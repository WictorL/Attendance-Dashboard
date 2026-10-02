// Same OAuth Client ID used in the HTML pages — verifyLogin checks that
// incoming tokens were actually issued for this app.
var GOOGLE_CLIENT_ID = '521793405590-qdr041egps4kkc41jda773qoi28sql8q.apps.googleusercontent.com';

function doGet(e) {
  var mode = e.parameter.mode;
  if (mode === 'verifyLogin') return handleVerifyLogin(e);
  if (mode === 'requestEmailCode') return handleRequestEmailCode(e);
  if (mode === 'verifyEmailCode') return handleVerifyEmailCode(e);
  if (mode === 'setPin') return handleSetPin(e);
  if (mode === 'verifyPin') return handleVerifyPin(e);
  if (mode === 'pinCheckIn') return handlePinCheckIn(e);
  if (mode === 'checkProximity') return handleCheckProximity(e);
  if (mode === 'data') return withCoachAuth(e, getDashboardData);
  if (mode === 'calendar') return withCoachAuth(e, getCalendarMonth);
  if (mode === 'members') return withCoachAuth(e, getMemberList);
  if (mode === 'profile') return getMemberProfile(e);
  return handleCheckIn(e); // default behaviour = check-in, same as your QR codes trigger
}

// ---------- Central lookups (both tabs live in this Router spreadsheet) ----------
//
// Gyms tab    — ONE row per gym:
//   Gym ID | Gym Name | Destination Sheet ID | Gym Latitude | Gym Longitude
// Routing tab — ONE row per QR code:
//   ID | Gym ID
//
// Both tabs are read by their HEADER names, so the column order doesn't matter —
// but the header text must match exactly.

function headerIndex(headerRow) {
  var col = {};
  headerRow.forEach(function(h, i) { col[String(h).trim()] = i; });
  return col;
}

// A gym's details (display name, which spreadsheet it uses, its location),
// looked up once and then reused for 5 minutes. Edits to the Gyms tab can
// therefore take up to 5 minutes to show up.
function getGymInfo(gymId) {
  var key = String(gymId || '').trim();
  if (!key) return null;

  var cacheKey = 'gyminfo_' + key;
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  var tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gyms');
  if (!tab) {
    Logger.log('No "Gyms" tab found in the Router spreadsheet.');
    return null;
  }
  var rows = tab.getDataRange().getValues();
  var col = headerIndex(rows[0]);

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][col['Gym ID']]).trim() === key) {
      var info = {
        gymName: rows[i][col['Gym Name']],
        destSheetId: String(rows[i][col['Destination Sheet ID']]).trim(),
        lat: rows[i][col['Gym Latitude']],
        lng: rows[i][col['Gym Longitude']]
      };
      cache.put(cacheKey, JSON.stringify(info), 300);
      return info;
    }
  }
  return null;
}

// Which gym a scanned QR code belongs to. The code → gym link is remembered for
// 10 minutes once found; codes that AREN'T found are never remembered, so a
// newly added code works straight away.
function findDestination(id) {
  var key = String(id || '').trim();
  if (!key) return null;

  var cache = CacheService.getScriptCache();
  var gymId = cache.get('route_' + key);

  if (!gymId) {
    var rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Routing').getDataRange().getValues();
    var col = headerIndex(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][col['ID']]).trim() === key) {
        gymId = String(rows[i][col['Gym ID']]).trim();
        break;
      }
    }
    if (!gymId) return null;
    cache.put('route_' + key, gymId, 600);
  }

  var info = getGymInfo(gymId);
  if (!info) return null;
  return {
    gym: info.gymName,
    gymId: gymId,
    sheetId: info.destSheetId,
    lat: info.lat,
    lng: info.lng
  };
}

// Straight-line distance between two lat/lng points, in metres (haversine formula)
function distanceMeters(lat1, lon1, lat2, lon2) {
  var R = 6371000;
  var toRad = function(deg) { return deg * Math.PI / 180; };
  var dLat = toRad(lat2 - lat1);
  var dLon = toRad(lon2 - lon1);
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// How far (in metres) a scan can be from a gym's registered location and still count.
// Kept generous on purpose — GPS accuracy indoors is often poor, and this is meant to catch
// "checking in from home," not to be a precise fence around the building.
var GEOFENCE_RADIUS_METERS = 150;

// The location rule for EVERY check-in. It "fails closed": if the gym's location
// isn't set up, or the phone didn't send a usable location, the check-in is
// refused — rather than quietly skipping the distance check, which is how a
// check-in from home used to get through.
// Returns null when the check-in is allowed, otherwise { reason, error }.
function geofenceProblem(gymLat, gymLng, rawLat, rawLng) {
  var gLat = parseFloat(gymLat), gLng = parseFloat(gymLng);
  if (isNaN(gLat) || isNaN(gLng)) {
    return { reason: 'not_setup', error: "Check-in isn't available yet, because this gym's location hasn't been set up." };
  }
  var lat = parseFloat(rawLat), lng = parseFloat(rawLng);
  if (isNaN(lat) || isNaN(lng)) {
    return { reason: 'no_location', error: 'We need your location to check you in. Please allow location access.' };
  }
  if (distanceMeters(lat, lng, gLat, gLng) > GEOFENCE_RADIUS_METERS) {
    return { reason: 'too_far', error: "You're too far from the gym to check in." };
  }
  return null;
}

// Has this member already checked in today? Only looks at the most recent 500
// attendance rows (new check-ins are always added at the bottom), so it stays
// quick however long the attendance history gets.
function alreadyCheckedInToday(attendanceSheet, memberId) {
  var lastRow = attendanceSheet.getLastRow();
  if (lastRow < 2) return false;
  var count = Math.min(500, lastRow - 1);
  var rows = attendanceSheet.getRange(lastRow - count + 1, 1, count, 3).getValues();
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var target = String(memberId).trim();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]).trim() !== target) continue;
    var ts = rows[i][2] instanceof Date ? rows[i][2] : new Date(rows[i][2]);
    if (!isNaN(ts) && Utilities.formatDate(ts, tz, 'yyyy-MM-dd') === today) return true;
  }
  return false;
}

// The profile page keeps a 60-second copy of each member's history; throw it
// away after a check-in so the new visit shows up straight away.
function clearProfileCache(memberId, gymId) {
  CacheService.getScriptCache().removeAll([
    'profile_' + memberId + '_' + gymId,
    'profile_' + memberId + '_self'
  ]);
}

function handleCheckIn(e) {
  var id = e.parameter.id;
  var dest = findDestination(id);
  var resultObj;

  if (!dest || !dest.sheetId) {
    resultObj = { success: false, error: 'Code not recognised or not yet assigned' };
  } else {
    var destSheet = SpreadsheetApp.openById(dest.sheetId);
    var members = destSheet.getSheetByName("Members").getDataRange().getValues();

    var name = null;
    for (var i = 1; i < members.length; i++) {
      if (members[i][0] === id) { name = members[i][1]; break; }
    }

    var problem = name ? geofenceProblem(dest.lat, dest.lng, e.parameter.lat, e.parameter.lng) : null;

    if (!name) {
      resultObj = { success: false, error: "ID not found in this gym's member list" };
    } else if (problem) {
      resultObj = { success: false, reason: problem.reason, error: problem.error };
    } else {
      destSheet.getSheetByName("Attendance").appendRow([id, name, new Date()]);
      clearProfileCache(id, dest.gymId);
      resultObj = { success: true, name: name };
    }
  }

  return respond(e, resultObj);
}

function getDashboardData(e) {
  var gymId = e.parameter.gymId;

  // The full calculation below re-reads and re-processes the ENTIRE attendance
  // history every time, which is the slowest part of the whole app. Caching the
  // result briefly means a refresh, or a second person loading the dashboard
  // moments later, gets an instant answer instead of paying that cost again.
  var cacheKey = 'dashdata_' + gymId;
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) return respond(e, JSON.parse(cached));

  var gymInfo = getGymInfo(gymId);
  var destId = gymInfo ? gymInfo.destSheetId : null;

  var resultObj;
  var gymDisplayName = gymInfo ? gymInfo.gymName : null;

  if (!destId) {
    resultObj = { error: 'Gym not found' };
  } else {
    var destSheet = SpreadsheetApp.openById(destId);
    var attendanceRows = destSheet.getSheetByName("Attendance").getDataRange().getValues();
    var membersRows = destSheet.getSheetByName("Members").getDataRange().getValues();

    var totalMembers = membersRows.length - 1;
    var now = new Date();
    var todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var weekStart = new Date(todayStart.getTime() - 6 * 24 * 60 * 60 * 1000);
    var riskCutoff = new Date(todayStart.getTime() - 10 * 24 * 60 * 60 * 1000);

    var totalCheckIns = 0, todayCount = 0, weekCount = 0;
    var lastName = null, lastTime = null;
    var weekdayTotals = [0, 0, 0, 0, 0, 0, 0];
    var calendarDays = {};

    var memberById = {};
    var beltStats = {};
    for (var m = 1; m < membersRows.length; m++) {
      var mid = membersRows[m][0], mname = membersRows[m][1], mbelt = membersRows[m][2] || 'Unspecified';
      memberById[mid] = { name: mname, belt: mbelt, lastVisit: null, totalVisits: 0, visitsThisWeek: 0 };
      if (!beltStats[mbelt]) beltStats[mbelt] = { memberCount: 0, visitCount: 0, members: [] };
      beltStats[mbelt].memberCount++;
      beltStats[mbelt].members.push(memberById[mid]);
    }

    var dedupedEvents = {};
    for (var j = 1; j < attendanceRows.length; j++) {
      var rawId = attendanceRows[j][0];
      var rawTs = new Date(attendanceRows[j][2]);
      if (isNaN(rawTs)) continue;
      var dedupDayKey = Utilities.formatDate(rawTs, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var dedupKey = rawId + '|' + dedupDayKey;
      if (!dedupedEvents[dedupKey] || rawTs > dedupedEvents[dedupKey].ts) {
        dedupedEvents[dedupKey] = { id: rawId, ts: rawTs, name: attendanceRows[j][1] };
      }
    }

    var todayList = [];

    for (var dedupKey in dedupedEvents) {
      var rid = dedupedEvents[dedupKey].id;
      var ts = dedupedEvents[dedupKey].ts;

      totalCheckIns++;
      if (ts >= todayStart) { todayCount++; todayList.push({ name: dedupedEvents[dedupKey].name, time: ts }); }
      if (ts >= weekStart) weekCount++;
      if (!lastTime || ts > lastTime) { lastTime = ts; lastName = dedupedEvents[dedupKey].name; }

      var dow = (ts.getDay() + 6) % 7;
      weekdayTotals[dow]++;
      if (ts.getFullYear() === now.getFullYear() && ts.getMonth() === now.getMonth()) {
        var dayKey = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        if (!calendarDays[dayKey]) calendarDays[dayKey] = [];
        calendarDays[dayKey].push(dedupedEvents[dedupKey].name);
      }

      if (memberById[rid]) {
        if (!memberById[rid].lastVisit || ts > memberById[rid].lastVisit) memberById[rid].lastVisit = ts;
        memberById[rid].totalVisits++;
        if (ts >= weekStart) memberById[rid].visitsThisWeek++;
        var b = memberById[rid].belt;
        if (beltStats[b]) beltStats[b].visitCount++;
      }
    }

    todayList.sort(function(a, b) { return b.time - a.time; });

    var allMembersOutput = [];
    for (var mid2 in memberById) {
      var mm = memberById[mid2];
      allMembersOutput.push({ name: mm.name, belt: mm.belt, totalVisits: mm.totalVisits, visitsThisWeek: mm.visitsThisWeek });
    }

    var atRisk = [];
    for (var id in memberById) {
      var mem = memberById[id];
      if (mem.lastVisit && mem.lastVisit < riskCutoff) {
        var daysSince = Math.floor((todayStart - mem.lastVisit) / (24 * 60 * 60 * 1000));
        atRisk.push({ name: mem.name, daysSince: daysSince });
      }
    }
    atRisk.sort(function(a, b) { return b.daysSince - a.daysSince; });

    var beltBreakdown = [];
    for (var belt in beltStats) {
      var s = beltStats[belt];
      var roster = s.members.map(function(mem) {
        var daysSince = mem.lastVisit ? Math.floor((todayStart - mem.lastVisit) / (24 * 60 * 60 * 1000)) : null;
        var label;
        if (!mem.lastVisit) label = 'Never checked in';
        else if (daysSince <= 0) label = 'Checked in today';
        else if (daysSince === 1) label = '1 day ago';
        else label = daysSince + ' days ago';
        return { name: mem.name, daysSince: daysSince, label: label };
      });
      roster.sort(function(a, b) {
        if (a.daysSince === null) return 1;
        if (b.daysSince === null) return -1;
        return a.daysSince - b.daysSince;
      });
      beltBreakdown.push({
        belt: belt,
        avgVisits: s.memberCount ? Math.round((s.visitCount / s.memberCount) * 10) / 10 : 0,
        memberCount: s.memberCount,
        members: roster
      });
    }
    beltBreakdown.sort(function(a, b) { return b.avgVisits - a.avgVisits; });

    resultObj = {
      gymName: gymDisplayName,
      totalMembers: totalMembers, totalCheckIns: totalCheckIns,
      today: todayCount, week: weekCount, lastName: lastName,
      lastTime: lastTime ? lastTime.toISOString() : null,
      atRisk: atRisk.slice(0, 5),
      beltBreakdown: beltBreakdown,
      weekdayTotals: weekdayTotals,
      calendarMonth: { year: now.getFullYear(), month: now.getMonth(), days: calendarDays },
      todayList: todayList,
      allMembers: allMembersOutput
    };
  }

  cache.put(cacheKey, JSON.stringify(resultObj), 60); // cached for 60 seconds
  return respond(e, resultObj);
}

// Returns per-day check-in counts for one specific month/year, so paging the calendar
// forward/back doesn't need to re-fetch the whole dashboard payload.
function getCalendarMonth(e) {
  var gymId = e.parameter.gymId;
  var year = parseInt(e.parameter.year, 10);
  var month = parseInt(e.parameter.month, 10);
  var gymInfo = getGymInfo(gymId);
  var destId = gymInfo ? gymInfo.destSheetId : null;

  var resultObj;
  if (!destId) {
    resultObj = { error: 'Gym not found' };
  } else {
    var destSheet = SpreadsheetApp.openById(destId);
    var attendanceRows = destSheet.getSheetByName("Attendance").getDataRange().getValues();

    var dedupedEvents = {};
    for (var j = 1; j < attendanceRows.length; j++) {
      var rawId = attendanceRows[j][0];
      var rawTs = new Date(attendanceRows[j][2]);
      if (isNaN(rawTs)) continue;
      if (rawTs.getFullYear() !== year || rawTs.getMonth() !== month) continue;
      var dayKey = Utilities.formatDate(rawTs, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var dedupKey = rawId + '|' + dayKey;
      if (!dedupedEvents[dedupKey] || rawTs > dedupedEvents[dedupKey].ts) {
        dedupedEvents[dedupKey] = { ts: rawTs, name: attendanceRows[j][1], dayKey: dayKey };
      }
    }

    var days = {};
    for (var k in dedupedEvents) {
      var ev = dedupedEvents[k];
      if (!days[ev.dayKey]) days[ev.dayKey] = [];
      days[ev.dayKey].push(ev.name);
    }
    resultObj = { year: year, month: month, days: days };
  }

  return respond(e, resultObj);
}

// Returns every member's id/name/belt for this gym, so the dashboard can offer
// a searchable list — selecting one then reuses getMemberProfile for their history.
function getMemberList(e) {
  var gymId = e.parameter.gymId;
  var gymInfo = getGymInfo(gymId);
  var destId = gymInfo ? gymInfo.destSheetId : null;

  var resultObj;
  if (!destId) {
    resultObj = { error: 'Gym not found' };
  } else {
    var destSheet = SpreadsheetApp.openById(destId);
    var membersRows = destSheet.getSheetByName("Members").getDataRange().getValues();
    var members = [];
    for (var m = 1; m < membersRows.length; m++) {
      members.push({ id: membersRows[m][0], name: membersRows[m][1], belt: membersRows[m][2] || 'Unspecified' });
    }
    members.sort(function(a, b) { return String(a.name).localeCompare(String(b.name)); });
    resultObj = { members: members };
  }

  return respond(e, resultObj);
}

function getMemberProfile(e) {
  var id = e.parameter.id;
  var gymId = e.parameter.gymId;

  // Same idea as the dashboard cache — a member reopening their profile, or the
  // coach looking up the same person twice, shouldn't re-scan the whole
  // attendance sheet each time within this short window.
  var cacheKey = 'profile_' + id + '_' + (gymId || 'self');
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) return respond(e, JSON.parse(cached));

  var sheetId = null;
  if (gymId) {
    var gymInfo = getGymInfo(gymId);
    sheetId = gymInfo ? gymInfo.destSheetId : null;
  } else {
    var dest = findDestination(id);
    if (dest) sheetId = dest.sheetId;
  }

  var resultObj;
  if (!sheetId) {
    resultObj = { error: 'not found' };
  } else {
    var destSheet = SpreadsheetApp.openById(sheetId);
    var membersRows = destSheet.getSheetByName("Members").getDataRange().getValues();
    var attendanceRows = destSheet.getSheetByName("Attendance").getDataRange().getValues();

    var member = null;
    for (var m = 1; m < membersRows.length; m++) {
      if (membersRows[m][0] === id) {
        member = { id: membersRows[m][0], name: membersRows[m][1], belt: membersRows[m][2] };
        break;
      }
    }

    if (!member) {
      resultObj = { error: 'not found' };
    } else {
      var rawVisits = attendanceRows.slice(1)
        .filter(function(r) { return r[0] === id; })
        .map(function(r) { return new Date(r[2]); })
        .filter(function(d) { return !isNaN(d); });

      var visitsByDay = {};
      rawVisits.forEach(function(d) {
        var dayKey = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        if (!visitsByDay[dayKey] || d > visitsByDay[dayKey]) visitsByDay[dayKey] = d;
      });
      var visits = Object.keys(visitsByDay).map(function(k) { return visitsByDay[k]; });

      resultObj = { member: member, visits: visits };
    }
  }

  cache.put(cacheKey, JSON.stringify(resultObj), 60); // cached for 60 seconds
  return respond(e, resultObj);
}

// Verifies a Google sign-in token and looks the signed-in email up in the
// Authorized Users tab, returning that person's role + which gym they belong to.
function handleVerifyLogin(e) {
  var idToken = e.parameter.token;
  if (!idToken) {
    return respond(e, { success: false, message: 'Missing sign-in token.' });
  }

  var email = verifyGoogleToken(idToken);
  if (!email) {
    return respond(e, { success: false, message: 'Could not verify sign-in. Please try again.' });
  }

  return completeSignIn(e, email);
}

// The shared last step for BOTH sign-in methods (Google, or an emailed code).
// Once someone has proven they own an email address, everything after that is
// identical — so it lives in one place rather than two copies drifting apart.
function completeSignIn(e, email) {
  var user = findAuthorizedUser(email);
  if (!user || user.status !== 'Active') {
    return respond(e, { success: false, message: 'No profile found for this email. Ask your coach to add you.' });
  }

  var profile = {
    email: email,
    name: user.name,
    role: user.role,
    gymId: user.gymId,
    gymName: findGymName(user.gymId)
  };

  // Members also need their own member ID, and whether they've already set a PIN —
  // the sign-in page uses hasPin to decide whether to prompt them to create one.
  if (user.role === 'Member') {
    profile.memberId = user.memberId;
    profile.hasPin = memberHasPin(user.gymId, user.memberId);
  } else if (user.role === 'Coach') {
    profile.hasPin = coachHasPin(email);
  }

  // viaSignIn marks this session as coming from a FULL sign-in (not just a PIN
  // unlock) — only sessions like this are allowed to set or reset a PIN.
  var sessionKey = createSession(Object.assign({ viaSignIn: true }, profile));
  return respond(e, { success: true, profile: profile, sessionKey: sessionKey });
}

// ---------- Email sign-in codes (for people without a Google account) ----------

var OTP_TTL_SECONDS = 600;   // a code lasts 10 minutes
var OTP_MAX_ATTEMPTS = 5;    // wrong guesses allowed before the code is destroyed
var OTP_MAX_PER_HOUR = 5;    // codes one email address can request per hour

// Step 1: someone types their email and asks for a code. The reply is the same
// whether or not the address is registered, so this can't be used to discover
// which emails belong to a gym.
function handleRequestEmailCode(e) {
  var email = String(e.parameter.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return respond(e, { success: false, message: 'Please enter a valid email address.' });
  }

  var cache = CacheService.getScriptCache();
  var key = sha256Hex('otp:' + email);

  if (cache.get('otpwait_' + key)) {
    return respond(e, { success: false, message: 'A code was just sent. Please wait a minute before asking for another.' });
  }
  var sentThisHour = Number(cache.get('otpcount_' + key) || 0);
  if (sentThisHour >= OTP_MAX_PER_HOUR) {
    return respond(e, { success: false, message: 'Too many codes requested. Please try again in an hour.' });
  }
  // Free Gmail accounts can send ~100 emails a day; fail politely if that runs out.
  if (MailApp.getRemainingDailyQuota() < 1) {
    return respond(e, { success: false, message: 'Email sign-in is busy right now. Please use Google, or try again tomorrow.' });
  }

  cache.put('otpwait_' + key, '1', 60);
  cache.put('otpcount_' + key, String(sentThisHour + 1), 3600);

  var genericReply = { success: true };
  var user = findAuthorizedUser(email);
  if (!user || user.status !== 'Active') return respond(e, genericReply);

  var code = generateSixDigitCode();
  cache.put('otp_' + key, JSON.stringify({
    hash: sha256Hex(code + ':' + email), // only a scrambled copy is kept, never the code itself
    attempts: 0,
    expiresAt: Date.now() + OTP_TTL_SECONDS * 1000
  }), OTP_TTL_SECONDS);

  try {
    sendSignInCodeEmail(email, code, findGymName(user.gymId));
  } catch (err) {
    cache.remove('otp_' + key);
    cache.remove('otpwait_' + key);
    return respond(e, { success: false, message: 'Could not send the email just now. Please try again in a moment.' });
  }
  return respond(e, genericReply);
}

// Step 2: they type the code back. Each code works once, expires after 10 minutes,
// and is destroyed after 5 wrong guesses.
function handleVerifyEmailCode(e) {
  var email = String(e.parameter.email || '').trim().toLowerCase();
  var code = String(e.parameter.code || '').replace(/\D/g, '');
  if (code.length !== 6) {
    return respond(e, { success: false, message: 'Enter the 6-digit code from the email.' });
  }

  var cache = CacheService.getScriptCache();
  var cacheKey = 'otp_' + sha256Hex('otp:' + email);
  var stored = cache.get(cacheKey);
  if (!stored) {
    return respond(e, { success: false, expired: true, message: 'This code has expired or was already used. Tap "Send a new code".' });
  }

  var entry = JSON.parse(stored);
  if (entry.hash !== sha256Hex(code + ':' + email)) {
    entry.attempts++;
    var secondsLeft = Math.floor((entry.expiresAt - Date.now()) / 1000);
    if (entry.attempts >= OTP_MAX_ATTEMPTS || secondsLeft < 1) {
      cache.remove(cacheKey);
      return respond(e, { success: false, expired: true, message: 'Too many incorrect tries. Please request a new code.' });
    }
    cache.put(cacheKey, JSON.stringify(entry), secondsLeft);
    var left = OTP_MAX_ATTEMPTS - entry.attempts;
    return respond(e, { success: false, message: "That code isn't right. " + left + (left === 1 ? ' try' : ' tries') + ' left.' });
  }

  cache.remove(cacheKey); // single use
  return completeSignIn(e, email);
}

function generateSixDigitCode() {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + ':' + Date.now());
  var n = 0;
  for (var i = 0; i < 4; i++) n = n * 256 + (bytes[i] < 0 ? bytes[i] + 256 : bytes[i]);
  return String(n % 1000000).padStart(6, '0');
}

function sendSignInCodeEmail(email, code, gymName) {
  var plain = 'Your sign-in code for ' + gymName + ' is ' + code + '.\n\n' +
    'It expires in 10 minutes. If you didn\'t ask for this, you can ignore this email.';
  var html =
    '<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:420px;margin:0 auto;padding:24px;color:#222;">' +
      '<p style="font-size:15px;margin:0 0 16px;">Your sign-in code for <strong>' + escapeHtml(gymName) + '</strong>:</p>' +
      '<div style="font-size:34px;font-weight:700;letter-spacing:8px;text-align:center;padding:18px;' +
        'background:#FFF1E8;border:2px solid #E8590C;border-radius:12px;color:#161616;">' + code + '</div>' +
      '<p style="font-size:13px;color:#777;margin:16px 0 0;">It expires in 10 minutes. ' +
        'If you didn\'t ask for this, you can safely ignore this email.</p>' +
    '</div>';
  MailApp.sendEmail({
    to: email,
    subject: code + ' is your OssTrack sign-in code',
    body: plain,
    htmlBody: html,
    name: 'OssTrack'
  });
}

function escapeHtml(text) {
  return String(text).replace(/[&<>"']/g, function(c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function sha256Hex(text) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return raw.map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

// RUN THIS ONCE from the editor after pasting this file in. It grants the script
// permission to send email (the same kind of one-time "Allow" step as before),
// and logs how many emails your account can still send today.
function authorizeEmail() {
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());
}

// Confirms a token was really issued by Google for THIS app, and returns the
// signed-in email if so (or null if invalid, expired, or for a different app).
// Checking with Google is the slowest part of every request, so a valid result is
// cached briefly — only the first request in a few minutes pays that cost.
function verifyGoogleToken(idToken) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'tok_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, idToken)
  );

  var cached = cache.get(cacheKey);
  if (cached) return cached === '__invalid__' ? null : cached;

  var url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    cache.put(cacheKey, '__invalid__', 60);
    return null;
  }

  var tokenData = JSON.parse(response.getContentText());
  if (tokenData.aud !== GOOGLE_CLIENT_ID) {
    cache.put(cacheKey, '__invalid__', 60);
    return null;
  }

  cache.put(cacheKey, tokenData.email, 300);
  return tokenData.email;
}

// Looks up one row in the Authorized Users tab by email, by header name (not
// column position) so reordering columns later won't silently break this.
// Expected columns: Email | GymID | Role | Name | Status | MemberID
function findAuthorizedUser(email) {
  var router = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = router.getSheetByName('Authorized Users');
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];

  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][col['Email']]).toLowerCase() === email.toLowerCase()) {
      return {
        gymId: rows[i][col['GymID']],
        role: rows[i][col['Role']],
        name: rows[i][col['Name']],
        status: rows[i][col['Status']],
        memberId: col['MemberID'] !== undefined ? rows[i][col['MemberID']] : null
      };
    }
  }
  return null;
}

// Looks up a gym's display name from the Gyms tab. Falls back to the ID
// itself if not found — a missing name should never block someone signing in.
function findGymName(gymId) {
  var info = getGymInfo(gymId);
  return info ? info.gymName : gymId;
}

// Wraps a result as JSON or JSONP depending on whether a callback was requested.
function respond(e, resultObj) {
  var json = JSON.stringify(resultObj);
  if (e.parameter.callback) {
    return ContentService.createTextOutput(e.parameter.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  var output = ContentService.createTextOutput(json);
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// Verifies a token belongs to an Active user whose role is in allowedRoles.
// Returns that user's record (gymId, role, name) if so, or null otherwise —
// callers must treat null as "reject this request," never fall through.
function requireRole(e, allowedRoles) {
  var idToken = e.parameter.token;
  if (!idToken) return null;

  var email = verifyGoogleToken(idToken);
  if (!email) return null;

  var user = findAuthorizedUser(email);
  if (!user || user.status !== 'Active') return null;
  if (allowedRoles.indexOf(user.role) === -1) return null;

  user.email = email; // attach it — several callers need it (e.g. setting a coach's PIN)
  return user;
}

// Creates a short-lived session key after any successful sign-in (Google or PIN),
// so later requests can prove "this device unlocked a moment ago" without needing
// a fresh Google token each time — which PIN-unlocked devices never have.
function createSession(profile) {
  var key = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  cache.put('session_' + key, JSON.stringify(profile), 21600); // 6 hours — the max CacheService allows
  return key;
}

// Looks up a session key created by createSession. Returns the profile it was
// issued for for if valid and the role matches, otherwise null.
function requireSession(e, allowedRoles) {
  var key = e.parameter.sessionKey;
  if (!key) return null;
  var cache = CacheService.getScriptCache();
  var cached = cache.get('session_' + key);
  if (!cached) return null;
  var profile = JSON.parse(cached);
  if (allowedRoles.indexOf(profile.role) === -1) return null;
  return profile;
}

// Finds a row in Authorized Users by email, returning its row number and column
// map together — shared by every coach-PIN function below so the sheet is only
// scanned once per helper, the same way everywhere.
function findAuthorizedUserRowIndex(sheet, email) {
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][col['Email']]).toLowerCase() === String(email).toLowerCase()) {
      return { rowNum: i + 1, col: col, row: rows[i] };
    }
  }
  return null;
}

function coachHasPin(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Authorized Users');
  var found = findAuthorizedUserRowIndex(sheet, email);
  if (!found || found.col['PIN Hash'] === undefined) return false;
  return !!found.row[found.col['PIN Hash']];
}

function setCoachPin(email, pin) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Authorized Users');
  var found = findAuthorizedUserRowIndex(sheet, email);
  if (!found || found.col['PIN Hash'] === undefined) return false;
  sheet.getRange(found.rowNum, found.col['PIN Hash'] + 1).setValue(hashPin(pin, email));
  return true;
}

// A coach's PIN is only ever checked alongside their known email (never used to
// anonymously identify someone the way a member's check-in PIN is), so there's
// no need to enforce uniqueness the way handleSetPin does for members.
function verifyCoachPin(email, pin) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Authorized Users');
  var found = findAuthorizedUserRowIndex(sheet, email);
  if (!found) return null;
  if (found.row[found.col['Status']] !== 'Active' || found.row[found.col['Role']] !== 'Coach') return null;
  var pinCol = found.col['PIN Hash'];
  var storedHash = pinCol !== undefined ? String(found.row[pinCol]).trim() : '';
  if (!storedHash || storedHash !== hashPin(pin, email)) return null;
  return {
    role: 'Coach',
    email: email,
    gymId: found.row[found.col['GymID']],
    gymName: findGymName(found.row[found.col['GymID']]),
    name: found.row[found.col['Name']]
  };
}

function withCoachAuth(e, handlerFn) {
  var user = requireSession(e, ['Coach']);
  if (!user) {
    return respond(e, { error: 'Not signed in, or your session has expired. Please unlock the app again.' });
  }
  e.parameter.gymId = user.gymId;
  return handlerFn(e);
}

// Finds a column's index by its header text, so column order can change later
// without silently breaking anything that reads it.
// Turns a PIN into an irreversible scrambled string before it's ever stored, so
// opening the sheet never shows anyone's actual PIN — only this hash. The salt
// (the member's own ID, or a coach's email) means two people choosing the same
// PIN still produce different-looking hashes.
function hashPin(pin, salt) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pin + ':' + salt);
  return raw.map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function findColumnIndex(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(headerName);
}

// Opens a gym's own Members sheet, given its GymID. Null if the gym isn't found.
function getMembersSheet(gymId) {
  var info = getGymInfo(gymId);
  if (!info) return null;
  return SpreadsheetApp.openById(info.destSheetId).getSheetByName('Members');
}

// Finds a member's row number in their gym's Members sheet, by ID (column A).
// Returns -1 if not found.
function findMemberRow(membersSheet, memberId) {
  var ids = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(memberId)) return i + 2;
  }
  return -1;
}

// Whether a member already has a PIN set — used right after sign-in to decide
// whether to prompt them to create one.
function memberHasPin(gymId, memberId) {
  var sheet = getMembersSheet(gymId);
  if (!sheet) return false;
  var pinCol = findColumnIndex(sheet, 'PIN Hash');
  if (pinCol === -1) return false;
  var row = findMemberRow(sheet, memberId);
  if (row === -1) return false;
  return !!sheet.getRange(row, pinCol + 1).getValue();
}

// Lets a signed-in member choose their own PIN. Requires a valid, fresh Google
// token — this is only ever called right after a real sign-in, never on its own.
// Setting a PIN needs proof of a FULL sign-in in the last few hours — a session
// that was only unlocked with a PIN can't be used to change the PIN.
function requireFreshSignIn(e) {
  if (e.parameter.sessionKey) {
    var profile = requireSession(e, ['Member', 'Coach']);
    return (profile && profile.viaSignIn) ? profile : null;
  }
  return requireRole(e, ['Member', 'Coach']); // older cached pages still send a Google token
}

function handleSetPin(e) {
  var user = requireFreshSignIn(e);
  if (!user) return respond(e, { success: false, message: 'Not signed in.' });

  var pin = String(e.parameter.pin || '').trim();
  if (!/^[0-9]{4}$/.test(pin)) {
    return respond(e, { success: false, message: 'PIN must be exactly 4 digits.' });
  }

  if (user.role === 'Coach') {
    var ok = setCoachPin(user.email, pin);
    if (!ok) return respond(e, { success: false, message: 'This account is missing a PIN column. Please contact support.' });
    return respond(e, { success: true });
  }

  // Member path — PINs must stay unique WITHIN a gym, since a member's PIN alone
  // (with no email or sign-in) is what identifies them during a PIN check-in.
  var membersSheet = getMembersSheet(user.gymId);
  if (!membersSheet) return respond(e, { success: false, message: 'Gym not found.' });

  var pinCol = findColumnIndex(membersSheet, 'PIN Hash');
  if (pinCol === -1) return respond(e, { success: false, message: 'This gym is not set up for PIN check-in yet.' });

  var row = findMemberRow(membersSheet, user.memberId);
  if (row === -1) return respond(e, { success: false, message: 'Member not found.' });

  // Each row has its own salt (its own member ID), so checking for a duplicate
  // means re-hashing the candidate PIN with EACH row's salt and comparing —
  // a plain string comparison won't work once salts differ per row.
  var allRows = membersSheet.getDataRange().getValues();
  for (var i = 1; i < allRows.length; i++) {
    var otherId = allRows[i][0];
    var otherHash = String(allRows[i][pinCol] || '').trim();
    if (otherId !== user.memberId && otherHash && otherHash === hashPin(pin, otherId)) {
      return respond(e, { success: false, message: 'That PIN is already in use. Please choose a different one.' });
    }
  }

  membersSheet.getRange(row, pinCol + 1).setValue(hashPin(pin, user.memberId));
  return respond(e, { success: true });
}

// Unlocks the app on a device that already remembers this member — deliberately
// does NOT require a fresh Google token (that's the whole point of remembering
// the device), and is never geofenced, since opening your profile isn't a check-in.
function handleVerifyPin(e) {
  var pin = String(e.parameter.pin || '').trim();

  if (e.parameter.role === 'Coach') {
    var coachProfile = verifyCoachPin(e.parameter.email, pin);
    if (!coachProfile) {
      return respond(e, { success: false, message: 'Incorrect PIN.' });
    }
    coachProfile.hasPin = true;
    var coachSessionKey = createSession(coachProfile);
    return respond(e, { success: true, profile: coachProfile, sessionKey: coachSessionKey });
  }

  var gymId = e.parameter.gymId;
  var memberId = e.parameter.memberId;

  var membersSheet = getMembersSheet(gymId);
  if (!membersSheet) return respond(e, { success: false, message: 'Gym not found.' });

  var pinCol = findColumnIndex(membersSheet, 'PIN Hash');
  var row = findMemberRow(membersSheet, memberId);
  if (pinCol === -1 || row === -1) {
    return respond(e, { success: false, message: 'Could not verify. Please sign in again.' });
  }

  var storedHash = String(membersSheet.getRange(row, pinCol + 1).getValue()).trim();
  if (!storedHash || storedHash !== hashPin(pin, memberId)) {
    return respond(e, { success: false, message: 'Incorrect PIN.' });
  }

  var memberProfile = {
    role: 'Member',
    gymId: gymId,
    gymName: findGymName(gymId),
    memberId: memberId,
    name: membersSheet.getRange(row, 2).getValue(),
    hasPin: true
  };
  var memberSessionKey = createSession(memberProfile);
  return respond(e, { success: true, profile: memberProfile, sessionKey: memberSessionKey });
}

// Checks a member in by PIN instead of a QR scan. Same geofence rule as scanning —
// resolved the same way handleCheckIn does, via the Gyms tab's stored coordinates.
function handlePinCheckIn(e) {
  var gymId = e.parameter.gymId;
  var memberId = e.parameter.memberId;
  var pin = String(e.parameter.pin || '').trim();

  var membersSheet = getMembersSheet(gymId);
  if (!membersSheet) return respond(e, { success: false, error: 'Gym not found.' });

  var pinCol = findColumnIndex(membersSheet, 'PIN Hash');
  var row = findMemberRow(membersSheet, memberId);
  if (pinCol === -1 || row === -1) {
    return respond(e, { success: false, error: 'Member not found.' });
  }

  var storedHash = String(membersSheet.getRange(row, pinCol + 1).getValue()).trim();
  if (!storedHash || storedHash !== hashPin(pin, memberId)) {
    return respond(e, { success: false, reason: 'pin', error: 'Incorrect PIN.' });
  }

  var gymInfo = getGymInfo(gymId);
  if (!gymInfo) return respond(e, { success: false, error: 'Gym not found.' });

  var problem = geofenceProblem(gymInfo.lat, gymInfo.lng, e.parameter.lat, e.parameter.lng);
  if (problem) return respond(e, { success: false, reason: problem.reason, error: problem.error });

  var attendance = SpreadsheetApp.openById(gymInfo.destSheetId).getSheetByName('Attendance');
  if (alreadyCheckedInToday(attendance, memberId)) {
    return respond(e, { success: false, reason: 'already', error: "You're already checked in today." });
  }

  var name = membersSheet.getRange(row, 2).getValue();
  attendance.appendRow([memberId, name, new Date()]);
  clearProfileCache(memberId, gymId);
  return respond(e, { success: true, name: name });
}

// Read-only proximity check — tells the member profile page whether the device's
// current location is within the gym's geofence, WITHOUT logging a check-in or
// needing a PIN. Used to show a "you're at the gym — check in?" prompt.
function handleCheckProximity(e) {
  var gymId = e.parameter.gymId;
  var lat = parseFloat(e.parameter.lat);
  var lng = parseFloat(e.parameter.lng);

  var gymInfo = getGymInfo(gymId);
  if (!gymInfo || isNaN(parseFloat(gymInfo.lat)) || isNaN(parseFloat(gymInfo.lng))) {
    return respond(e, { withinRange: null, notSetUp: true });
  }
  if (isNaN(lat) || isNaN(lng)) return respond(e, { withinRange: null });

  var distance = distanceMeters(lat, lng, parseFloat(gymInfo.lat), parseFloat(gymInfo.lng));
  return respond(e, { withinRange: distance <= GEOFENCE_RADIUS_METERS });
}
