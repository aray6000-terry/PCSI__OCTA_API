/**
 * 線上打卡 (第一階段)：打卡、外勤拍照、打卡地點 (公司 / 工地)、每日出勤判定
 * -------------------------------------------------------------
 * 安裝：在 Apps Script 新增指令碼檔案 Attendance，貼上本檔即可 (與 Code.gs 同一專案)。
 * 第一次有人打卡時，會自動建立 3 張工作表：attendance_logs、attendance_daily、work_locations，
 * 並以公司地址自動建立「公司」打卡地點 (半徑 200 公尺)。
 *
 * 規則 (2026-11-19 決定)：
 *  - 彈性上班 08:30～09:00 到班；應下班 = 到班 + 9.5 小時 (含午休 12:00～13:30)，即 18:00～18:30
 *  - 08:30 前到班以 08:30 計；09:00 後到班算遲到 (不另給寬限)
 *  - 在登錄地點範圍外打卡 = 外勤打卡：必須選 / 填工地並拍現場照片
 *  - 最高主管 (老闆) 免打卡
 *  - 原始打卡紀錄只新增、不修改 (勞基法第 30 條：逐日記載至分鐘、保存 5 年)
 */

const ATT = {
  SHEETS: { LOGS: "attendance_logs", DAILY: "attendance_daily", LOCATIONS: "work_locations" },
  COMPANY_ADDRESS: "桃園市中壢區民權路二段147號",
  DEFAULT_RADIUS_M: 200,
  ARRIVE_FROM: 8 * 60 + 30,     // 08:30
  ARRIVE_UNTIL: 9 * 60,         // 09:00
  DAY_SPAN_MIN: 9 * 60 + 30,    // 9.5 小時 (含午休)
  LUNCH_START: 12 * 60,
  LUNCH_END: 13 * 60 + 30,
  STAY_AFTER_ALERT_MIN: 30,     // 超過應下班 30 分鐘才標示
  OFFSITE_MISMATCH_M: 1000,     // 外勤打卡離所選工地超過 1 公里 → 地點不符
  SAME_DEVICE_WINDOW_MIN: 10,   // 同一裝置 10 分鐘內替不同人打卡 → 標示
  DUPLICATE_WINDOW_SEC: 60,     // 1 分鐘內重複按同一種卡 → 視為同一筆
  PHOTO_FOLDER_PROP: "ATT_PHOTO_FOLDER_ID",
  PHOTO_MAX_BYTES: 3 * 1024 * 1024
};

const ATT_HEADERS = {
  LOGS: ["id", "user_id", "punched_at", "date", "type", "location_id", "location_name", "distance_m", "in_range",
         "lat", "lng", "accuracy_m", "site_name", "note", "photo_url", "device_id", "flags", "created_at"],
  DAILY: ["date", "user_id", "status", "first_in", "last_out", "expected_in", "expected_out",
          "late_min", "early_min", "stay_after_min", "offsite", "leave_ids", "flags", "judged_at"],
  LOCATIONS: ["id", "name", "type", "address", "lat", "lng", "radius_m", "start_date", "end_date", "created_by", "created_at"]
};

// ======================== 工作表 ========================

function ensureAttendanceSheets_(ss) {
  ss = ss || getSpreadsheet();
  Object.keys(ATT.SHEETS).forEach(key => {
    const name = ATT.SHEETS[key];
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, ATT_HEADERS[key].length).setValues([ATT_HEADERS[key]]).setFontWeight("bold");
      sh.setFrozenRows(1);
      // 一律存成文字，避免試算表把日期時間自動轉成其他格式
      sh.getRange(1, 1, sh.getMaxRows(), ATT_HEADERS[key].length).setNumberFormat("@");
    }
  });
  const locSheet = ss.getSheetByName(ATT.SHEETS.LOCATIONS);
  if (locSheet.getLastRow() <= 1) seedCompanyLocation_(locSheet);
  return ss;
}

function seedCompanyLocation_(locSheet) {
  let lat = "", lng = "";
  try {
    const geo = geocodeAddress_(ATT.COMPANY_ADDRESS);
    if (geo) { lat = geo.lat; lng = geo.lng; }
  } catch (e) {
    console.error("公司地址定位失敗：" + e);
  }
  locSheet.appendRow(["LOC-COMPANY", "百成系統整合（民權路二段147號）", "公司", ATT.COMPANY_ADDRESS, lat, lng,
    ATT.DEFAULT_RADIUS_M, "", "", "SYSTEM", nowStr_()]);
}

function geocodeAddress_(address) {
  const res = Maps.newGeocoder().setRegion("tw").setLanguage("zh-TW").geocode(address);
  if (res && res.status === "OK" && res.results && res.results.length) {
    const loc = res.results[0].geometry.location;
    return { lat: Math.round(loc.lat * 1e6) / 1e6, lng: Math.round(loc.lng * 1e6) / 1e6, formatted: res.results[0].formatted_address };
  }
  return null;
}

// ======================== 工具 ========================

function nowStr_() { return Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss"); }
function todayStr_() { return Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd"); }

// "yyyy-MM-dd HH:mm(:ss)" → 當天第幾分鐘 (不經過 Date 解析，避免時區誤差)
function minOfDay_(s) {
  const m = String(s || "").match(/(\d{1,2}):(\d{2})/);
  return m ? parseInt(m[1], 10) * 60 + parseInt(m[2], 10) : null;
}
function hhmm_(min) {
  if (min === null || min === undefined || min === "") return "";
  const h = Math.floor(min / 60), m = min % 60;
  return ("0" + h).slice(-2) + ":" + ("0" + m).slice(-2);
}

function distanceM_(lat1, lng1, lat2, lng2) {
  const R = 6371000, toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1), dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return Math.round(2 * R * Math.asin(Math.sqrt(a)));
}

function isLocationActive_(loc, dateStr) {
  const s = String(loc.start_date || "").slice(0, 10), e = String(loc.end_date || "").slice(0, 10);
  return (!s || s <= dateStr) && (!e || e >= dateStr) && loc.lat !== "" && loc.lng !== "";
}

function canManageLocations_(user) {
  return !!user && (user.role === "Manager" || isAdminOrHR_(user));
}

// 主管可看直屬部屬；人資、管理者、超級使用者可看全部
function canViewAttendanceOf_(me, target) {
  if (!me || !target) return false;
  if (me.id === target.id || isAdminOrHR_(me)) return true;
  return me.role === "Manager" && String(target.manager_id || "") === me.id;
}

// ======================== 打卡 ========================

/**
 * params: { type: "IN"|"OUT", lat, lng, accuracy, siteName, note, photo (data URL), deviceId }
 * 時間一律以伺服器時間為準；打卡人一律為登入者本人 (params.userId 由 token 決定)
 */
function punch(params) {
  const me = params._me;
  const type = params.type === "OUT" ? "OUT" : (params.type === "IN" ? "IN" : null);
  if (!type) return { success: false, message: "請選擇上班或下班打卡。" };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: "系統忙碌中，請再按一次。" };
  try {
    const ss = ensureAttendanceSheets_();
    const logSheet = ss.getSheetByName(ATT.SHEETS.LOGS);
    const now = nowStr_(), today = now.slice(0, 10);
    const logs = sheetToObjects(logSheet);

    // 1 分鐘內重複按同一種卡：回傳原本那筆，不重複記錄
    const nowMin = minOfDay_(now);
    const dup = logs.find(l => l.user_id === me.id && l.date === today && l.type === type &&
      Math.abs(minOfDay_(l.punched_at) - nowMin) * 60 <= ATT.DUPLICATE_WINDOW_SEC);
    if (dup) return { success: true, duplicate: true, message: "剛剛已經打過卡了。", record: dup, today: todaySummary_(ss, me) };

    // 定位比對
    const lat = parseFloat(params.lat), lng = parseFloat(params.lng);
    const hasGps = isFinite(lat) && isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180;
    const accuracy = isFinite(parseFloat(params.accuracy)) ? Math.round(parseFloat(params.accuracy)) : "";
    const locations = sheetToObjects(ss.getSheetByName(ATT.SHEETS.LOCATIONS)).filter(l => isLocationActive_(l, today));
    let nearest = null;
    if (hasGps) {
      locations.forEach(l => {
        const d = distanceM_(lat, lng, parseFloat(l.lat), parseFloat(l.lng));
        if (!nearest || d < nearest.distance) nearest = { loc: l, distance: d };
      });
    }
    const inRange = !!(nearest && nearest.distance <= (parseFloat(nearest.loc.radius_m) || ATT.DEFAULT_RADIUS_M));

    const flags = [];
    if (!hasGps) flags.push("無定位");
    const siteName = String(params.siteName || "").trim().slice(0, 80);
    const note = String(params.note || "").trim().slice(0, 200);
    let photoUrl = "";

    // 範圍外 = 外勤打卡：必須有工地名稱與現場照片
    if (!inRange) {
      if (!siteName) return { success: false, needOffsite: true, message: "您不在公司或已登錄工地的範圍內，請選擇或填寫工地名稱，並拍一張現場照片。" };
      if (!params.photo) return { success: false, needOffsite: true, message: "外勤打卡需要拍一張現場照片。" };
      const saved = saveAttendancePhoto_(params.photo, me, type, now);
      if (!saved.success) return saved;
      photoUrl = saved.url;
      // 選的是已登錄工地、但實際位置離很遠 → 地點不符
      const chosen = locations.find(l => l.name === siteName);
      if (chosen && hasGps && distanceM_(lat, lng, parseFloat(chosen.lat), parseFloat(chosen.lng)) > ATT.OFFSITE_MISMATCH_M) {
        flags.push("地點不符");
      }
    }

    // 同一裝置短時間內替不同人打卡
    const deviceId = String(params.deviceId || "").slice(0, 40);
    if (deviceId) {
      const other = logs.find(l => l.device_id === deviceId && l.user_id !== me.id && l.date === today &&
        Math.abs(minOfDay_(l.punched_at) - nowMin) <= ATT.SAME_DEVICE_WINDOW_MIN);
      if (other) flags.push("同裝置多人打卡");
    }

    const record = {
      id: "ATT-" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 9000 + 1000),
      user_id: me.id, punched_at: now, date: today, type: type,
      location_id: inRange ? nearest.loc.id : "", location_name: inRange ? nearest.loc.name : "外勤",
      distance_m: nearest ? nearest.distance : "", in_range: inRange ? "TRUE" : "FALSE",
      // 隱私：只有外勤打卡才保留座標，範圍內只留距離
      lat: inRange ? "" : (hasGps ? lat : ""), lng: inRange ? "" : (hasGps ? lng : ""),
      accuracy_m: accuracy, site_name: inRange ? "" : siteName, note: note, photo_url: photoUrl,
      device_id: deviceId, flags: flags.join("、"), created_at: now
    };
    logSheet.appendRow(ATT_HEADERS.LOGS.map(h => record[h] === undefined ? "" : String(record[h])));

    const label = type === "IN" ? "上班" : "下班";
    return {
      success: true,
      message: `${label}打卡成功：${now.slice(11, 16)}，${inRange ? record.location_name : "外勤（" + siteName + "）"}` + (flags.length ? `（${flags.join("、")}）` : ""),
      record: record,
      today: todaySummary_(ss, me)
    };
  } finally {
    lock.releaseLock();
  }
}

function saveAttendancePhoto_(dataUrl, me, type, now) {
  const m = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!m) return { success: false, message: "照片格式不正確，請重新拍照。" };
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > ATT.PHOTO_MAX_BYTES) return { success: false, message: "照片太大，請重新拍照。" };

  const props = PropertiesService.getScriptProperties();
  let root = null;
  const rootId = props.getProperty(ATT.PHOTO_FOLDER_PROP);
  if (rootId) { try { root = DriveApp.getFolderById(rootId); } catch (e) { root = null; } }
  if (!root) {
    root = DriveApp.createFolder("請假系統_外勤打卡照片");
    props.setProperty(ATT.PHOTO_FOLDER_PROP, root.getId());
  }
  const month = now.slice(0, 7);
  const it = root.getFoldersByName(month);
  const folder = it.hasNext() ? it.next() : root.createFolder(month);
  const ext = m[1] === "image/png" ? "png" : (m[1] === "image/webp" ? "webp" : "jpg");
  const name = `${now.slice(0, 10)}_${now.slice(11, 19).replace(/:/g, "")}_${me.id}_${type}.${ext}`;
  const file = folder.createFile(Utilities.newBlob(bytes, m[1], name));
  return { success: true, url: file.getUrl() };
}

// ======================== 每日判定 ========================

function holidayMap_(ss) {
  const map = {};
  sheetToObjects(ss.getSheetByName(CONFIG.SHEETS.HOLIDAYS)).forEach(h => {
    const d = String(h.date || "").slice(0, 10).replace(/\//g, "-");
    if (d) map[d] = { name: h.name, isWorkday: h.is_workday === true || String(h.is_workday).toUpperCase() === "TRUE" };
  });
  return map;
}

function isWorkday_(dateStr, holidays) {
  const h = holidays[dateStr];
  if (h) return h.isWorkday;
  const p = dateStr.split("-").map(Number);
  const dow = new Date(p[0], p[1] - 1, p[2]).getDay();
  return dow !== 0 && dow !== 6;
}

// 當天被請假涵蓋的時段 (分鐘)，只看已核准 (含銷假審核中) 與審核中
function leaveCoverage_(requests, userId, dateStr) {
  const out = { approved: [], pending: [], ids: [] };
  requests.forEach(r => {
    if (r.user_id !== userId) return;
    const st = String(r.status || "");
    if (["APPROVED", "CANCEL_PENDING", "PENDING"].indexOf(st) === -1) return;
    const s = String(r.start_time || "").replace(/\//g, "-"), e = String(r.end_time || "").replace(/\//g, "-");
    const sd = s.slice(0, 10), ed = e.slice(0, 10);
    if (!sd || dateStr < sd || dateStr > ed) return;
    const from = dateStr === sd ? minOfDay_(s) : 0;
    const to = dateStr === ed ? minOfDay_(e) : 24 * 60;
    (st === "PENDING" ? out.pending : out.approved).push([from, to]);
    out.ids.push(r.id);
  });
  return out;
}

/**
 * 判定某人某天的出勤。回傳 attendance_daily 的一列 (物件)。
 * isToday：當天尚未結束時，不判缺下班卡 / 早退。
 */
function judgeDay_(user, dateStr, dayLogs, requests, holidays, isToday) {
  const row = { date: dateStr, user_id: user.id, status: "", first_in: "", last_out: "", expected_in: "", expected_out: "",
    late_min: 0, early_min: 0, stay_after_min: 0, offsite: "", leave_ids: "", flags: "", judged_at: nowStr_() };

  const ins = dayLogs.filter(l => l.type === "IN").map(l => minOfDay_(l.punched_at)).sort((a, b) => a - b);
  const outs = dayLogs.filter(l => l.type === "OUT").map(l => minOfDay_(l.punched_at)).sort((a, b) => a - b);
  const firstIn = ins.length ? ins[0] : null;
  const lastOut = outs.length ? outs[outs.length - 1] : null;
  row.first_in = hhmm_(firstIn);
  row.last_out = hhmm_(lastOut);
  const offsiteSites = dayLogs.filter(l => String(l.in_range).toUpperCase() !== "TRUE").map(l => l.site_name).filter(Boolean);
  row.offsite = Array.from(new Set(offsiteSites)).join("、");
  row.flags = Array.from(new Set(dayLogs.map(l => l.flags).filter(Boolean).join("、").split("、").filter(Boolean))).join("、");

  if (isTopAdmin_(user)) { row.status = "免打卡"; return row; }
  if (!isWorkday_(dateStr, holidays)) {
    row.status = dayLogs.length ? "假日出勤" : "休息日";
    return row;
  }

  const cov = leaveCoverage_(requests, user.id, dateStr);
  row.leave_ids = cov.ids.join(",");
  if (cov.pending.length) { row.status = "請假審核中"; return row; }

  // 已核准請假：從早上開始的假 (例如上午假、08:30～10:00) 與一直請到下班的假 (例如下午假、16:00～18:00)
  const DAY_END = 18 * 60;
  const startLeaves = cov.approved.filter(([f, t]) => f <= ATT.ARRIVE_FROM && t > ATT.ARRIVE_FROM);
  const endLeaves = cov.approved.filter(([f, t]) => t >= DAY_END && f < DAY_END);
  const coversStart = startLeaves.length > 0;
  const coversEnd = endLeaves.length > 0;
  let leaveEnd = coversStart ? Math.max.apply(null, startLeaves.map(([, t]) => t)) : null;
  let leaveStart = coversEnd ? Math.min.apply(null, endLeaves.map(([f]) => f)) : null;
  if (coversStart && leaveEnd >= ATT.LUNCH_START && leaveEnd < ATT.LUNCH_END) leaveEnd = ATT.LUNCH_END;
  if (coversStart && coversEnd && leaveEnd >= leaveStart) { row.status = "請假"; return row; }
  if (coversStart && leaveEnd >= DAY_END) { row.status = "請假"; return row; }

  // 應上班時間
  let expectedIn, effArrival, late = 0;
  if (coversStart) {
    // 早上請假：請假結束就要到 (不適用彈性)
    expectedIn = leaveEnd;
    effArrival = expectedIn;
    if (firstIn !== null && firstIn > expectedIn) late = firstIn - expectedIn;
  } else {
    expectedIn = ATT.ARRIVE_UNTIL;
    if (firstIn === null) effArrival = ATT.ARRIVE_UNTIL;
    else if (firstIn <= ATT.ARRIVE_FROM) effArrival = ATT.ARRIVE_FROM;
    else if (firstIn <= ATT.ARRIVE_UNTIL) effArrival = firstIn;
    else { effArrival = ATT.ARRIVE_UNTIL; late = firstIn - ATT.ARRIVE_UNTIL; }
  }

  // 應下班時間
  let expectedOut;
  if (coversEnd) {
    // 下午假 (13:30 起)：上午做滿 3.5 小時 (12:00～12:30，依到班時間)；其他時間開始的假：請假開始就可下班
    expectedOut = (leaveStart <= ATT.LUNCH_END && !coversStart)
      ? Math.min(effArrival + 210, ATT.ARRIVE_UNTIL + 210)
      : Math.min(leaveStart, ATT.LUNCH_START <= leaveStart && leaveStart < ATT.LUNCH_END ? ATT.LUNCH_START : leaveStart);
  } else if (coversStart) {
    expectedOut = DAY_END;
  } else {
    expectedOut = Math.min(effArrival + ATT.DAY_SPAN_MIN, ATT.ARRIVE_UNTIL + ATT.DAY_SPAN_MIN);
  }
  row.expected_in = coversStart ? hhmm_(expectedIn) : "08:30~09:00";
  row.expected_out = hhmm_(expectedOut);

  const issues = [];
  if (firstIn === null && lastOut === null) {
    row.status = isToday ? "尚未打卡" : "未打卡";
    return row;
  }
  if (firstIn === null) issues.push("缺上班卡");
  if (late > 0) { row.late_min = late; issues.push("遲到"); }
  if (lastOut === null) {
    if (!isToday) issues.push("缺下班卡");
  } else if (lastOut < expectedOut) {
    if (!isToday) { row.early_min = expectedOut - lastOut; issues.push("早退"); }
  } else {
    row.stay_after_min = lastOut - expectedOut;
    if (row.stay_after_min >= ATT.STAY_AFTER_ALERT_MIN) issues.push("晚下班");
  }

  row.status = issues.length ? issues.join("、") : (isToday && lastOut === null ? "上班中" : "正常");
  if (cov.approved.length && row.status === "正常") row.status = "正常（部分請假）";
  return row;
}

function todaySummary_(ss, me) {
  const today = todayStr_();
  const logs = sheetToObjects(ss.getSheetByName(ATT.SHEETS.LOGS)).filter(l => l.user_id === me.id && l.date === today);
  const requests = sheetToObjects(ss.getSheetByName(CONFIG.SHEETS.LEAVE_REQUESTS));
  const judged = judgeDay_(me, today, logs, requests, holidayMap_(ss), true);
  return { date: today, logs: logs.map(stripLog_), judged: judged };
}

function stripLog_(l) {
  return {
    id: l.id, punched_at: l.punched_at, date: l.date, type: l.type, location_name: l.location_name,
    distance_m: l.distance_m, in_range: String(l.in_range).toUpperCase() === "TRUE", site_name: l.site_name,
    note: l.note, has_photo: !!l.photo_url, photo_url: l.photo_url, flags: l.flags
  };
}

/**
 * 查詢某人某月的出勤：每天的判定 + 打卡明細 + 月統計
 * params: { month: "yyyy-MM", targetUserId? }
 */
function getAttendance(params) {
  const me = params._me;
  const ss = ensureAttendanceSheets_();
  const users = sheetToObjects(ss.getSheetByName(CONFIG.SHEETS.USERS));
  const target = users.find(u => u.id === (params.targetUserId || me.id));
  if (!target) return { success: false, message: "找不到這位同仁。" };
  if (!canViewAttendanceOf_(me, target)) return { success: false, message: "您沒有權限查看這位同仁的出勤。" };

  const today = todayStr_();
  const month = /^\d{4}-\d{2}$/.test(String(params.month || "")) ? params.month : today.slice(0, 7);
  const [y, mo] = month.split("-").map(Number);
  const lastDay = new Date(y, mo, 0).getDate();

  const logs = sheetToObjects(ss.getSheetByName(ATT.SHEETS.LOGS)).filter(l => l.user_id === target.id && String(l.date).slice(0, 7) === month);
  const requests = sheetToObjects(ss.getSheetByName(CONFIG.SHEETS.LEAVE_REQUESTS));
  const holidays = holidayMap_(ss);

  const days = [];
  const stats = { workdays: 0, attended: 0, late: 0, lateMin: 0, early: 0, earlyMin: 0, missing: 0, absent: 0, offsite: 0, leave: 0, stayAfter: 0 };
  for (let d = 1; d <= lastDay; d++) {
    const ds = month + "-" + ("0" + d).slice(-2);
    if (ds > today) break;
    const dayLogs = logs.filter(l => l.date === ds);
    const j = judgeDay_(target, ds, dayLogs, requests, holidays, ds === today);
    j.holiday = holidays[ds] && !holidays[ds].isWorkday ? holidays[ds].name : "";
    j.logs = dayLogs.map(stripLog_);
    days.push(j);
    if (isWorkday_(ds, holidays) && j.status !== "免打卡") stats.workdays++;
    if (j.first_in || j.last_out) stats.attended++;
    if (j.late_min > 0) { stats.late++; stats.lateMin += j.late_min; }
    if (j.early_min > 0) { stats.early++; stats.earlyMin += j.early_min; }
    if (/缺上班卡|缺下班卡/.test(j.status)) stats.missing++;
    if (j.status === "未打卡") stats.absent++;
    if (j.offsite) stats.offsite++;
    if (/請假/.test(j.status)) stats.leave++;
    if (/晚下班/.test(j.status)) stats.stayAfter++;
  }

  // 可查看的同仁清單 (給主管 / 人資切換)
  const viewable = users.filter(u => canViewAttendanceOf_(me, u) && !isUserDisabled_(u))
    .map(u => ({ id: u.id, name: u.name, department_name: u.department_name }));

  return {
    success: true,
    month: month,
    user: { id: target.id, name: target.name, department_name: target.department_name },
    days: days.reverse(),
    stats: stats,
    viewable: viewable,
    // 照片連結只給可管理出勤的人看 (人資 / 管理者)；一般同仁看到「已拍照」
    canSeePhotos: isAdminOrHR_(me),
    // 人資部 / 管理部：匯出用的同仁清單
    canExport: canExportAttendance_(me),
    exportUsers: canExportAttendance_(me)
      ? users.filter(u => !isUserDisabled_(u) && !isTopAdmin_(u)).map(u => ({ id: u.id, name: u.name, department_name: u.department_name }))
      : []
  };
}

// ======================== 今日出勤總覽 (一次看全部同仁) ========================

/**
 * 某一天所有同仁的打卡狀態。超級使用者 / 人資 / 管理者看全部；主管看直屬部屬 (含自己)。
 * params: { date?: "yyyy-MM-dd" (預設今天) }
 * 不列已停用帳號與免打卡的最高主管。
 */
function getAttendanceBoard(params) {
  const me = params._me;
  const isAll = isAdminOrHR_(me);
  if (!isAll && me.role !== "Manager") return { success: false, message: "只有主管、人資或管理者可以查看出勤總覽。" };

  const today = todayStr_();
  let date = /^\d{4}-\d{2}-\d{2}$/.test(String(params.date || "")) ? params.date : today;
  if (date > today) date = today;
  const isToday = date === today;

  const ss = ensureAttendanceSheets_();
  const users = sheetToObjects(ss.getSheetByName(CONFIG.SHEETS.USERS))
    .filter(u => u.id && !isUserDisabled_(u) && !isTopAdmin_(u) && canViewAttendanceOf_(me, u));
  const logs = sheetToObjects(ss.getSheetByName(ATT.SHEETS.LOGS)).filter(l => String(l.date).slice(0, 10) === date);
  const requests = sheetToObjects(ss.getSheetByName(CONFIG.SHEETS.LEAVE_REQUESTS));
  const holidays = holidayMap_(ss);
  const nowMin = minOfDay_(nowStr_());
  const canSeePhotos = isAdminOrHR_(me);

  const rows = users.map(u => {
    const dayLogs = logs.filter(l => l.user_id === u.id).sort((a, b) => String(a.punched_at).localeCompare(String(b.punched_at)));
    const j = judgeDay_(u, date, dayLogs, requests, holidays, isToday);
    const hire = String(u.hire_date || "").slice(0, 10).replace(/\//g, "-");
    if (hire && date < hire) j.status = "尚未到職";
    // 今天過了 09:00 還沒打上班卡 → 標示「未到班」
    const overdue = isToday && j.status === "尚未打卡" && nowMin > ATT.ARRIVE_UNTIL;
    return {
      user_id: u.id, name: u.name, department_name: u.department_name || "",
      status: overdue ? "未到班" : j.status,
      first_in: j.first_in, last_out: j.last_out, expected_out: j.expected_out,
      late_min: j.late_min, early_min: j.early_min, stay_after_min: j.stay_after_min,
      offsite: j.offsite, flags: j.flags,
      logs: dayLogs.map(l => {
        const s = stripLog_(l);
        if (!canSeePhotos) s.photo_url = "";
        return s;
      })
    };
  });

  // 有狀況的排前面：未到班 / 缺卡 → 遲到早退 → 還沒到 09:00 → 正常 → 請假 → 休息
  const order = s => /未到班|未打卡|缺/.test(s) ? 0 : /遲到|早退/.test(s) ? 1 : /尚未打卡/.test(s) ? 2 : /請假/.test(s) ? 4 : /休息|免打卡|到職/.test(s) ? 5 : 3;
  rows.sort((a, b) => order(a.status) - order(b.status) ||
    String(a.department_name).localeCompare(String(b.department_name)) || String(a.user_id).localeCompare(String(b.user_id)));

  const count = re => rows.filter(r => re.test(r.status)).length;
  return {
    success: true, date: date, isToday: isToday, now: isToday ? hhmm_(nowMin) : "",
    workday: isWorkday_(date, holidays),
    holiday: holidays[date] && !holidays[date].isWorkday ? holidays[date].name : "",
    scope: isAll ? "全部同仁" : "直屬部屬",
    canSeePhotos: canSeePhotos,
    rows: rows,
    summary: {
      total: rows.length,
      arrived: rows.filter(r => r.first_in).length,
      late: rows.filter(r => r.late_min > 0).length,
      notIn: count(/未到班|未打卡|尚未打卡/),
      missing: count(/缺上班卡|缺下班卡/),
      leave: count(/請假/),
      offsite: rows.filter(r => r.offsite).length,
      out: rows.filter(r => r.last_out).length
    }
  };
}

// ======================== 出勤匯出 (人資部 / 管理部) ========================

const ATT_EXPORT_DEPTS = ["DEPT_HR", "DEPT_MGMT"];
const ATT_EXPORT_DEPT_NAMES = ["人資部", "人力資源部", "管理部"];
const ATT_EXPORT_MAX_DAYS = 366;

// 人資部、管理部的同仁 (不論角色)，以及人資 / 管理者角色、超級使用者，可以匯出出勤
function canExportAttendance_(user) {
  if (!user) return false;
  if (isAdminOrHR_(user)) return true;
  return ATT_EXPORT_DEPTS.indexOf(String(user.department_id || "").trim()) !== -1 ||
    ATT_EXPORT_DEPT_NAMES.indexOf(String(user.department_name || "").trim()) !== -1;
}

/**
 * 匯出出勤：每日明細 + 每人統計
 * params: { from: "yyyy-MM-dd", to: "yyyy-MM-dd", targetUserId: 員工編號 或 "ALL" }
 * 「全部」不含已停用帳號與免打卡的最高主管 (有打卡紀錄的已停用同仁仍會列入)
 */
function exportAttendance(params) {
  const me = params._me;
  if (!canExportAttendance_(me)) return { success: false, message: "只有人資部、管理部可以匯出出勤紀錄。" };

  const today = todayStr_();
  const re = /^\d{4}-\d{2}-\d{2}$/;
  const from = String(params.from || ""), to0 = String(params.to || "");
  if (!re.test(from) || !re.test(to0)) return { success: false, message: "請選擇開始與結束日期。" };
  if (from > to0) return { success: false, message: "開始日期不能晚於結束日期。" };
  if (from > today) return { success: false, message: "開始日期不能晚於今天。" };
  const to = to0 > today ? today : to0;
  const dayCount = Math.round((Date.parse(to) - Date.parse(from)) / 86400000) + 1;
  if (dayCount > ATT_EXPORT_MAX_DAYS) return { success: false, message: "一次最多匯出 366 天，請縮小日期範圍。" };

  const ss = ensureAttendanceSheets_();
  const users = sheetToObjects(ss.getSheetByName(CONFIG.SHEETS.USERS)).filter(u => u.id);
  const logs = sheetToObjects(ss.getSheetByName(ATT.SHEETS.LOGS)).filter(l => {
    const d = String(l.date).slice(0, 10);
    return d >= from && d <= to;
  });

  const all = String(params.targetUserId || "ALL") === "ALL";
  let targets;
  if (all) {
    const hasLogs = {};
    logs.forEach(l => { hasLogs[l.user_id] = true; });
    targets = users.filter(u => !isTopAdmin_(u) && (!isUserDisabled_(u) || hasLogs[u.id]));
  } else {
    const t = users.find(u => u.id === params.targetUserId);
    if (!t) return { success: false, message: "找不到這位同仁。" };
    targets = [t];
  }
  targets.sort((a, b) => String(a.department_name).localeCompare(String(b.department_name)) || String(a.id).localeCompare(String(b.id)));

  const requests = sheetToObjects(ss.getSheetByName(CONFIG.SHEETS.LEAVE_REQUESTS));
  const holidays = holidayMap_(ss);
  const dates = [];
  for (let i = 0; i < dayCount; i++) {
    dates.push(Utilities.formatDate(new Date(Date.parse(from) + i * 86400000 + 12 * 3600000), "UTC", "yyyy-MM-dd"));
  }

  const rows = [], summary = [];
  targets.forEach(u => {
    const myLogs = logs.filter(l => l.user_id === u.id);
    const hire = String(u.hire_date || "").slice(0, 10).replace(/\//g, "-");
    const st = { workdays: 0, attended: 0, late: 0, lateMin: 0, early: 0, earlyMin: 0, missing: 0, absent: 0, offsite: 0, leave: 0, stayAfter: 0, stayAfterMin: 0 };
    dates.forEach(ds => {
      if (hire && ds < hire) return;
      const dayLogs = myLogs.filter(l => String(l.date).slice(0, 10) === ds);
      const j = judgeDay_(u, ds, dayLogs, requests, holidays, ds === today);
      const workday = isWorkday_(ds, holidays);
      if (workday && j.status !== "免打卡") st.workdays++;
      if (j.first_in || j.last_out) st.attended++;
      if (j.late_min > 0) { st.late++; st.lateMin += j.late_min; }
      if (j.early_min > 0) { st.early++; st.earlyMin += j.early_min; }
      if (/缺上班卡|缺下班卡/.test(j.status)) st.missing++;
      if (j.status === "未打卡") st.absent++;
      if (j.offsite) st.offsite++;
      if (/請假/.test(j.status)) st.leave++;
      if (/晚下班/.test(j.status)) { st.stayAfter++; st.stayAfterMin += j.stay_after_min; }
      // 只列工作日與有打卡的假日，休息日不佔列
      if (!workday && !dayLogs.length) return;
      rows.push({
        date: ds, user_id: u.id, name: u.name, department_name: u.department_name || "",
        holiday: holidays[ds] && !holidays[ds].isWorkday ? holidays[ds].name : "",
        status: j.status, first_in: j.first_in, last_out: j.last_out, expected_in: j.expected_in, expected_out: j.expected_out,
        late_min: j.late_min, early_min: j.early_min, stay_after_min: j.stay_after_min,
        offsite: j.offsite, flags: j.flags,
        punches: dayLogs.map(l => (l.type === "IN" ? "上 " : "下 ") + String(l.punched_at).slice(11, 16) + " " +
          (String(l.in_range).toUpperCase() === "TRUE" ? l.location_name : "外勤：" + (l.site_name || ""))).join("；")
      });
    });
    summary.push(Object.assign({ user_id: u.id, name: u.name, department_name: u.department_name || "" }, st));
  });

  return {
    success: true, from: from, to: to, scope: all ? "ALL" : targets[0].id,
    scopeName: all ? "全部同仁" : targets[0].name,
    rows: rows, summary: summary, exportedBy: me.name, exportedAt: nowStr_()
  };
}

// ======================== 打卡地點 (公司 / 工地) ========================

function listLocations(params) {
  const ss = ensureAttendanceSheets_();
  const today = todayStr_();
  const locs = sheetToObjects(ss.getSheetByName(ATT.SHEETS.LOCATIONS)).map(l => Object.assign({}, l, { active: isLocationActive_(l, today) }));
  // 最近 30 天外勤打卡用過的工地名稱 (給外勤打卡下拉選單)
  const since = Utilities.formatDate(new Date(Date.now() - 30 * 86400000), "Asia/Taipei", "yyyy-MM-dd");
  const recentSites = Array.from(new Set(
    sheetToObjects(ss.getSheetByName(ATT.SHEETS.LOGS)).filter(l => l.site_name && l.date >= since).map(l => l.site_name)
  )).slice(-30);
  return {
    success: true,
    locations: locs.map(l => ({ id: l.id, name: l.name, type: l.type, address: l.address, lat: l.lat, lng: l.lng,
      radius_m: l.radius_m, start_date: l.start_date, end_date: l.end_date, active: l.active })),
    recentSites: recentSites,
    canManage: canManageLocations_(params._me)
  };
}

/**
 * 新增 / 修改打卡地點。params: { id?, name, type, address?, lat?, lng?, radius_m, start_date, end_date }
 * 沒給座標時以地址自動定位。
 */
function saveLocation(params) {
  const me = params._me;
  if (!canManageLocations_(me)) return { success: false, message: "只有主管、人資或管理者可以設定打卡地點。" };
  const name = String(params.name || "").trim().slice(0, 60);
  if (!name) return { success: false, message: "請輸入地點名稱。" };
  const type = params.type === "公司" ? "公司" : "工地";
  let lat = parseFloat(params.lat), lng = parseFloat(params.lng);
  const address = String(params.address || "").trim().slice(0, 120);
  if (!(isFinite(lat) && isFinite(lng))) {
    if (!address) return { success: false, message: "請輸入地址，或在現場按「使用目前位置」。" };
    const geo = geocodeAddress_(address);
    if (!geo) return { success: false, message: "找不到這個地址的位置，請確認地址，或在現場按「使用目前位置」。" };
    lat = geo.lat; lng = geo.lng;
  }
  const radius = Math.min(1000, Math.max(50, parseInt(params.radius_m, 10) || ATT.DEFAULT_RADIUS_M));
  const dateOk = v => !v || /^\d{4}-\d{2}-\d{2}$/.test(v);
  if (!dateOk(params.start_date) || !dateOk(params.end_date)) return { success: false, message: "日期格式不正確。" };
  if (params.start_date && params.end_date && params.end_date < params.start_date) return { success: false, message: "結束日期不能早於開始日期。" };

  const ss = ensureAttendanceSheets_();
  const sheet = ss.getSheetByName(ATT.SHEETS.LOCATIONS);
  const values = { name, type, address, lat: String(lat), lng: String(lng), radius_m: String(radius),
    start_date: params.start_date || "", end_date: params.end_date || "" };

  if (params.id) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === params.id) {
        Object.keys(values).forEach(k => sheet.getRange(i + 1, ATT_HEADERS.LOCATIONS.indexOf(k) + 1).setValue(values[k]));
        return { success: true, message: `已更新打卡地點「${name}」。` };
      }
    }
    return { success: false, message: "找不到這個打卡地點。" };
  }
  const id = "LOC-" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMddHHmmss");
  const rec = Object.assign({ id, created_by: me.id, created_at: nowStr_() }, values);
  sheet.appendRow(ATT_HEADERS.LOCATIONS.map(h => rec[h] === undefined ? "" : rec[h]));
  return { success: true, message: `已新增打卡地點「${name}」，半徑 ${radius} 公尺。` };
}

// ======================== 每日自動判定 (由每日早上的通知排程呼叫) ========================

function judgeAttendanceForDate_(dateStr) {
  const ss = ensureAttendanceSheets_();
  const users = sheetToObjects(ss.getSheetByName(CONFIG.SHEETS.USERS)).filter(u => u.id && !isUserDisabled_(u));
  const logs = sheetToObjects(ss.getSheetByName(ATT.SHEETS.LOGS)).filter(l => l.date === dateStr);
  const requests = sheetToObjects(ss.getSheetByName(CONFIG.SHEETS.LEAVE_REQUESTS));
  const holidays = holidayMap_(ss);
  const dailySheet = ss.getSheetByName(ATT.SHEETS.DAILY);

  // 同一天重算：先刪掉舊的判定再寫入
  const data = dailySheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === dateStr) dailySheet.deleteRow(i + 1);
  }
  const rows = users.map(u => {
    const j = judgeDay_(u, dateStr, logs.filter(l => l.user_id === u.id), requests, holidays, false);
    return ATT_HEADERS.DAILY.map(h => String(j[h] === undefined ? "" : j[h]));
  });
  if (rows.length) dailySheet.getRange(dailySheet.getLastRow() + 1, 1, rows.length, ATT_HEADERS.DAILY.length).setValues(rows);
  return rows.length;
}

// 手動執行：重算昨天
function judgeYesterdayAttendance() {
  const y = Utilities.formatDate(new Date(Date.now() - 86400000), "Asia/Taipei", "yyyy-MM-dd");
  const n = judgeAttendanceForDate_(y);
  Logger.log(`已判定 ${y} 共 ${n} 人的出勤`);
  return n;
}
