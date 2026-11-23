/**
 * 程式版本備份到 Google Drive (GitHub → Drive 版本控管)
 * -------------------------------------------------------------
 * GitHub 是正式的版本控管 (每次修改都是一個 commit，可隨時還原)；
 * 本程式把每一個版本的完整程式碼，另外快照備份到 Google Drive：
 *
 *   請假系統/
 *   ├── 版本紀錄 (試算表：每個版本一列，含日期、說明、資料夾與 GitHub 連結)
 *   ├── 2026-11-03_1254_cd6a419/   ← 每個版本一個資料夾，保留原本的目錄結構
 *   │     ├── 原始碼_cd6a419.zip     ← 整包下載用
 *   │     ├── index.html, css/, js/, google-apps-script/ ...
 *   └── ...
 *
 * 使用方式 (在 Apps Script 編輯器)：
 *   1. 新增指令碼檔案 DriveBackup，貼上本檔，存檔。
 *   2. 執行 backupHistoryToDrive  → 一次備份 BACKUP_CONFIG.INITIAL_COMMITS 列出的歷史版本。
 *   3. 執行 installBackupTrigger  → 之後每天自動檢查 GitHub，有新版本就備份。
 *   隨時也可手動執行 backupLatestToDrive 立即備份最新版。
 */

const BACKUP_CONFIG = {
  OWNER: "aray6000-terry",
  REPO: "PCSI__OCTA_API",
  BRANCH: "main",
  FOLDER_ID: "1WOS5oFhmpWN7fHUySonmUDNgVlYzjDeb",   // Google Drive「請假系統」資料夾
  LOG_NAME: "版本紀錄",
  // 第一次執行 backupHistoryToDrive 時要補備份的歷史版本 (舊 → 新)
  INITIAL_COMMITS: ["759d612", "46d3686", "cd6a419", "d0cb8e5"],
  // 版本中文說明 (沒有列在這裡的版本會使用 GitHub 上的 commit 說明)
  NOTES: {
    "759d612": "原始版本（新增請假日期後 7 天銷假失效限制）",
    "46d3686": "資安修正：後端強制登入驗證、密碼雜湊、簽核權限檢查、移除後門",
    "cd6a419": "XSS 防護、跨年度額度自動建立、婚喪產假可申請、特休算法前後端統一",
    "d0cb8e5": "新增 GitHub → Google Drive 每日版本備份",
    "926ab90": "修正備份程式：改用 GitHub 公開下載連結，避開 API 次數限制",
    "6dd27ea": "勞基法規則：補休 1:1、生理假半薪併病假、事假 14 日 / 病假 30 日上限警告",
    "11d75a3": "老闆 (最高 Admin) 可自核自己的假單，紀錄標註本人自核",
    "28a5c1a": "老闆核准即生效，不再轉人資第二階"
  }
};

/** 每天自動備份 (台北時間凌晨 2 點左右) */
function installBackupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "backupLatestToDrive")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("backupLatestToDrive").timeBased().everyDays(1).atHour(2).inTimezone("Asia/Taipei").create();
  return "已設定：每天凌晨自動把 GitHub 最新版本備份到 Google Drive";
}

/** 備份歷史版本 (已備份過的會自動略過) */
function backupHistoryToDrive() {
  const results = BACKUP_CONFIG.INITIAL_COMMITS.map(sha => backupCommitToDrive_(sha));
  Logger.log(results.join("\n"));
  return results;
}

/** 備份 main 分支最新版本 (已備份過就略過) */
function backupLatestToDrive() {
  const msg = backupCommitToDrive_(BACKUP_CONFIG.BRANCH);
  Logger.log(msg);
  return msg;
}

// ---------------------------------------------------------------
// repo 已改為「私人」：需要一把唯讀的 GitHub 存取權杖 (Fine-grained token，只給 PCSI__OCTA_API 的 Contents: Read-only)，
// 存在「指令碼屬性」GITHUB_TOKEN。帶權杖時每小時有 5000 次額度，不會再遇到匿名 403。
// 若沒有權杖、repo 又是公開的，會自動改用公開連結。

function githubToken_() {
  return PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN") || "";
}

function githubFetch_(url, accept) {
  const token = githubToken_();
  const headers = { "User-Agent": "leave-system-drive-backup", "Accept": accept || "application/vnd.github+json" };
  if (token) headers.Authorization = "Bearer " + token;
  return UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true, followRedirects: true });
}

function backupCommitToDrive_(ref) {
  const commits = listRecentCommits_();
  const commit = (ref === BACKUP_CONFIG.BRANCH) ? commits[0] : commits.find(c => c.sha.indexOf(ref) === 0);
  if (!commit) return `找不到版本 ${ref}（只能備份最近 20 個版本）`;

  const short = commit.sha.substring(0, 7);
  const log = getOrCreateLog_();
  const done = log.getDataRange().getValues().some(r => String(r[2]) === short);
  if (done) return `略過 ${short}：已經備份過`;

  // 下載該版本完整原始碼 zip
  const zipUrl = githubToken_()
    ? `https://api.github.com/repos/${BACKUP_CONFIG.OWNER}/${BACKUP_CONFIG.REPO}/zipball/${commit.sha}`
    : `https://codeload.github.com/${BACKUP_CONFIG.OWNER}/${BACKUP_CONFIG.REPO}/zip/${commit.sha}`;
  const res = githubFetch_(zipUrl);
  if (res.getResponseCode() !== 200) {
    throw new Error(`下載原始碼失敗 (HTTP ${res.getResponseCode()})。repo 為私人時，請在指令碼屬性設定 GITHUB_TOKEN。`);
  }
  const zipBlob = res.getBlob().setName(`原始碼_${short}.zip`).setContentType("application/zip");

  const folderName = Utilities.formatDate(commit.date, "Asia/Taipei", "yyyy-MM-dd_HHmm") + "_" + short;
  const versionFolder = DriveApp.getFolderById(BACKUP_CONFIG.FOLDER_ID).createFolder(folderName);
  versionFolder.createFile(zipBlob);

  // 解壓縮並保留原本的目錄結構 (去掉 zip 內最外層的 REPO-sha/ 資料夾)
  const folderCache = { "": versionFolder };
  let count = 0;
  Utilities.unzip(zipBlob.copyBlob()).forEach(b => {
    const full = b.getName();
    if (/\/$/.test(full)) return; // 資料夾項目
    const parts = full.split("/").slice(1);
    if (!parts.length || !parts[parts.length - 1]) return;
    const fileName = parts.pop();
    getSubFolder_(versionFolder, folderCache, parts).createFile(b.setName(fileName));
    count++;
  });

  const note = BACKUP_CONFIG.NOTES[short] || commit.title;
  const version = "v" + log.getLastRow(); // 第 1 列是標題，第一個版本為 v1
  log.appendRow([
    version,
    Utilities.formatDate(commit.date, "Asia/Taipei", "yyyy-MM-dd HH:mm"),
    short,
    note,
    count,
    versionFolder.getUrl(),
    `https://github.com/${BACKUP_CONFIG.OWNER}/${BACKUP_CONFIG.REPO}/commit/${commit.sha}`,
    Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm")
  ]);
  return `已備份 ${version} ${short}「${note}」共 ${count} 個檔案 → ${folderName}`;
}

// 讀取最近 20 個版本 (最新在前)：有權杖用 GitHub API，否則用公開的 commit 動態
function listRecentCommits_() {
  if (githubToken_()) {
    const res = githubFetch_(`https://api.github.com/repos/${BACKUP_CONFIG.OWNER}/${BACKUP_CONFIG.REPO}/commits?sha=${BACKUP_CONFIG.BRANCH}&per_page=20`);
    if (res.getResponseCode() !== 200) throw new Error(`讀取 GitHub 版本清單失敗 (HTTP ${res.getResponseCode()})，請確認 GITHUB_TOKEN 是否有效。`);
    return JSON.parse(res.getContentText()).map(c => ({
      sha: c.sha,
      title: String(c.commit.message || "").split("\n")[0],
      date: new Date(c.commit.committer.date)
    }));
  }
  const url = `https://github.com/${BACKUP_CONFIG.OWNER}/${BACKUP_CONFIG.REPO}/commits/${BACKUP_CONFIG.BRANCH}.atom`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error(`讀取 GitHub 版本清單失敗 (HTTP ${res.getResponseCode()})`);
  const atom = XmlService.getNamespace("http://www.w3.org/2005/Atom");
  const entries = XmlService.parse(res.getContentText()).getRootElement().getChildren("entry", atom);
  return entries.map(e => {
    const id = e.getChildText("id", atom) || "";
    return {
      sha: (id.match(/[0-9a-f]{40}/) || [""])[0],
      title: (e.getChildText("title", atom) || "").trim(),
      date: new Date(e.getChildText("updated", atom))
    };
  }).filter(c => c.sha);
}

function getSubFolder_(base, cache, parts) {
  let path = "", cur = base;
  parts.forEach(p => {
    path = path ? path + "/" + p : p;
    if (!cache[path]) cache[path] = cur.createFolder(p);
    cur = cache[path];
  });
  return cur;
}

function getOrCreateLog_() {
  const root = DriveApp.getFolderById(BACKUP_CONFIG.FOLDER_ID);
  const it = root.getFilesByName(BACKUP_CONFIG.LOG_NAME);
  let ss;
  if (it.hasNext()) {
    ss = SpreadsheetApp.open(it.next());
  } else {
    ss = SpreadsheetApp.create(BACKUP_CONFIG.LOG_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(root);
  }
  const sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.appendRow(["版本", "版本日期", "Commit", "修改說明", "檔案數", "Drive 資料夾", "GitHub", "備份時間"]);
    sh.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#2563eb").setFontColor("#ffffff");
    sh.setFrozenRows(1);
  }
  return sh;
}
