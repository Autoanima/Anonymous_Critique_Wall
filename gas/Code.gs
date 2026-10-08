/**
 * 匿名評圖牆 — Google Apps Script 後端（v2：密語與匿名比對都在伺服器端驗證）
 *
 * 使用方式：
 * 1. 開一份新的 Google 試算表。
 * 2. 上方選單「擴充功能」→「Apps Script」，把整個檔案內容貼進去（取代原本的內容），存檔。
 * 3. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」。
 *    - 執行身分：我 (你自己的帳號)
 *    - 具有存取權的使用者：任何人
 *    - 部署後複製「網頁應用程式網址」，之後在網頁的「連接設定」畫面貼上就好，不用改程式碼。
 * 4. 之後如果有改這份程式碼，要重新「管理部署作業」→ 編輯 → 部署新版本，網址才會生效最新的程式。
 *
 * 安全性設計：
 * - doGet() 只回傳統計數字（幾件作品、幾筆評分），不含任何姓名或分數細節。
 * - 學生端的「排除自己的作品」「挑下一件要評的作品」都在伺服器內計算，姓名永遠不會傳到別的同學的瀏覽器。
 * - 教師專用的動作（看總表、改設定）都要求密語，密語比對在伺服器端做，前端拿不到密語原文。
 * - 「作品牆」（gallery）是刻意公開的，不需要密語：老師要求同學也能瀏覽全部作品、姓名、分數，
 *   跟評分過程中的匿名機制是分開的兩件事，評分時同學還是看不到彼此身分。
 * - 學生上傳的圖片會自動存進你 Google 帳號裡一個叫「匿名評圖牆 - 作品圖片」的雲端硬碟資料夾（第一次有人交作品時自動建立），
 *   並自動設成「知道連結的人可檢視」，不需要老師或學生手動處理權限。
 */

/**
 * 這個函式本身沒有實際功能，只是用來手動觸發一次 Google 的權限授權畫面。
 * 用法：上方函式下拉選單選 authorize → 按執行 ▶ → 跳出授權視窗就照著同意到底。
 * 如果沒跳出視窗、直接在記錄看到「授權沒問題」，代表權限本來就是好的。
 *
 * 注意：這裡故意實際建立一個小測試檔案（不是只用讀取類的呼叫），
 * 因為 Google 只會照你「實際執行過的程式碼」去要求對應範圍的權限——
 * 只呼叫讀取類的方法（例如 getRootFolder）沒辦法連帶要到「建立檔案／
 * 資料夾、設定分享」這種寫入類的權限，這正是圖片上傳失敗的真正原因。
 */
function authorize() {
  var tinyPng = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
  var url = saveUploadedImage('authorize_test.png', tinyPng, 'image/png');
  SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('授權沒問題，測試圖片網址：' + url + '（這個測試檔案可以之後手動刪掉）');
}

// ---- 多班級 ----
// 同一個後端可以同時服務好幾個班。班級用「班級代號」區分，前端每次請求都會帶 cls 參數：
// - 代號是空的 = 原本的那一班，分頁名稱、快取、屬性全部維持原樣（舊資料完全不用搬）。
// - 其他班的所有分頁名稱前面加「代號_」（例如 乙_Submissions、乙_Config），快取 key、Script Properties 也各自加上班級，
//   所以資料彼此完全看不到。每個班有自己的教師密語、名單、作業、圖片資料夾。
// 每一次請求（執行）都是全新的環境，這個全域變數只屬於這一次請求，不會影響到別的請求。
var CURRENT_CLASS = '';
function cleanClassId(v) {
  var t = String(v == null ? '' : v).trim();
  return /^[A-Za-z0-9一-鿿]{1,12}$/.test(t) ? t : '';
}
function clsPrefix() { return CURRENT_CLASS ? CURRENT_CLASS + '_' : ''; }
function clsNs() { return CURRENT_CLASS ? CURRENT_CLASS + ':' : ''; }
function clsPropKey(k) { return CURRENT_CLASS ? k + '__' + CURRENT_CLASS : k; }
function classExists(id, fresh) {
  if (!id) return true;
  var cache = CacheService.getScriptCache();
  if (!fresh && cache.get('clsok:' + id)) return true;
  var ok = !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(id + '_Config');
  if (ok) cache.put('clsok:' + id, '1', 600);
  return ok;
}
/** 每次請求一開始呼叫：決定這次是哪一班。班級不存在就回傳錯誤（不會因為連結打錯字就偷偷建出一個新班）。 */
function enterClass(e) {
  var raw = e && e.parameter ? e.parameter.cls : '';
  CURRENT_CLASS = cleanClassId(raw);
  if (raw && !CURRENT_CLASS) return jsonOut({ ok: false, error: 'no_such_class' });
  if (CURRENT_CLASS && !classExists(CURRENT_CLASS)) return jsonOut({ ok: false, error: 'no_such_class' });
  return null;
}

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var full = clsPrefix() + name;
  var sh = ss.getSheetByName(full);
  if (!sh) sh = ss.insertSheet(full);
  return sh;
}

// ---- 資料快取（作品／評分／設定／作品牆結果）----
// 全班看到的東西大多一樣，每個人都重新讀整張試算表、重算一次太浪費。
// 做法：快取的 key 帶「資料版本」，任何會寫入的動作做完就換一個新版本（舊版本的快取自動作廢，
// 不用逐一刪除）；讀取端是「先拿版本、再讀資料、用當初拿到的版本存回去」，所以即使剛好有人在
// 讀的同時被改了，存回去的也只是已經作廢版本的 key，不會有人讀到，不會把舊資料誤當新資料。
// 另外設 60 秒的存活時間，給「直接去試算表手動改資料」這種不會換版本的情況一個保險。
var CACHE_CHUNK_CHARS = 30000; // 單一快取值上限 100KB，中文一字最多 3 bytes，30000 字最多 90KB
var DATA_CACHE_TTL = 60;
var _dataVersionByClass = {}; // 每個班各自的資料版本（同一次請求裡可能會切到別班建立新班級）

// 版本號除了時間還加亂數：同一毫秒內連續換版本也一定不會撞號。
function newDataVersion() {
  return Date.now() + '-' + Math.floor(Math.random() * 1000000);
}

function dataVersion() {
  var ns = clsNs();
  if (_dataVersionByClass[ns]) return _dataVersionByClass[ns];
  var cache = CacheService.getScriptCache();
  var v = cache.get(ns + 'dataVersion');
  if (!v) { v = newDataVersion(); cache.put(ns + 'dataVersion', v, 21600); }
  _dataVersionByClass[ns] = v;
  return v;
}

function bumpDataVersion() {
  var ns = clsNs();
  _dataVersionByClass[ns] = newDataVersion();
  try { CacheService.getScriptCache().put(ns + 'dataVersion', _dataVersionByClass[ns], 21600); } catch (e) {}
}

function cachedData(name, compute, fresh) {
  var cache = CacheService.getScriptCache();
  var base = clsNs() + name + '_' + dataVersion();
  if (!fresh) {
    try {
      var meta = cache.get(base);
      if (meta) {
        var n = parseInt(meta, 10);
        var keys = [];
        for (var i = 0; i < n; i++) keys.push(base + '_' + i);
        var got = cache.getAll(keys);
        var parts = [];
        for (var j = 0; j < n; j++) {
          if (got[keys[j]] == null) { parts = null; break; }
          parts.push(got[keys[j]]);
        }
        if (parts) return JSON.parse(parts.join(''));
      }
    } catch (e) {}
  }
  var value = compute();
  try {
    var json = JSON.stringify(value);
    var chunks = {};
    var count = 0;
    for (var p = 0; p < json.length; p += CACHE_CHUNK_CHARS) {
      chunks[base + '_' + count] = json.substring(p, p + CACHE_CHUNK_CHARS);
      count++;
    }
    cache.putAll(chunks, DATA_CACHE_TTL);
    cache.put(base, String(count), DATA_CACHE_TTL);
  } catch (e) {}
  return value;
}

// School timetable uses Asia/Taipei regardless of the viewer's device timezone.
function defaultClassSchedule(){
  return [['08:05','08:55'],['09:10','10:00'],['10:10','11:00'],['11:10','12:00'],['13:00','13:50'],['14:00','14:50'],['15:05','15:55']].map(function(p){return {start:p[0],end:p[1]};});
}
function scheduleMinute(value){
  if(!/^\d{2}:\d{2}$/.test(String(value))) return NaN;
  var p=value.split(':').map(Number);
  return p[0]<24 && p[1]<60 ? p[0]*60+p[1] : NaN;
}
function validClassSchedule(schedule){
  if(!Array.isArray(schedule)||schedule.length!==7) return false;
  var last=-1;
  return schedule.every(function(p){
    if(!p) return false;
    var start=scheduleMinute(p.start),end=scheduleMinute(p.end);
    var valid=isFinite(start)&&isFinite(end)&&start<end&&start>=last;
    last=end; return valid;
  });
}
function inClassAt(schedule,now){
  var minute=((now+8*3600000)%86400000)/60000;
  return schedule.some(function(p){return minute>=scheduleMinute(p.start)&&minute<scheduleMinute(p.end);});
}
function classElapsedMs(schedule,from,to){
  if(!(to>from)) return 0;
  var dayMs=86400000,offset=8*3600000;
  var first=Math.floor((from+offset)/dayMs),last=Math.floor((to+offset)/dayMs),total=0;
  // Count whole intervening days in constant time, including long-closed tabs.
  var daily=schedule.reduce(function(n,p){return n+(scheduleMinute(p.end)-scheduleMinute(p.start))*60000;},0);
  if(last-first>1) total+=(last-first-1)*daily;
  [first,last].filter(function(d,i,a){return a.indexOf(d)===i;}).forEach(function(d){
    var base=d*dayMs-offset;
    schedule.forEach(function(p){total+=Math.max(0,Math.min(to,base+scheduleMinute(p.end)*60000)-Math.max(from,base+scheduleMinute(p.start)*60000));});
  });
  return total;
}
function equipmentTotalMinutes(eq){
  if(eq && Number(eq.totalMinutes)>0) return Number(eq.totalMinutes);
  var maxN=0;
  ((eq&&eq.groups)||[]).forEach(function(g){
    var present={}; (g.members||[]).forEach(function(m){present[String(m.seat)]=m.present!==false;});
    var n=(g.order||[]).filter(function(p){return present[String(p.seat)]!==false;}).length;
    maxN=Math.max(maxN,n);
  });
  return ((eq&&eq.perPersonMinutes)||6)*Math.max(1,maxN);
}
function equipmentElapsedMs(eq,now){
  if(!eq||!eq.startedAt) return 0;
  if(eq.clockVersion!==2) return Math.max(0,(eq.pausedAt||now)-eq.startedAt-(eq.pausedTotalMs||0));
  var elapsed=Number(eq.elapsedMs)||0;
  if(!eq.pausedAt&&!eq.endedAt&&eq.resumedAt) elapsed+=classElapsedMs(eq.classSchedule||defaultClassSchedule(),eq.resumedAt,now);
  return Math.min(equipmentTotalMinutes(eq)*60000,Math.max(0,elapsed));
}

function defaultConfig() {
  return {
    className: '',          // 班級名稱：顯示在網頁標題旁，好幾個班共用同一個網頁時方便辨認
    criteria: [
      { key: 'appeal', label: '有沒有吸引力？' },
      { key: 'likeness', label: '一眼知道是肖像嗎？' },
      { key: 'composition', label: '構圖舒服嗎？' },
      { key: 'color', label: '色彩有一致性嗎？' },
      { key: 'originality', label: '有沒有自己的特色？' },
      { key: 'completion', label: '整體設計完成度' }
    ],
    reviewsPerStudent: 5,
    teacherPasscode: '0928',
    namesRevealedByAssignment: {}, // { 作業名稱: true/false }，每份作業分開設定要不要開放姓名
    revealMinReviewers: 15, // 某份作業要有幾位同學「評完」，分數／雷達圖和作者姓名才自動對全班公開
    revealMinReviews: 5,    // 每位同學在這份作業要評滿幾件，才算「評完」
    assignmentCriteria: {},
    criteriaArchive: {},
    classSchedule: defaultClassSchedule(),
    assignmentDescriptions: {}, // { 作業名稱: 要求說明文字 }，首頁交作品前會顯示給學生看
    bonusByWork: {},        // { 作品編號 workCode: 老師加分 0~3 }，在作品牆播放時由老師直接加，算進總平均（上限 5 分）
    bonusNotes: {},         // { 作品編號: { reason, at } }，老師作品加分的原因和時間（最近一次）
    attendanceBonus: {},    // { 座號: { seat, name, points } }，老師在點名時幫同學按「加分」的累計次數；每次 ATTENDANCE_BONUS 分，另外加在總分，不影響作品牆
    awardsByWork: {},       // { 作品編號: { kind, assignment, at, bonus, seat, name } }，榮譽榜（前三名＋人氣獎）登記的額外加分；不算進作品牆的總平均，另外加在「含榮譽總分」
    flaggedReviewers: {},   // { 評分者識別碼: { items: { 作品編號: { at, mean, othersMean, gap, mode } } } }，老師認為評分不合理而登記的同學；他們的評分畫面會看到提醒
    returnedByWork: {},     // { 作品編號 workCode: { reason, at } }，老師退回的作品（例如拿別人作業上傳）：總平均扣 RETURN_PENALTY，不再出現在作品牆和同學的評分抽件裡
    assignments: ['第一次作業']
  };
}

/**
 * 每一次請求都會呼叫，但檢查四個分頁的標題列要打好幾次試算表 API（每次動輒上百毫秒），
 * 全班輪詢的時候光這一步就很可觀——所以檢查通過後在快取裡記一個旗標，10 分鐘內不再重複檢查。
 */
function ensureHeaders() {
  var cache = CacheService.getScriptCache();
  if (cache.get(clsNs() + 'headersOk')) return;
  ensureHeadersUncached();
  cache.put(clsNs() + 'headersOk', '1', 600);
}

function ensureHeadersUncached() {
  var subs = getSheet('Submissions');
  if (subs.getLastRow() === 0) {
    // "assignment"、"deviceId" 都加在最後一欄，不是插在中間——這樣以後就算再加欄位，
    // 舊資料列也不會被錯位讀取（之前把 studentSeat 插在中間就讓舊資料全部跑掉過一次）。
    subs.appendRow(['id', 'studentName', 'studentSeat', 'studentNorm', 'workCode', 'title', 'driveLink', 'createdAt', 'assignment', 'deviceId', 'imageHash']);
  } else if (!subs.getRange(1, 11).getValue()) {
    subs.getRange(1, 11).setValue('imageHash'); // 舊試算表補上新欄位的標題（imageHash：用來擋同一個人重複交一模一樣的圖）
  }
  var revs = getSheet('Reviews');
  if (revs.getLastRow() === 0) {
    revs.appendRow(['id', 'workId', 'reviewerNorm', 'scoresJson', 'createdAt']);
  }
  var cfg = getSheet('Config');
  if (cfg.getLastRow() === 0) {
    cfg.appendRow(['json']);
    cfg.appendRow([JSON.stringify(defaultConfig())]);
  }
  var eq = getSheet('Equipment');
  if (eq.getLastRow() === 0) {
    eq.appendRow(['json']);
    eq.appendRow([JSON.stringify(defaultEquipment())]);
  }
  var acc = getSheet('Accounts');
  if (acc.getLastRow() === 0) acc.appendRow(['seat', 'pwHmac', 'setAt']);
}

// ---- 器材輪值（多組別，全班共用同一個時間軸，誰在用不用手動按「換人」，用時間算出來）----

// 顏色不存在資料裡，完全由前端依「目前排第幾個位置」現算（每組都一樣：
// 第 1 個位置用色盤第 1 色、第 2 個位置用第 2 色⋯），所以這裡的 order 只要存
// seat／name，不用管顏色，也不會有「換色盤後舊資料查不到顏色」的問題。

// 器材種類是全班共用的一份清單（不分組別），同一種要借好幾份就重複列名字
// （例如記憶卡要 2 張，清單裡就寫兩行「記憶卡」）。每組各自有一份「有沒有歸還」
// 的勾選狀態，長度跟 itemTypes 對齊，是「這堂課」的即時狀態，不留歷史紀錄——
// 重新排定順序（等於開新的一堂課）時會整批重置回「還沒還」。
var DEFAULT_EQUIPMENT_ITEM_TYPES = ['單眼相機', '讀卡機', '記憶卡', '記憶卡', '電池'];

function defaultEquipment() {
  return {
    groups: [],           // [{name, members:[{seat,name,present}], order:[{seat,name}], leader, equipmentStatus:[{reported,reportedBySeat,reportedByName,reportedAt,confirmed}...]}]
    itemTypes: DEFAULT_EQUIPMENT_ITEM_TYPES.slice(),
    totalMinutes: 30,
    clockVersion: 2, elapsedMs: 0, resumedAt: null, endedAt: null,
    classSchedule: defaultClassSchedule(),
    perPersonMinutes: 6,
    durationMin: 1,        // 教師模式那個左右滑桿的可調範圍下限（分鐘）
    durationMax: 15,       // 上限，老師可以在教師設定裡自己改
    startedAt: null,       // 全班共用的開始時間戳
    pausedAt: null,        // 非 null 代表目前暫停中，值是暫停當下的時間戳
    pausedTotalMs: 0       // 累計暫停掉、要從經過時間裡扣掉的毫秒數
  };
}

function defaultEquipmentStatus() {
  return { reported: false, reportedBySeat: '', reportedByName: '', reportedAt: null, confirmed: false };
}

/** 讓每組的 equipmentStatus 陣列長度跟 itemTypes 對齊：多的砍掉、少的補「還沒回報」，儘量保留原本的回報/確認紀錄。 */
function syncEquipmentReturned(eq) {
  var n = (eq.itemTypes || []).length;
  (eq.groups || []).forEach(function (g) {
    var old = g.equipmentStatus || [];
    var next = [];
    for (var i = 0; i < n; i++) next.push(old[i] || defaultEquipmentStatus());
    g.equipmentStatus = next;
  });
}

function readEquipment() {
  var sh = getSheet('Equipment');
  if (sh.getLastRow() < 2) return defaultEquipment();
  var v = sh.getRange(2, 1).getValue();
  try {
    var parsed = JSON.parse(v);
    if (parsed && parsed.groups) {
      // 舊資料可能是在加 durationMin/durationMax／itemTypes 這幾個欄位之前存的，補上預設值，
      // 不然舊班級讀出來的滑桿範圍、器材清單會是 undefined。
      if (parsed.durationMin == null) parsed.durationMin = 1;
      if (parsed.durationMax == null) parsed.durationMax = 15;
      if (parsed.itemTypes == null) parsed.itemTypes = DEFAULT_EQUIPMENT_ITEM_TYPES.slice();
      // 舊資料可能還是「單純 true/false」的 equipmentReturned 陣列（兩段式確認流程上線之前存的），
      // 把它轉成新的物件格式，已經勾過的視為「已回報也已確認」，避免升級後看起來全部憑空消失。
      (parsed.groups || []).forEach(function (g) {
        if (g.equipmentReturned && !g.equipmentStatus) {
          g.equipmentStatus = g.equipmentReturned.map(function (v) {
            return v ? { reported: true, reportedBySeat: '', reportedByName: '', reportedAt: null, confirmed: true } : defaultEquipmentStatus();
          });
          delete g.equipmentReturned;
        }
        if (g.waivedSeats == null) g.waivedSeats = [];
        if (g.waiverLog == null) g.waiverLog = [];
      });
      syncEquipmentReturned(parsed);
      return parsed;
    }
  } catch (e) {}
  return defaultEquipment();
}

function writeEquipment(obj) {
  var sh = getSheet('Equipment');
  if (sh.getLastRow() < 1) sh.appendRow(['json']);
  if (sh.getLastRow() < 2) sh.appendRow([JSON.stringify(obj)]);
  else sh.getRange(2, 1).setValue(JSON.stringify(obj));
  putEquipmentCache(obj);
}

/**
 * 全班每台裝置都會定時問一次 equipmentGet，直接讀試算表太慢；
 * 結果放進快取 10 秒，寫入的時候順手換成最新的，所以最多只會有幾秒的延遲。
 */
function putEquipmentCache(obj) {
  try { CacheService.getScriptCache().put(clsNs() + 'equipmentJson', JSON.stringify(obj), 10); } catch (e) {}
}

function findGroup(eq, name) {
  for (var i = 0; i < eq.groups.length; i++) if (eq.groups[i].name === name) return eq.groups[i];
  return null;
}

function handleEquipmentGet() {
  var hit=CacheService.getScriptCache().get(clsNs()+'equipmentJson'),eq=null;
  if(hit){try{eq=JSON.parse(hit);}catch(e){}}
  if(!eq) eq=readEquipment();
  if(eq.clockVersion!==2){
    var lock=LockService.getScriptLock();
    if(lock.tryLock(1000)){
      try{eq=readEquipment();if(eq.clockVersion!==2){upgradeEquipmentTimer(eq);writeEquipment(eq);}}
      finally{lock.releaseLock();}
    }
  }
  putEquipmentCache(eq);return {ok:true,equipment:eq};
}

/**
 * 分組名單：payload.groups = [{name, members:[{seat,name}]}]。
 * 保留舊資料裡「誰出席」「順序顏色」的設定（用座號比對），不會因為老師
 * 只是想改個名字或加一個人，就把已經點過的名、排好的順序整個洗掉。
 */
function handleEquipmentSetGroups(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var eq = readEquipment();
  var oldByName = {};
  eq.groups.forEach(function (g) { oldByName[g.name] = g; });
  var newGroups = (payload.groups || []).map(function (g) {
    var old = oldByName[g.name];
    var oldPresence = {};
    if (old) (old.members || []).forEach(function (m) { oldPresence[m.seat] = m.present; });
    var members = (g.members || []).map(function (m) {
      var seat = String(m.seat || '').trim(), name = String(m.name || '').trim();
      return { seat: seat, name: name, present: oldPresence.hasOwnProperty(seat) ? oldPresence[seat] : true };
    }).filter(function (m) { return m.name; });

    // order（轉圈順序）盡量沿用舊的，不是整組重新洗牌：
    // 名單裡刪掉的人就跟著從 order 拿掉；還在名單裡的人保留原本的順序位置，
    // 名字如果改過（例如修正打錯字）就跟著更新；新加進來、原本沒在 order 裡的人
    // 直接接在最後面。這樣老師事後幫組別加人，不用整組重新排順序、
    // 也不會像「重新排定順序」那樣把全班共用的計時歸零，影響到正在進行的課堂。
    var memberBySeats = {};
    members.forEach(function (m) { memberBySeats[m.seat] = m; });
    var oldOrder = old ? (old.order || []) : [];
    var order = oldOrder
      .filter(function (p) { return memberBySeats.hasOwnProperty(p.seat); })
      .map(function (p) { return { seat: p.seat, name: memberBySeats[p.seat].name }; });
    members.forEach(function (m) {
      var already = order.some(function (p) { return p.seat === m.seat; });
      if (!already) order.push({ seat: m.seat, name: m.name });
    });

    return {
      name: g.name, members: members, order: order, leader: old ? (old.leader || '') : '',
      equipmentStatus: old ? old.equipmentStatus : [],
      waivedSeats: old ? (old.waivedSeats || []) : [],
      waiverLog: old ? (old.waiverLog || []) : []
    };
  });
  eq.groups = newGroups;
  syncEquipmentReturned(eq);
  writeEquipment(eq);
  return { ok: true, equipment: eq };
}

/** payload.attendance = { 組名: { 座號: true/false } } */
/**
 * 點名存檔——如果全班共用的計時「還沒開始」，存完點名就直接幫全部組別重新排一次
 * 隨機順序，不用老師再多按一次「重新排定順序」；組長不再是另外抽的，就是排序
 * 第一個的組員（前端顯示時直接從 order[0] 算，這裡不用另外存 leader）。
 * 如果計時「已經開始」了，就不重排，避免把正在進行的課堂計時歸零——這種情況下
 * 出席狀態的異動已經靠前端 presentOrderFor() 即時反映在畫面上，不需要重排。
 */
function handleEquipmentSetAttendance(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var eq = readEquipment();
  var att = payload.attendance || {};
  eq.groups.forEach(function (g) {
    var ga = att[g.name];
    if (!ga) return;
    g.members.forEach(function (m) { if (ga.hasOwnProperty(m.seat)) m.present = !!ga[m.seat]; });
  });
  if (!eq.startedAt) {
    eq.groups.forEach(function (g) {
      var all = g.members.slice();
      for (var i = all.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var t = all[i]; all[i] = all[j]; all[j] = t;
      }
      g.order = all.map(function (m) { return { seat: m.seat, name: m.name }; });
      // 重排順序等於重新開始一輪，之前的棄權紀錄不該繼續套用在新的順序上。
      g.waivedSeats = [];
      g.waiverLog = [];
    });
  }
  writeEquipment(eq);
  return { ok: true, equipment: eq };
}

/**
 * 幫「所有」組別重新洗牌排順序，並把全班的共用時間軸歸零（要另外按「開始」）。
 * 注意：這裡故意把全部組員（不管今天出席與否）都排進 order 裡——實際「今天算不算
 * 在轉圈裡」是前端 presentOrderFor() 每次即時看「目前的點名結果」動態算出來的，
 * 不是排這個順序的當下才決定。這樣之後老師改點名（不管是把人改缺席、還是遲到的人
 * 回來改成有來），畫面才會立刻反映最新狀態，不需要每次點名異動後都重新排一次順序。
 * 如果這裡排序時就先把缺席的人濾掉，遲到的人事後被重新點名「有來」也不會生效，
 * 因為他根本不在 order 名單裡，這正是先前回報「點名改了但輪值名單沒有跟著變」的原因。
 */
function handleEquipmentBuildOrder(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var eq = readEquipment();
  eq.groups.forEach(function (g) {
    var all = g.members.slice();
    for (var i = all.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var t = all[i]; all[i] = all[j]; all[j] = t;
    }
    g.order = all.map(function (m) { return { seat: m.seat, name: m.name }; });
    // 重新排順序等於是開新的一堂課，器材會重新借出去，歸還勾選也跟著整批重置；
    // 棄權紀錄也是針對這一輪的順序才有意義，一起清空。
    g.equipmentStatus = (eq.itemTypes || []).map(function () { return defaultEquipmentStatus(); });
    g.waivedSeats = [];
    g.waiverLog = [];
  });
  eq.startedAt = null;
  eq.pausedAt = null;
  eq.pausedTotalMs = 0;
  eq.elapsedMs=0;eq.resumedAt=null;eq.endedAt=null;
  writeEquipment(eq);
  return { ok: true, equipment: eq };
}

/**
 * 學生本人放棄這一輪的使用權利——故意不檢查教師密語，情境是學生自己臨時決定
 * 不需要用器材，跟評分/交作品一樣是課堂信任機制。放棄之後，他原本還沒用到的
 * 時間會平均分給還沒輪到的其他非棄權組員（equipGroupState 裡計算），不是
 * 讓時段直接消失，也不會動到已經輪過的人。同一人重複呼叫沒有副作用。
 */
function handleEquipmentWaive(payload) {
  var eq = readEquipment();
  var g = findGroup(eq, payload.groupName);
  var seat = String(payload.seat || '').trim();
  var name = String(payload.name || '').trim();
  if (!g || !seat) return { ok: false, error: 'not_found' };
  settleEquipmentTimer(eq,Date.now());
  g.waivedSeats = g.waivedSeats || [];
  if (g.waivedSeats.indexOf(seat) < 0) {
    g.waivedSeats.push(seat);
    g.waiverLog = g.waiverLog || [];
    g.waiverLog.push({ seat: seat, name: name, at: Date.now(), elapsedMs: equipmentElapsedMs(eq,Date.now()) });
  }
  writeEquipment(eq);
  return { ok: true, equipment: eq };
}



function upgradeEquipmentTimer(eq){
  if(eq.clockVersion===2) return eq;
  var now=Date.now();
  eq.totalMinutes=equipmentTotalMinutes(eq);
  eq.elapsedMs=equipmentElapsedMs(eq,now);
  eq.resumedAt=eq.startedAt&&!eq.pausedAt ? now : null;
  eq.clockVersion=2; eq.endedAt=null;
  eq.classSchedule=readConfig().classSchedule||defaultClassSchedule();
  return eq;
}
function settleEquipmentTimer(eq,now){
  upgradeEquipmentTimer(eq);
  eq.elapsedMs=equipmentElapsedMs(eq,now);
  eq.resumedAt=eq.startedAt&&!eq.pausedAt&&!eq.endedAt ? now : null;
}
function handleEquipmentStop(payload){
  if(!checkPasscode(payload)) return {ok:false,error:'unauthorized'};
  var eq=readEquipment(); settleEquipmentTimer(eq,Date.now());
  eq.endedAt=Date.now(); eq.resumedAt=null; eq.pausedAt=null;
  writeEquipment(eq); return {ok:true,equipment:eq};
}
function handleEquipmentResetTimer(payload){
  if(!checkPasscode(payload)) return {ok:false,error:'unauthorized'};
  var eq=readEquipment(); upgradeEquipmentTimer(eq);
  eq.startedAt=null;eq.resumedAt=null;eq.pausedAt=null;eq.endedAt=null;eq.elapsedMs=0;eq.pausedTotalMs=0;
  eq.groups.forEach(function(g){g.waivedSeats=[];g.waiverLog=[];});
  writeEquipment(eq);return {ok:true,equipment:eq};
}

function handleEquipmentSetDuration(payload) {
  if(!checkPasscode(payload)) return {ok:false,error:'unauthorized'};
  var minutes=Number(payload.minutes);
  if(!isFinite(minutes)||minutes<1||minutes>1440) return {ok:false,error:'invalid_duration'};
  var eq=readEquipment();settleEquipmentTimer(eq,Date.now());
  eq.totalMinutes=minutes;
  writeEquipment(eq);return {ok:true,equipment:eq};
}

/** 老師在「教師設定」裡自訂滑桿可以調整的分鐘數範圍；目前的 perPersonMinutes 如果超出新範圍就跟著夾回範圍內。 */
function handleEquipmentSetDurationRange(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var eq = readEquipment();
  var lo = Math.max(1, parseInt(payload.min, 10) || 1);
  var hi = Math.max(lo, parseInt(payload.max, 10) || lo);
  eq.durationMin = lo;
  eq.durationMax = hi;
  eq.perPersonMinutes = Math.max(lo, Math.min(hi, eq.perPersonMinutes));
  writeEquipment(eq);
  return { ok: true, equipment: eq };
}

/** 老師在「教師借用管理」編輯器材清單（每行一項，同款要借幾份就重複列幾行）。 */
function handleEquipmentSetItemTypes(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var eq = readEquipment();
  var items = (payload.itemTypes || []).map(function (s) { return String(s || '').trim(); }).filter(Boolean);
  eq.itemTypes = items.length ? items : DEFAULT_EQUIPMENT_ITEM_TYPES.slice();
  syncEquipmentReturned(eq);
  writeEquipment(eq);
  return { ok: true, equipment: eq };
}

/**
 * 學生點「已歸還」用的——故意不檢查教師密語，因為情境是下課時學生自己隨手勾，
 * 跟評分、交作品一樣是低風險課堂情境下的信任機制，不是給老師專用的管理動作。
 */
/**
 * 學生「送出」本組的器材回報——故意不檢查教師密語，情境是下課時學生自己隨手回報，
 * 跟評分、交作品一樣是低風險課堂情境下的信任機制。分段確認流程：
 * 這裡只負責記「誰、什麼時候回報了哪些項目」，要等老師另外呼叫
 * handleEquipmentConfirmReturned 確認過，才算真正完成歸還。
 * payload.statuses 是跟 itemTypes 對齊的 true/false 陣列（true=這次回報「已歸還」）；
 * 如果某一項從已回報被改回未回報，等於撤銷回報，連帶把老師原本的確認也一起清掉——
 * 東西沒真的還回來，不該留著一個過期的「已確認」。
 */
function handleEquipmentReportReturned(payload) {
  var eq = readEquipment();
  var g = findGroup(eq, payload.groupName);
  if (!g) return { ok: false, error: 'not_found' };
  var statuses = payload.statuses || [];
  var seat = String(payload.reporterSeat || '').trim();
  var name = String(payload.reporterName || '').trim();
  if (!seat || !name) return { ok: false, error: 'missing_identity' };
  g.equipmentStatus.forEach(function (st, i) {
    var wantReported = !!statuses[i];
    if (wantReported === st.reported) return;
    if (wantReported) {
      st.reported = true;
      st.reportedBySeat = seat;
      st.reportedByName = name;
      st.reportedAt = Date.now();
    } else {
      st.reported = false;
      st.reportedBySeat = '';
      st.reportedByName = '';
      st.reportedAt = null;
      st.confirmed = false;
    }
  });
  writeEquipment(eq);
  return { ok: true, equipment: eq };
}

/** 老師勾選收到之後按「確認」——只有已經被學生回報過的項目才能確認，順序反過來沒有意義。 */
function handleEquipmentConfirmReturned(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var eq = readEquipment();
  var g = findGroup(eq, payload.groupName);
  if (!g) return { ok: false, error: 'not_found' };
  var confirms = payload.confirms || [];
  g.equipmentStatus.forEach(function (st, i) {
    if (st.reported && confirms[i]) st.confirmed = true;
  });
  writeEquipment(eq);
  return { ok: true, equipment: eq };
}

/** 老師專用：只重置器材歸還狀態（全部打回「還沒還」），不動輪值順序跟計時進度。 */
function handleEquipmentResetReturned(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var eq = readEquipment();
  eq.groups.forEach(function (g) {
    g.equipmentStatus = (eq.itemTypes || []).map(function () { return defaultEquipmentStatus(); });
  });
  writeEquipment(eq);
  return { ok: true, equipment: eq };
}

/** 開始（或重新開始）全班共用的時間軸，從頭計時。 */
function handleEquipmentStart(payload) {
  if(!checkPasscode(payload)) return {ok:false,error:'unauthorized'};
  var eq=readEquipment();upgradeEquipmentTimer(eq);
  if(!eq.groups.some(function(g){return presentOrderForServer(g).length>0;})) return {ok:false,error:'no_order'};
  eq.startedAt=Date.now();eq.resumedAt=eq.startedAt;eq.pausedAt=null;eq.endedAt=null;eq.elapsedMs=0;eq.pausedTotalMs=0;
  eq.classSchedule=readConfig().classSchedule||defaultClassSchedule();
  eq.groups.forEach(function(g){g.waivedSeats=[];g.waiverLog=[];});
  writeEquipment(eq);return {ok:true,equipment:eq};
}

/** 暫停／繼續全班共用的時間軸（下課、臨時狀況都可以用這個）。 */
function handleEquipmentTogglePause(payload) {
  if(!checkPasscode(payload)) return {ok:false,error:'unauthorized'};
  var eq=readEquipment();settleEquipmentTimer(eq,Date.now());
  if(!eq.startedAt||eq.endedAt) return {ok:false,error:'not_started'};
  if(eq.pausedAt){eq.pausedAt=null;eq.resumedAt=Date.now();}
  else {eq.pausedAt=Date.now();eq.resumedAt=null;}
  writeEquipment(eq);return {ok:true,equipment:eq};
}

function readConfig(fresh) {
  return cachedData('cfg', readConfigUncached, fresh);
}

function readConfigUncached() {
  var cfg = getSheet('Config');
  if (cfg.getLastRow() < 2) return defaultConfig();
  var v = cfg.getRange(2, 1).getValue();
  try {
    var parsed = JSON.parse(v);
    if (parsed && parsed.criteria) {
      // 舊資料可能是加 revealMinReviewers／revealMinReviews／assignmentDescriptions 之前存的，補上預設值
      // （舊的 scoreRevealThreshold／namesRevealThreshold 已經不再使用，留在試算表裡也不會有影響）。
      if (parsed.revealMinReviewers == null) parsed.revealMinReviewers = 15;
      if (parsed.revealMinReviews == null) parsed.revealMinReviews = 5;
      if (parsed.assignmentDescriptions == null) parsed.assignmentDescriptions = {};
      if (!parsed.assignmentCriteria) parsed.assignmentCriteria = {};
      if (!parsed.criteriaArchive) parsed.criteriaArchive = {};
      if (!validClassSchedule(parsed.classSchedule)) parsed.classSchedule = defaultClassSchedule();
      return parsed;
    }
  } catch (e) {}
  return defaultConfig();
}

function writeConfig(obj) {
  var cfg = getSheet('Config');
  if (cfg.getLastRow() < 2) cfg.appendRow(['json']);
  cfg.getRange(2, 1).setValue(JSON.stringify(obj));
}

/**
 * 用來比對兩個「作業名稱」是否算同一份。
 * 這裡刻意不用單純的 === 比較字串，是因為 Google 試算表在寫入像
 * "01"、"02" 這種看起來像數字的文字時，會自動把它存成數字（前面的 0
 * 會不見），造成同一份作業在不同地方讀出來一次是字串"01"、一次是數字1，
 * 直接比較會判定成「不是同一份」，結果作業繳交狀況、抽件評分、教師總表
 * 篩選全部都對不起來。統一用這個函式正規化過後再比較就不會有這個問題，
 * 不需要去改試算表裡已經存進去的舊資料。
 */
function normAssignment(a) {
  var s = String(a == null ? '' : a).trim();
  if (s !== '' && !isNaN(Number(s))) return String(Number(s));
  return s;
}

/**
 * 判斷某一份作業目前要不要開放姓名／座號。
 * 新版設定是每份作業分開存（namesRevealedByAssignment），但相容舊版：
 * 舊版只有一個全班共用的 namesRevealed 開關，還沒切換過設定的試算表
 * 裡不會有 namesRevealedByAssignment 這個欄位，這種情況才退回用舊欄位判斷，
 * 避免老師原本已經開放的姓名，一升級就整個被關掉。
 */
function isNamesRevealedFor(cfg, assignment) {
  var map = cfg.namesRevealedByAssignment;
  if (map && typeof map === 'object') {
    var target = normAssignment(assignment);
    for (var k in map) {
      if (map.hasOwnProperty(k) && normAssignment(k) === target) return !!map[k];
    }
    return false;
  }
  return !!cfg.namesRevealed;
}

function readSubmissions(fresh) {
  return cachedData('subs', readSubmissionsUncached, fresh);
}

// ---- 座號格式統一 ----
// 學生可能輸入「 1 」「01」「１」（全形），這幾個都是同一個人。
// canonSeatServer：給「比對」用的標準形式（半形、去空白、純數字去掉前導零）；
// padSeat：給「顯示／存檔／檔名」用的標準形式（半形、去空白、純數字不足兩位補 0，例如 1 → 01）。
// 試算表裡舊資料的身分識別碼（seat::name）可能是舊格式，所以讀取時一律先轉成標準形式再比對，不用改舊資料。
function toHalfWidthServer(s) {
  return String(s == null ? '' : s)
    .replace(/[！-～]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xfee0); })
    .replace(/　/g, ' ');
}
function canonSeatServer(s) {
  var t = toHalfWidthServer(s).replace(/\s+/g, '').toLowerCase();
  return /^\d+$/.test(t) ? t.replace(/^0+(?=\d)/, '') : t;
}
function padSeat(s) {
  var t = toHalfWidthServer(s).replace(/\s+/g, '');
  return (/^\d+$/.test(t) && t.length < 2) ? '0' + t : t;
}
// ---- 老師測試帳號 ----
// 座號 99 是保留給老師的測試帳號：能用全部的學生端功能，但不屬於任何器材分組、不會被抽給同學評分、
// 作品不出現在作品牆、它評的分不算進任何分數或「評完」人數。沒有教師密語就不能用（見 authFailed）。
var TEST_SEAT = '99';
function isTestNorm(norm) { return seatOfNorm(norm) === TEST_SEAT; }
function dropTestSubs(subs) { return subs.filter(function (s) { return !isTestNorm(s.studentNorm); }); }
function dropTestReviews(revs) { return revs.filter(function (r) { return !isTestNorm(r.reviewerNorm); }); }

function canonNorm(n) {
  var s = String(n == null ? '' : n);
  var i = s.indexOf('::');
  if (i < 0) return s;
  return canonSeatServer(s.substring(0, i)) + '::' + s.substring(i + 2);
}

// ---- 學生密碼登入 ----
// 學生用班級名單登入，第一次登入自己設密碼（老師建議用身分證字號），之後新裝置要輸入密碼。
// 隱私設計：密碼（身分證字號）在學生手機上先轉成 SHA-256 亂碼才送出，這裡從頭到尾拿不到原文；
// 試算表裡存的是「亂碼 + 只存在 Apps Script 指令碼屬性裡的祕密金鑰」再算一次 HMAC 的結果，
// 就算有人拿到試算表，沒有那把金鑰也沒辦法離線猜身分證字號。
// 登入成功後發一個 token（座號 + 到期日 + 密碼設定時間，用同一把金鑰簽章），之後學生端動作都要帶，
// 老師重設密碼 = 刪掉那一列，舊 token 立刻全部失效。
var AUTH_ACTIONS = {
  getMySubmissions: 'studentNorm', addSubmission: 'studentNorm', deleteSubmission: 'studentNorm', downloadBundle: 'studentNorm',
  getNextWork: 'reviewerNorm', getMyReviewProgress: 'reviewerNorm', addReview: 'reviewerNorm'
};
var TOKEN_DAYS = 120;
var LOGIN_MAX_FAILS = 8;

function getPepper() {
  var props = PropertiesService.getScriptProperties();
  var v = props.getProperty('AUTH_PEPPER');
  if (!v) { v = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('AUTH_PEPPER', v); }
  return v;
}
function hmacB64(msg) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(msg, getPepper()));
}
function safeEqual(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var d = 0;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function seatOfNorm(norm) {
  var s = String(norm || '');
  var i = s.indexOf('::');
  return canonSeatServer(i < 0 ? s : s.substring(0, i));
}
function readAccounts(fresh) { return cachedData('acct', readAccountsUncached, fresh); }
function readAccountsUncached() {
  var sh = getSheet('Accounts');
  var out = {};
  if (sh.getLastRow() < 2) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    var seat = canonSeatServer(vals[i][0]);
    if (seat && vals[i][1]) out[seat] = { hash: String(vals[i][1]), setAt: Number(vals[i][2]) || 0 };
  }
  return out;
}
function makeToken(seat, setAt) {
  var exp = Date.now() + TOKEN_DAYS * 86400000;
  return exp + '.' + hmacB64('tok|' + clsNs() + seat + '|' + exp + '|' + setAt);
}
function verifyToken(token, seat, acct) {
  var t = String(token || '');
  var dot = t.indexOf('.');
  if (dot < 1) return false;
  var exp = Number(t.substring(0, dot));
  if (!exp || exp < Date.now()) return false;
  return safeEqual(t.substring(dot + 1), hmacB64('tok|' + clsNs() + seat + '|' + exp + '|' + acct.setAt));
}
/** 這個動作要不要擋下：座號已經設過密碼，就一定要帶有效 token（老師已解鎖教師專區的話例外，方便老師測試）。 */
function authFailed(action, payload) {
  var field = AUTH_ACTIONS[action];
  if (!field || !payload || !payload[field]) return false;
  var seat = seatOfNorm(payload[field]);
  var teacherOk = !!(payload.teacherPass && checkPasscode({ passcode: payload.teacherPass }));
  if (seat === TEST_SEAT) return !teacherOk; // 99 號老師測試帳號：一定要帶有效的教師密語，同學不能自己打 99 進來
  var acct = readAccounts()[seat];
  if (!acct) return false; // 還沒設定密碼的座號維持原本規則
  if (teacherOk) return false;
  return !verifyToken(payload.token, seat, acct);
}
/** 只驗證教師密語對不對（99 號測試帳號登入用），比 teacherFetch 輕很多。 */
function handleVerifyTeacher(payload) {
  return { ok: checkPasscode(payload) };
}
function handleAccountStatus(payload) {
  return { ok: true, exists: !!readAccounts()[canonSeatServer(payload.seat)] };
}
function handleLogin(payload) {
  var seat = canonSeatServer(payload.seat);
  var cache = CacheService.getScriptCache();
  var key = clsNs() + 'loginFail_' + seat;
  var fails = parseInt(cache.get(key) || '0', 10);
  if (fails >= LOGIN_MAX_FAILS) return { ok: false, error: 'locked' };
  var acct = readAccounts(true)[seat];
  if (!acct) return { ok: false, error: 'no_account' };
  if (!safeEqual(hmacB64('pw|' + clsNs() + seat + '|' + payload.pwHash), acct.hash)) {
    cache.put(key, String(fails + 1), 600);
    return { ok: false, error: 'bad_password', left: Math.max(0, LOGIN_MAX_FAILS - fails - 1) };
  }
  cache.remove(key);
  return { ok: true, token: makeToken(seat, acct.setAt) };
}
function handleRegister(payload) {
  var seat = canonSeatServer(payload.seat);
  var hash = String(payload.pwHash || '');
  if (!seat || !/^[0-9a-f]{64}$/.test(hash)) return { ok: false, error: 'bad_request' };
  if (seat === TEST_SEAT) return { ok: false, error: 'reserved' };
  // 座號一定要在班級名單（器材分組）裡，不能亂造帳號；名單還沒設定時不擋。
  var eq = readEquipment();
  var rosterSeats = [];
  (eq.groups || []).forEach(function (g) { (g.members || []).forEach(function (m) { rosterSeats.push(canonSeatServer(m.seat)); }); });
  if (rosterSeats.length && rosterSeats.indexOf(seat) < 0) return { ok: false, error: 'not_in_roster' };
  if (readAccounts(true)[seat]) return { ok: false, error: 'exists' };
  var sh = getSheet('Accounts');
  if (sh.getLastRow() === 0) sh.appendRow(['seat', 'pwHmac', 'setAt']);
  var setAt = Date.now();
  sh.appendRow([seat, hmacB64('pw|' + clsNs() + seat + '|' + hash), setAt]);
  return { ok: true, token: makeToken(seat, setAt) };
}
/** 老師專用：忘記密碼的同學，刪掉他的帳號，他下次登入就會重新設定；舊的登入全部失效。 */
function handleResetPassword(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var seat = canonSeatServer(payload.seat);
  var sh = getSheet('Accounts');
  var removed = false;
  for (var r = sh.getLastRow(); r >= 2; r--) {
    if (canonSeatServer(sh.getRange(r, 1).getValue()) === seat) { sh.deleteRow(r); removed = true; }
  }
  return { ok: true, removed: removed };
}

function readSubmissionsUncached() {
  var sh = getSheet('Submissions');
  var vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues() : [];
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (!r[0]) continue;
    out.push({ id: r[0], studentName: r[1], studentSeat: padSeat(r[2]), studentNorm: canonNorm(r[3]), workCode: r[4], title: r[5], driveLink: r[6], createdAt: r[7], assignment: r[8] || '', deviceId: r[9] || '', imageHash: r[10] || '' });
  }
  return out;
}

function readReviews(fresh) {
  return cachedData('revs', readReviewsUncached, fresh);
}

function readReviewsUncached() {
  var sh = getSheet('Reviews');
  var vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (!r[0]) continue;
    var scores = {};
    try { scores = JSON.parse(r[3]); } catch (e) {}
    out.push({ id: r[0], workId: r[1], reviewerNorm: canonNorm(r[2]), scores: scores, createdAt: r[4] });
  }
  return out;
}

function reviewCountFor(reviews, workId) {
  var n = 0;
  for (var i = 0; i < reviews.length; i++) if (reviews[i].workId === workId) n++;
  return n;
}
function authorKey(s) {
  return s.studentNorm + '|' + normAssignment(s.assignment);
}
/**
 * 同一個人在同一份作業，不管交了幾件（誤傳兩張一樣的、或是真的交了好幾件），
 * 每位評分者最多只會被分到其中一件：評過其中一件，同作者同作業的其他件就不會再抽給他。
 * 這樣任何一位作者佔用的「評分名額」，對每位評分者來說最多只有一個，
 * 重複上傳不會洗掉其他同學作品該被評到的次數。
 */
function candidatesFor(submissions, reviews, norm) {
  var reviewed = {};
  reviews.forEach(function (r) { if (r.reviewerNorm === norm) reviewed[r.workId] = 1; });
  var reviewedAuthors = {};
  submissions.forEach(function (s) { if (reviewed[s.id]) reviewedAuthors[authorKey(s)] = 1; });
  return submissions.filter(function (s) {
    return s.studentNorm !== norm && !reviewed[s.id] && !reviewedAuthors[authorKey(s)] && !isTestNorm(s.studentNorm);
  });
}
function hash32(str) {
  var h = 2166136261;
  for (var i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619); }
  return h >>> 0;
}
/**
 * 優先抽「目前被評次數最少」的作品，次數一樣的再用「評分者＋作品＋第幾件」算出來的固定雜湊值決定先後，
 * 不是每次重新擲骰子。這樣同一個人在同一個進度重新整理看到的會是同一件（不會越刷越換），
 * 而且伺服器可以事先算出「評完這件之後的下一件大概是哪件」，讓前端提前把圖載好。
 */
function pickNext(submissions, reviews, norm, step) {
  var cand = candidatesFor(submissions, reviews, norm);
  if (!cand.length) return null;
  var counts = {};
  reviews.forEach(function (r) { counts[r.workId] = (counts[r.workId] || 0) + 1; });
  cand.forEach(function (c) { c._n = counts[c.id] || 0; c._h = hash32(norm + '|' + c.id + '|' + (step || 0)); });
  cand.sort(function (a, b) { return (a._n - b._n) || (a._h - b._h); });
  return cand[0];
}

function criteriaForAssignment(cfg,assignment){
  var map=cfg.assignmentCriteria||{},wanted=normAssignment(assignment);
  var key=Object.keys(map).filter(function(k){return normAssignment(k)===wanted;})[0];
  return key!==undefined && Array.isArray(map[key]) && map[key].length ? map[key] : cfg.criteria;
}
function uniqueCriteria(lists){
  var seen={},out=[];
  lists.forEach(function(list){(list||[]).forEach(function(c){if(!seen[c.key]){seen[c.key]=true;out.push({key:c.key,label:c.label});}});});
  return out;
}
function criteriaForWork(cfg,work,reviews){
  var current=criteriaForAssignment(cfg,work.assignment);
  var old=(cfg.criteriaArchive||{})[String(work.assignment)]||cfg.criteria;
  var relevant=old.filter(function(c){return reviews.some(function(r){return r.workId===work.id && typeof r.scores[c.key]==='number';});});
  return uniqueCriteria([current,relevant]);
}
function validCriteria(list){
  if(!Array.isArray(list)||!list.length||list.length>50) return false;
  var seen={};
  return list.every(function(c){if(!c||typeof c.key!=='string'||!c.key||c.key==='__proto__'||seen[c.key]||typeof c.label!=='string'||!c.label.trim()) return false;seen[c.key]=true;return true;});
}

function statsFor(criteria, reviews, workId) {
  var subReviews = reviews.filter(function (r) { return r.workId === workId; });
  var perCrit = {};
  criteria.forEach(function (c) {
    var sum = 0, n = 0;
    subReviews.forEach(function (r) { if (typeof r.scores[c.key] === 'number') { sum += r.scores[c.key]; n++; } });
    perCrit[c.key] = n ? sum / n : null;
  });
  var overallSum = 0, overallN = 0;
  subReviews.forEach(function (r) {
    var vals = criteria.map(function (c) { return r.scores[c.key]; }).filter(function (v) { return typeof v === 'number'; });
    if (vals.length) { overallSum += vals.reduce(function (a, b) { return a + b; }, 0) / vals.length; overallN++; }
  });
  return { count: subReviews.length, perCrit: perCrit, overall: overallN ? overallSum / overallN : null };
}

function getUploadFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty(clsPropKey('UPLOAD_FOLDER_ID'));
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) { /* 資料夾被刪了，重新建一個 */ }
  }
  // 存圖已經搬到全域鎖外面，第一次同時有好幾個人上傳時可能同時走到這裡；
  // 用另一把鎖（使用者鎖，不是全域鎖）擋住，進來後再查一次，避免建出好幾個資料夾。
  var folderLock = LockService.getUserLock();
  folderLock.waitLock(20000);
  try {
    folderId = props.getProperty(clsPropKey('UPLOAD_FOLDER_ID'));
    if (folderId) {
      try { return DriveApp.getFolderById(folderId); } catch (e) { /* 還是找不到，往下重建 */ }
    }
    var folder = DriveApp.createFolder('匿名評圖牆 - 作品圖片' + (CURRENT_CLASS ? ' - ' + CURRENT_CLASS : ''));
    props.setProperty(clsPropKey('UPLOAD_FOLDER_ID'), folder.getId());
    try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return folder;
  } finally {
    folderLock.releaseLock();
  }
}

/**
 * 每張圖單獨 setSharing 要多打一次雲端硬碟 API（通常多花一秒左右），
 * 改成「資料夾本身」設一次知道連結的人可檢視，裡面的檔案會繼承。
 * 舊資料夾（之前建立、還沒設過分享）第一次上傳時補設一次。
 * 為了避免繼承沒生效導致圖片看不到，第一次上傳會驗證新檔案真的是公開的；
 * 驗證失敗就自動退回「每張圖單獨設分享」並記起來，之後都走舊做法。
 */
function ensureFileShared(file, folder) {
  var props = PropertiesService.getScriptProperties();
  var cache = CacheService.getScriptCache();
  var perFile = function () { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); };
  if (props.getProperty(clsPropKey('PER_FILE_SHARING')) === '1') { perFile(); return; }
  if (props.getProperty(clsPropKey('FOLDER_SHARED')) !== '1') {
    try {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      props.setProperty(clsPropKey('FOLDER_SHARED'), '1');
    } catch (e) { perFile(); props.setProperty(clsPropKey('PER_FILE_SHARING'), '1'); return; }
  }
  if (cache.get(clsNs() + 'folderShareVerified')) return;
  var inherited = false;
  try { inherited = (file.getSharingAccess() === DriveApp.Access.ANYONE_WITH_LINK); } catch (e) {}
  if (inherited) {
    cache.put(clsNs() + 'folderShareVerified', '1', 21600);
  } else {
    perFile();
    props.setProperty(clsPropKey('PER_FILE_SHARING'), '1');
  }
}

function sanitizeFilenamePart(s) {
  return String(s || '')
    .replace(/[\/\\:*?"<>|]/g, '')
    .replace(/\s+/g, '')
    .trim() || '未命名';
}

function saveUploadedImage(filename, base64, mime) {
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mime || 'image/jpeg', filename);
  var folder = getUploadFolder();
  var file = folder.createFile(blob);
  ensureFileShared(file, folder);
  // uc?export=view 常被 Google 擋下來改顯示病毒掃描警告頁，改用縮圖服務比較穩定。
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600';
}

/** 上傳到雲端硬碟的檔名：座號_姓名_作業名稱_第N個作品.jpg（沒有作業名稱就省略那一段）。 */
function uploadFilename(seat, studentName, assignment, n) {
  var parts = [sanitizeFilenamePart(padSeat(seat)), sanitizeFilenamePart(studentName)];
  if (String(assignment == null ? '' : assignment).trim()) parts.push(sanitizeFilenamePart(assignment));
  parts.push('第' + n + '個作品');
  return parts.join('_') + '.jpg';
}

/**
 * 老師專用：把「已經交過」的圖檔，依照現在的檔名格式一次改名。
 * 只動「本系統自己建立的上傳資料夾」裡的檔案——學生貼外部連結、或貼到老師其他雲端硬碟檔案的連結都不會被碰到。
 * 「第N個作品」是這位學生的第幾件（照試算表裡的先後順序算）。
 * 這個動作只改雲端硬碟檔名，不寫試算表，圖片網址是用檔案 ID，改名不影響作品牆顯示；
 * 所以不拿全域鎖（以免整批改名好幾分鐘期間，全班都交不了作品）。
 * 一次最多處理約 4 分鐘，沒做完會回傳 remaining，老師再按一次就接著做（已經改對的會直接略過）。
 */
function handleRenameUploads(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var started = Date.now();
  var folderId = PropertiesService.getScriptProperties().getProperty(clsPropKey('UPLOAD_FOLDER_ID'));
  if (!folderId) return { ok: true, renamed: 0, unchanged: 0, skipped: 0, failed: 0, remaining: 0 };
  var folder;
  try { folder = DriveApp.getFolderById(folderId); } catch (e) { return { ok: true, renamed: 0, unchanged: 0, skipped: 0, failed: 0, remaining: 0 }; }
  var filesById = {};
  var it = folder.getFiles();
  while (it.hasNext()) { var f = it.next(); filesById[f.getId()] = f; }

  var subs = readSubmissions(true);
  var seqByStudent = {};
  var renamed = 0, unchanged = 0, skipped = 0, failed = 0, remaining = 0;
  for (var i = 0; i < subs.length; i++) {
    var s = subs[i];
    seqByStudent[s.studentNorm] = (seqByStudent[s.studentNorm] || 0) + 1;
    var m = String(s.driveLink || '').match(/[?&]id=([a-zA-Z0-9_-]+)/);
    var file = m ? filesById[m[1]] : null;
    if (!file) { skipped++; continue; }
    var want = uploadFilename(s.studentSeat, s.studentName, s.assignment, seqByStudent[s.studentNorm]);
    try {
      if (file.getName() === want) { unchanged++; continue; }
      if (Date.now() - started > 240000) { remaining++; continue; }
      file.setName(want);
      renamed++;
    } catch (e) { failed++; }
  }
  return { ok: true, renamed: renamed, unchanged: unchanged, skipped: skipped, failed: failed, remaining: remaining };
}

function addSubmission(payload) {
  var sh = getSheet('Submissions');
  var lastRow = sh.getLastRow();
  var existing = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 11).getValues() : []; // id, studentName, studentSeat, studentNorm, workCode, title, driveLink, createdAt, assignment, deviceId, imageHash
  for (var i = 0; i < existing.length; i++) {
    if (existing[i][0] === payload.id) return { code: existing[i][4], duplicate: false, reused: true }; // 已經交過了（安全重試），回傳原本的編號
  }
  var used = {};
  var mineCount = 0;
  var wantAssignment = normAssignment(payload.assignment);
  for (var j = 0; j < existing.length; j++) {
    used[existing[j][4]] = true;
    if (canonNorm(existing[j][3]) === payload.studentNorm) {
      mineCount++;
      // 同一個人在同一份作業重複交一模一樣的圖（或同一個網址），不再多存一件。
      if (payload.imageHash && existing[j][10] === payload.imageHash && normAssignment(existing[j][8]) === wantAssignment) {
        return { code: existing[j][4], duplicate: true, reused: true };
      }
    }
  }
  var code;
  do {
    code = 'P' + (1000 + Math.floor(Math.random() * 9000));
  } while (used[code]);
  var link = payload.driveLink || '';
  var seat = padSeat(payload.studentSeat);
  if (payload.savedLink) {
    link = payload.savedLink; // 圖已經在鎖外面存好了（見 prepareSubmissionImage），鎖裡面只剩寫試算表這一步
  } else if (payload.imageBase64) {
    link = saveUploadedImage(uploadFilename(seat, payload.studentName, payload.assignment, mineCount + 1), payload.imageBase64, payload.imageMime);
  }
  sh.appendRow([payload.id, payload.studentName, seat, payload.studentNorm, code, payload.title, link, payload.createdAt, payload.assignment || '', payload.deviceId || '', payload.imageHash || '']);
  // 作業名稱那一欄要強制設成「純文字」格式再寫一次，不然像 "01" 這種
  // 看起來像數字的名稱，試算表會自動存成數字、把前面的 0 吃掉
  // （細節見 normAssignment 的註解）。這裡對舊資料不會有影響，只補救之後新交的作品。
  var assignRange = sh.getRange(sh.getLastRow(), 9);
  assignRange.setNumberFormat('@');
  assignRange.setValue(payload.assignment || '');
  return { code: code, duplicate: false };
}

/**
 * 回傳 true（刪掉了）、false（找不到）、'wrong_device'（不是原本交作品的那台裝置）。
 * 登入只要輸入座號就行，任何人都能「變成」別人，所以刪除多一道鎖：必須是交作品的同一台裝置
 * （裝置代碼是隨機產生、只存在那支手機／電腦、畫面上看不到，學生拿不到別人的）。
 * 沒有裝置代碼的舊作品（這個欄位加入之前交的）維持原本規則，不擋。
 */
function deleteSubmission(payload) {
  var sh = getSheet('Submissions');
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  var vals = sh.getRange(2, 1, lastRow - 1, 10).getValues();
  var rowIndex = -1, driveLink = '';
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === payload.id && canonNorm(vals[i][3]) === payload.studentNorm) {
      var rowDevice = String(vals[i][9] || '');
      if (rowDevice && String(payload.deviceId || '') !== rowDevice) return 'wrong_device';
      // 被老師退回的作品要留著當紀錄（扣分、退回原因都掛在它上面），學生不能自己刪掉，要另外再交一件。
      if (returnedInfo(readConfig(), vals[i][4])) return 'returned';
      rowIndex = i; driveLink = vals[i][6]; break;
    }
  }
  if (rowIndex === -1) return false;
  sh.deleteRow(rowIndex + 2);

  // 把這件作品收到的評分一併刪掉，避免總表算到已經不存在的作品
  var revSh = getSheet('Reviews');
  var revLast = revSh.getLastRow();
  if (revLast > 1) {
    var revVals = revSh.getRange(2, 1, revLast - 1, 2).getValues();
    for (var j = revVals.length - 1; j >= 0; j--) {
      if (revVals[j][1] === payload.id) revSh.deleteRow(j + 2);
    }
  }

  // 盡量把雲端硬碟裡對應的圖片也丟進垃圾桶，失敗也不影響刪除結果
  try {
    var m = String(driveLink).match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (m) DriveApp.getFileById(m[1]).setTrashed(true);
  } catch (e) {}

  return true;
}

/**
 * 老師測試完畢、正式開始使用前的「一鍵還原」：只清掉測試期間累積出來的
 * 「資料」——作品、評分、雲端硬碟裡對應的圖片、器材輪值目前計時到哪裡／
 * 今天點名結果——不會動到老師已經設定好的「設定」，像評分題目、作業清單、
 * 密語、姓名／分數公開規則、器材分組名單與顏色、每人使用時間，這些都是
 * 正式上線也要繼續用的設定，保留下來。
 * 因為這個動作會把所有學生的作品、評分永久刪除、無法復原，前端會要求
 * 老師另外打一次密語做二次確認，這裡也再檢查一次 confirmPasscode，
 * 光靠 checkPasscode 那一次驗證還不夠。
 */
/**
 * 還原前先把整份試算表（含程式）複製一份到雲端硬碟，放在原檔同一個資料夾，檔名「【還原前備份】班級名稱 時間」。
 * 備份失敗就不繼續刪（見 resetAllData）。備份裡的圖片網址指向原本的圖片；還原會把圖片丟進垃圾桶（保留 30 天），
 * 到垃圾桶還原圖片後，備份裡的作品就能正常顯示。
 */
function backupSpreadsheet(label) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  var file = DriveApp.getFileById(ss.getId());
  var name = '【還原前備份】' + label + ' ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  var parents = file.getParents();
  var copy = parents.hasNext() ? file.makeCopy(name, parents.next()) : file.makeCopy(name);
  return { name: name, url: copy.getUrl() };
}

function resetAllData(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  if (!payload.confirmPasscode || payload.confirmPasscode !== payload.passcode) {
    return { ok: false, error: 'confirm_mismatch' };
  }
  // 防呆：一定要先設定班級名稱，而且要輸入「這個資料庫」的班級名稱才會執行。
  // 比對是在伺服器做的——網頁萬一連到別班的資料庫（例如複製試算表時網址貼錯），輸入的班級名稱會對不上，就不會動到那一班。
  var guardCfg = readConfig(true);
  var guardName = String(guardCfg.className || '').trim();
  if (!guardName) return { ok: false, error: 'class_name_required' };
  if (String(payload.confirmClassName || '').trim() !== guardName) return { ok: false, error: 'class_mismatch', className: guardName };
  // 刪之前先備份整份試算表；備份不成功就不刪
  var backup;
  try { backup = backupSpreadsheet(guardName); } catch (be) { return { ok: false, error: 'backup_failed', detail: String(be && be.message || be) }; }

  var subsSheet = getSheet('Submissions');
  var subsLast = subsSheet.getLastRow();
  if (subsLast > 1) {
    // 只把「這個班自己的上傳資料夾」裡的圖丟進垃圾桶。
    // 複製試算表開新班級時，複製來的作品資料還指著原本那一班的圖片；如果不檢查就照單全收，
    // 在新班按「一鍵還原」會把原本那一班的圖全部刪掉。新班的資料夾（UPLOAD_FOLDER_ID）是各班自己的，
    // 複製試算表不會帶過來，所以新班這裡會是空的，什麼圖都不會動到。
    var ownIds = {};
    try {
      var ownFolderId = PropertiesService.getScriptProperties().getProperty(clsPropKey('UPLOAD_FOLDER_ID'));
      if (ownFolderId) {
        var it = DriveApp.getFolderById(ownFolderId).getFiles();
        while (it.hasNext()) ownIds[it.next().getId()] = true;
      }
    } catch (e) {}
    var subsVals = subsSheet.getRange(2, 1, subsLast - 1, 9).getValues();
    subsVals.forEach(function (row) {
      try {
        var m = String(row[6] || '').match(/[?&]id=([a-zA-Z0-9_-]+)/);
        if (m && ownIds[m[1]]) DriveApp.getFileById(m[1]).setTrashed(true);
      } catch (e) {}
    });
    subsSheet.deleteRows(2, subsLast - 1);
  }

  var revSheet = getSheet('Reviews');
  var revLast = revSheet.getLastRow();
  if (revLast > 1) revSheet.deleteRows(2, revLast - 1);

  var eq = readEquipment();
  eq.startedAt = null;
  eq.pausedAt = null;
  eq.pausedTotalMs = 0;
  eq.elapsedMs=0;eq.resumedAt=null;eq.endedAt=null;
  (eq.groups || []).forEach(function (g) {
    (g.members || []).forEach(function (m) { m.present = false; });
  });
  writeEquipment(eq);

  // 作品都清掉了，掛在作品編號上的老師加分和退回標記也一併清掉（「退回紀錄」分頁是歷史紀錄，不動）。
  var cfgAfterReset = readConfig(true);
  cfgAfterReset.bonusByWork = {};
  cfgAfterReset.returnedByWork = {};
  cfgAfterReset.flaggedReviewers = {};
  cfgAfterReset.awardsByWork = {};
  cfgAfterReset.attendanceBonus = {};
  cfgAfterReset.bonusNotes = {};
  writeConfig(cfgAfterReset);

  // 回到「全新班級」的狀態：同學登入密碼，加上各種紀錄分頁和匯出報表分頁，全部清空。
  // 複製試算表開新班級時，這些會連舊班級的內容一起被複製過來——密碼是用原班的金鑰加密的，
  // 新班用不了，還會讓同座號的新同學以為「已經設過密碼」登入不了；紀錄和報表則是原班同學的姓名和分數。
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Accounts', '加扣分紀錄', '評分異常紀錄', '榮譽榜加分', '點名加分', '退回紀錄'].forEach(function (name) {
    var sh = ss.getSheetByName(clsPrefix() + name);
    if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1); // 留下標題列
  });
  ['評分未達標名單', '作業繳交狀況', '評分總表', '加分彙整', '評分進度', '器材輪借報表', '器材放棄紀錄'].forEach(function (name) {
    var sh = ss.getSheetByName(clsPrefix() + name);
    if (sh) sh.clear(); // 這些是「匯出」時重新產生的報表，沒有資料要保留
  });

  return { ok: true, backup: backup };
}

function upsertReview(payload) {
  var sh = getSheet('Reviews');
  var lastRow = sh.getLastRow();
  var ids = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues() : [];
  var rowVals = [payload.id, payload.workId, payload.reviewerNorm, JSON.stringify(payload.scores), payload.createdAt];
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === payload.id) {
      sh.getRange(i + 2, 1, 1, rowVals.length).setValues([rowVals]);
      return;
    }
  }
  sh.appendRow(rowVals);
}

function checkPasscode(payload) {
  var cfg = readConfig();
  return !!(payload && payload.passcode && payload.passcode === cfg.teacherPasscode);
}

// ---- action handlers ----

function handleGetMySubmissions(payload) {
  var norm = payload.studentNorm;
  var cfg = readConfig();
  var mine = readSubmissions().filter(function (s) { return s.studentNorm === norm; })
    .map(function (s) { return { id: s.id, workCode: s.workCode, title: s.title, driveLink: s.driveLink, assignment: s.assignment, imageHash: s.imageHash, returned: returnedInfo(cfg, s.workCode) }; });
  return { submissions: mine };
}

function handleAddSubmission(payload) {
  // 交作品只收直接上傳的圖片，不再接受貼外部連結（舊資料裡已經存在的連結不受影響）。
  if (!payload.imageBase64 && !payload.savedLink && !payload.skipImage) return { ok: false, error: 'image_required' };
  var res = addSubmission(payload);
  return { workCode: res.code, duplicate: res.duplicate, reused: !!res.reused };
}

function handleDeleteSubmission(payload) {
  var res = deleteSubmission(payload);
  if (res === 'wrong_device') return { ok: false, error: 'wrong_device' };
  if (res === 'returned') return { ok: false, error: 'returned' };
  return { ok: res };
}

/**
 * 評分現在依「作業」分開計算：payload.assignment 指定要評哪一份作業，
 * 候選作品、已評份數、達標與否都只看這份作業底下的資料，不會跟其他
 * 作業的作品或評分紀錄混在一起。
 */
function handleGetNextWork(payload) {
  var norm = payload.reviewerNorm;
  var assignment = payload.assignment || '';
  var cfg = readConfig();
  var allSubmissions = readSubmissions();
  var submissions = assignment ? allSubmissions.filter(function (s) { return normAssignment(s.assignment) === normAssignment(assignment); }) : allSubmissions;
  var reviews = readReviews();
  var idsInScope = {};
  submissions.forEach(function (s) { idsInScope[s.id] = true; });
  var myDone = reviews.filter(function (r) { return r.reviewerNorm === norm && idsInScope[r.workId]; }).length;
  var required = cfg.reviewsPerStudent;
  var goal = Math.max(required, payload.extraTarget || 0);
  if (myDone >= goal) {
    var leftover = candidatesFor(dropReturned(cfg, submissions), reviews, norm).length;
    return { done: true, myDone: myDone, required: required, leftover: leftover, criteria: criteriaForAssignment(cfg,assignment) };
  }
  var pool = dropReturned(cfg, submissions); // 被退回的作品不再抽給同學評
  var work = pickNext(pool, reviews, norm, myDone);
  if (!work) return { done: true, myDone: myDone, required: required, leftover: 0, criteria: criteriaForAssignment(cfg,assignment) };
  // 順便算出「評完這件之後」大概會抽到哪一件，只給前端提前把圖載好（用固定雜湊決定先後，所以通常會一致）。
  var upcoming = null;
  if (myDone + 1 < goal) {
    var afterThis = reviews.concat([{ workId: work.id, reviewerNorm: norm, scores: {} }]);
    var nextWork = pickNext(pool, afterThis, norm, myDone + 1);
    if (nextWork) upcoming = { id: nextWork.id, driveLink: nextWork.driveLink };
  }
  return {
    done: false, myDone: myDone, required: required, criteria: criteriaForAssignment(cfg,work.assignment),
    work: { id: work.id, workCode: work.workCode, title: work.title, driveLink: work.driveLink },
    upcoming: upcoming
  };
}

/**
 * 「我要評分」選作業畫面用的：這個人（不用密語，只認 reviewerNorm）在每一份
 * 作業各評了幾件，讓學生自己選作業之前就能看到哪幾份還沒評到規定的件數。
 */
function handleGetMyReviewProgress(payload) {
  var norm = payload.reviewerNorm;
  var cfg = readConfig();
  var submissions = readSubmissions();
  var reviews = readReviews();
  var progress = {};
  (cfg.assignments || []).forEach(function (a) {
    var idsInScope = {};
    submissions.forEach(function (s) { if (normAssignment(s.assignment) === normAssignment(a)) idsInScope[s.id] = true; });
    progress[a] = reviews.filter(function (r) { return r.reviewerNorm === norm && idsInScope[r.workId]; }).length;
  });
  return { ok: true, assignments: cfg.assignments || [], required: cfg.reviewsPerStudent, progress: progress };
}

function handleAddReview(payload) {
  var work=readSubmissions().filter(function(s){return s.id===payload.workId;})[0];
  if(!work) return {ok:false,error:'no_work'};
  var criteria=criteriaForAssignment(readConfig(),work.assignment),scores=payload.scores||{};
  if(!criteria.every(function(c){return typeof scores[c.key]==='number'&&isFinite(scores[c.key])&&scores[c.key]>=1&&scores[c.key]<=5;})) return {ok:false,error:'criteria_changed'};
  payload.scores={}; criteria.forEach(function(c){payload.scores[c.key]=scores[c.key];});
  upsertReview(payload); return {ok:true};
}

/** 老師加分只認 0~3 的整數，其他一律當 0。 */
var RETURN_PENALTY = 0.1;
/**
 * 所有加分、扣分（作品加分、退回、榮譽榜加分、點名加分，包含取消／撤回）都會在試算表「加扣分紀錄」分頁多一列總帳：
 * 時間、類型、座號、姓名、作業、作品編號、分數變動（加分為正、扣分為負）、原因。
 * 各功能自己的分頁（退回紀錄、榮譽榜加分、點名加分）另外保留。entries 每一項是 { type, seat, name, assignment, workCode, delta, reason }。
 */
function logScoreChanges(entries) {
  if (!entries || !entries.length) return;
  var log = getSheet('加扣分紀錄');
  if (log.getLastRow() === 0) log.appendRow(['時間', '類型', '座號', '姓名', '作業', '作品編號', '分數變動', '原因']);
  var start = log.getLastRow() + 1;
  var range = log.getRange(start, 1, entries.length, 8);
  range.setNumberFormat('@'); // 座號、作業名稱像 "01" 的話不要被試算表吃掉前面的 0
  log.getRange(start, 1, entries.length, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  log.getRange(start, 7, entries.length, 1).setNumberFormat('+0.0##;-0.0##;0');
  var now = new Date();
  range.setValues(entries.map(function (e) {
    return [now, String(e.type || ''), String(e.seat || ''), String(e.name || ''), String(e.assignment || ''), String(e.workCode || ''), e.delta, String(e.reason || '')];
  }));
}
function submissionByCode(workCode) {
  var found = null;
  readSubmissions().forEach(function (s) { if (String(s.workCode) === String(workCode)) found = s; });
  return found;
}
var ATTENDANCE_BONUS = 0.5; // 點名區按一次「加分」加的總分
var AWARD_BONUS = 0.5; // 榮譽榜（每份作業前三名＋人氣獎）每位得獎作者額外加的總分
function awardOf(cfg, workCode) {
  var a = (cfg.awardsByWork || {})[String(workCode)];
  return a ? (Number(a.bonus) || 0) : 0;
}
function awardKindOf(cfg, workCode) {
  var a = (cfg.awardsByWork || {})[String(workCode)];
  return a ? String(a.kind || '') : '';
}
function returnedInfo(cfg, workCode) {
  var r = (cfg.returnedByWork || {})[String(workCode)];
  return r ? { reason: String(r.reason || ''), at: r.at || 0 } : null;
}
/** 把退回的作品從清單裡拿掉（評分抽件、公開作品牆用；老師的總表不拿掉，要看得到）。 */
function dropReturned(cfg, subs) {
  var map = cfg.returnedByWork || {};
  return subs.filter(function (s) { return !map[String(s.workCode)]; });
}

function bonusFor(cfg, workCode) {
  var n = parseInt((cfg.bonusByWork || {})[String(workCode)], 10);
  if (isNaN(n) || n < 0) return 0;
  return Math.min(3, n);
}

function buildRows(cfg, submissions, reviews) {
  var rows = submissions.map(function (s) {
    var workCriteria=criteriaForWork(cfg,s,reviews);
    var st = statsFor(workCriteria, reviews, s.id);
    // 老師加分直接加在總平均上，最高 5 分；還沒有任何評分的作品沒有總平均，就不套用（雷達圖各題的平均不受影響）。
    var bonus = bonusFor(cfg, s.workCode);
    var overall = (st.overall != null && bonus > 0) ? Math.min(5, st.overall + bonus) : st.overall;
    // 退回的作品：總平均再扣 0.1（最低 0 分），退回原因和時間另外記在「退回紀錄」分頁。
    var returned = returnedInfo(cfg, s.workCode);
    if (overall != null && returned) overall = Math.max(0, Math.round((overall - RETURN_PENALTY) * 1e6) / 1e6);
    return { workCode: s.workCode, studentName: s.studentName, studentSeat: s.studentSeat, title: s.title, assignment: s.assignment, driveLink: s.driveLink, deviceId: s.deviceId || '', criteria: workCriteria, perCrit: st.perCrit, overall: overall, baseOverall: st.overall, bonus: bonus, bonusNote: bonus > 0 ? ((cfg.bonusNotes || {})[String(s.workCode)] || null) : null, returned: returned, award: awardOf(cfg, s.workCode), awardKind: awardKindOf(cfg, s.workCode), count: st.count };
  });
  rows.sort(function (a, b) {
    var av = a.overall == null ? -1 : a.overall, bv = b.overall == null ? -1 : b.overall;
    return bv - av;
  });
  return rows;
}

/**
 * 公開的作品牆資料：不需要密語，同學跟老師都能看到全部人的作品跟分數。
 * 姓名／座號是否附上是「每份作業分開判斷」（isNamesRevealedFor），老師可以
 * 只開放某幾份作業的姓名，其他還沒開放的作業姓名還是會被拿掉。
 * 分數（含雷達圖）跟姓名的自動開放，現在改成看「全班的評分參與度」而不是單件作品被評幾次：
 * 某一份作業，只要有 cfg.revealMinReviewers 位同學、各自在這份作業評完至少 cfg.revealMinReviews 件，
 * 這份作業就「解鎖」——分數和姓名一起開放。人夠多的時候，統計上已經很難從分數反推是誰打的。
 * 老師仍然可以用 isNamesRevealedFor 手動提早開放某份作業的姓名（分數只走解鎖這條路）。
 * 未解鎖時 perCrit／overall 一律回傳 null（count 本身還是會回傳）。
 * 這些規則都是伺服器端強制執行的，不是前端畫面上的裝飾而已。
 */
function handleGallery(payload) {
  return cachedData('gallery', computeGallery);
}

/**
 * 首頁「榮譽榜」：每份作業的前三名作品作者＋人氣獎一名。
 * 資料直接取自作品牆那份已經套用解鎖規則的結果（computeGallery），所以規則完全一致：
 * 一份作業要等「評完的同學」達到門檻（解鎖）才會公布得獎者，還沒解鎖的只回傳進度（讓首頁顯示「再 N 位同學就揭曉」），
 * 不會把還沒公開的分數或姓名漏出去。被退回的作品、老師測試帳號的作品本來就不在作品牆資料裡。
 * 名次用「同分同名次」：同分的人並列，所以前三名偶爾會超過三個人。
 * 人氣獎＝「最吸引目光」：取題目裡「吸引力」那一題（找不到就用第一題）平均最高的作品，
 * 而且得獎者要是前三名以外的人，讓更多人有機會被看見；前三名以外沒人的話就不頒。
 */
function popularityCriterion(criteria) {
  var list = criteria || [];
  for (var i = 0; i < list.length; i++) if (/吸引|人氣|喜歡|喜愛/.test(String(list[i].label))) return list[i];
  return list[0] || null;
}

function handleHonors() {
  return cachedData('honors', computeHonors);
}

function computeHonors() {
  var g = cachedData('gallery', computeGallery);

  var items = [];
  (g.assignments || []).forEach(function (a) {
    var rows = g.rows.filter(function (r) { return normAssignment(r.assignment) === normAssignment(a); });
    if (!rows.length) return;
    var pop=popularityCriterion(criteriaForAssignment(readConfig(),a));
    var unlocked = rows.some(function (r) { return r.unlocked; });
    var item = {
      assignment: a, unlocked: unlocked, submissions: rows.length,
      have: (g.revealProgress || {})[normAssignment(a)] || 0, need: g.revealMinReviewers
    };
    if (unlocked) {
      var scored = rows.filter(function (r) { return r.overall != null; }).map(function (r) {
        return { r: r, score: Math.round(r.overall * 100) / 100 };
      }).sort(function (x, y) { return y.score - x.score; });
      scored.forEach(function (x) { x.rank = 1 + scored.filter(function (y) { return y.score > x.score; }).length; });
      var top = scored.filter(function (x) { return x.rank <= 3; }).slice(0, 6);
      var topSeats = {};
      top.forEach(function (x) { topSeats[x.r.studentSeat] = true; });
      var awards = readConfig().awardsByWork || {};
      item.top = top.map(function (x) {
        return { rank: x.rank, seat: x.r.studentSeat, name: x.r.studentName, score: x.score, driveLink: x.r.driveLink, workCode: x.r.workCode, awarded: !!awards[String(x.r.workCode)] };
      });
      item.popular = null;
      if (pop) {
        var best = null;
        scored.forEach(function (x) {
          if (topSeats[x.r.studentSeat]) return;
          var v = x.r.perCrit ? x.r.perCrit[pop.key] : null;
          if (v == null) return;
          if (!best || v > best.v) best = { x: x, v: v };
        });
        if (best) item.popular = { seat: best.x.r.studentSeat, name: best.x.r.studentName, score: Math.round(best.v * 100) / 100, label: pop.label, driveLink: best.x.r.driveLink, workCode: best.x.r.workCode, awarded: !!awards[String(best.x.r.workCode)] };
      }
    }
    items.push(item);
  });
  return { ok: true, items: items, revealMinReviews: g.revealMinReviews };
}

function qualifiedReviewersByAssignment(submissions, reviews, minReviews) {
  reviews = dropTestReviews(reviews);
  var assignByWork = {};
  submissions.forEach(function (s) { assignByWork[s.id] = normAssignment(s.assignment); });
  var counts = {};
  reviews.forEach(function (r) {
    var a = assignByWork[r.workId];
    if (a === undefined) return;
    counts[a] = counts[a] || {};
    counts[a][r.reviewerNorm] = (counts[a][r.reviewerNorm] || 0) + 1;
  });
  var out = {};
  for (var a in counts) {
    var n = 0;
    for (var who in counts[a]) if (counts[a][who] >= minReviews) n++;
    out[a] = n;
  }
  return out;
}

function computeGallery() {
  var cfg = readConfig();
  // 老師的測試帳號（99 號）交的作品和評的分都不給同學看、也不算進解鎖人數
  var submissions = dropTestSubs(readSubmissions());
  var reviews = dropTestReviews(readReviews());
  // 被老師退回的作品不再公開（也不算進解鎖人數以外的任何顯示），只留在教師總表。
  var rows = buildRows(cfg, dropReturned(cfg, submissions), reviews);
  var minReviewers = (cfg.revealMinReviewers == null) ? 15 : cfg.revealMinReviewers;
  var minReviews = (cfg.revealMinReviews == null) ? 5 : cfg.revealMinReviews;
  var qualified = qualifiedReviewersByAssignment(submissions, reviews, minReviews);
  rows = rows.map(function (r) {
    var unlocked = (qualified[normAssignment(r.assignment)] || 0) >= minReviewers;
    var namesOk = isNamesRevealedFor(cfg, r.assignment) || unlocked;
    var perCrit = r.perCrit;
    var overall = r.overall;
    if (!unlocked) {
      perCrit = {};
      for (var k in r.perCrit) perCrit[k] = null;
      overall = null;
    }
    return {
      workCode: r.workCode,
      studentName: namesOk ? r.studentName : null,
      studentSeat: namesOk ? r.studentSeat : null,
      title: r.title,
      assignment: r.assignment,
      driveLink: r.driveLink,
      criteria: r.criteria,
      perCrit: perCrit,
      overall: overall,
      bonus: unlocked ? r.bonus : 0,
      count: r.count,
      unlocked: unlocked
    };
  });
  return {
    rows: rows, criteria: uniqueCriteria(rows.map(function(r){return r.criteria;})), assignmentCriteria: cfg.assignmentCriteria || {}, assignments: cfg.assignments || [],
    revealMinReviewers: minReviewers, revealMinReviews: minReviews, revealProgress: qualified
  };
}

/**
 * 評分進度現在依作業分開算：reviewersByAssignment[作業名稱] 是那份作業
 * 底下每個人交了什麼、評了幾件、有沒有達標，彼此獨立計算。
 */
function buildReviewersByAssignment(cfg, submissions, reviews) {
  submissions = dropTestSubs(submissions);
  reviews = dropTestReviews(reviews);
  var result = {};
  // 只評分、還沒交這份作業的人：姓名座號優先從他「別份作業」交過的作品找；全都沒交過，
  // 就從身分識別碼（座號::姓名）拆出來，不要把「19::劉子涵」整串當成姓名、座號留空。
  var anyName = {}, anySeat = {};
  submissions.forEach(function (s) { anyName[s.studentNorm] = s.studentName; anySeat[s.studentNorm] = s.studentSeat; });
  function nameFromNorm(norm) { var i = String(norm).indexOf('::'); return i < 0 ? String(norm) : String(norm).substring(i + 2); }
  (cfg.assignments || []).forEach(function (a) {
    var subsForA = submissions.filter(function (s) { return normAssignment(s.assignment) === normAssignment(a); });
    var idsForA = {};
    subsForA.forEach(function (s) { idsForA[s.id] = true; });
    var nameByNorm = {}, seatByNorm = {}, codesByNorm = {};
    subsForA.forEach(function (s) {
      nameByNorm[s.studentNorm] = s.studentName;
      seatByNorm[s.studentNorm] = s.studentSeat;
      codesByNorm[s.studentNorm] = (codesByNorm[s.studentNorm] || []).concat([s.workCode]);
    });
    var reviewCountByNorm = {};
    reviews.forEach(function (r) { if (idsForA[r.workId]) reviewCountByNorm[r.reviewerNorm] = (reviewCountByNorm[r.reviewerNorm] || 0) + 1; });
    var allNorms = {};
    Object.keys(nameByNorm).forEach(function (n) { allNorms[n] = 1; });
    Object.keys(reviewCountByNorm).forEach(function (n) { allNorms[n] = 1; });
    var reviewers = Object.keys(allNorms).map(function (norm) {
      var count = reviewCountByNorm[norm] || 0;
      return {
        name: nameByNorm[norm] || anyName[norm] || nameFromNorm(norm),
        seat: seatByNorm[norm] || anySeat[norm] || padSeat(seatOfNorm(norm)),
        submitted: !!nameByNorm[norm],
        workCodes: (codesByNorm[norm] || []).join('、'),
        reviewCount: count,
        met: count >= cfg.reviewsPerStudent
      };
    });
    reviewers.sort(function (a, b) { return a.reviewCount - b.reviewCount; });
    result[a] = reviewers;
  });
  return result;
}

function bySeatAsc(a, b) {
  var na = parseInt(a.seat, 10), nb = parseInt(b.seat, 10);
  if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
  return String(a.seat).localeCompare(String(b.seat));
}
// 下面三個是給「匯出到試算表」用的簡化版，邏輯跟前端 index.html 的同名函式一致
// （出席與否、組長是誰），但這裡是獨立算一次，不共用前端的模組變數。
function normSeatServer(s) { return canonSeatServer(s); }
function presentOrderForServer(group) {
  var presentBySeat = {};
  (group.members || []).forEach(function (m) { presentBySeat[normSeatServer(m.seat)] = m.present !== false; });
  return (group.order || []).filter(function (p) {
    var v = presentBySeat[normSeatServer(p.seat)];
    return v !== false;
  });
}
function leaderSeatForServer(group) {
  var order = presentOrderForServer(group || {});
  return order.length ? order[0].seat : null;
}
/**
 * 老師要求把網頁上所有的報表都直接整理進試算表裡，方便他在試算表那邊統一
 * 總覽、算成績——不是塞進網頁畫面，全班快 30 個人，網頁上的表格太擠、
 * 看不清楚，直接開試算表用內建的排序/篩選、公式反而方便。
 * 每次點擊都會把下面這幾個分頁的內容整個清掉重寫成最新的資料，不是累加寫入：
 * 評分未達標名單、作業繳交狀況、評分總表、評分進度、器材輪借報表、器材放棄紀錄。
 */
function handleExportReportsToSheet(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var cfg = readConfig(true);
  var submissions = readSubmissions(true);
  var reviews = readReviews(true);
  var reviewersByAssignment = buildReviewersByAssignment(cfg, submissions, reviews);
  var assignments = cfg.assignments || [];

  var shortfallRows = [];
  assignments.forEach(function (a) {
    (reviewersByAssignment[a] || []).forEach(function (rv) {
      if (!rv.met) shortfallRows.push({ seat: rv.seat, name: rv.name, assignment: a, reviewCount: rv.reviewCount });
    });
  });
  shortfallRows.sort(function (x, y) {
    var s = bySeatAsc(x, y);
    return s !== 0 ? s : String(x.assignment).localeCompare(String(y.assignment));
  });
  var shortfallSheet = getSheet('評分未達標名單');
  shortfallSheet.clear();
  shortfallSheet.appendRow(['座號', '姓名', '作業', '已評件數', '需求件數', '還差幾件']);
  if (shortfallRows.length) {
    shortfallSheet.getRange(2, 1, shortfallRows.length, 6).setValues(shortfallRows.map(function (r) {
      return [r.seat, r.name, r.assignment, r.reviewCount, cfg.reviewsPerStudent, cfg.reviewsPerStudent - r.reviewCount];
    }));
  }

  var rosterMap = {};
  submissions.forEach(function (s) {
    var key = String(s.studentSeat || '').trim().toLowerCase();
    if (key && !rosterMap[key]) rosterMap[key] = { seat: s.studentSeat, name: s.studentName };
  });
  assignments.forEach(function (a) {
    (reviewersByAssignment[a] || []).forEach(function (rv) {
      var key = String(rv.seat || '').trim().toLowerCase();
      if (key && !rosterMap[key]) rosterMap[key] = { seat: rv.seat, name: rv.name };
    });
  });
  var roster = Object.keys(rosterMap).map(function (k) { return rosterMap[k]; }).sort(bySeatAsc);
  var submittedSet = {};
  submissions.forEach(function (s) {
    submittedSet[String(s.studentSeat || '').trim().toLowerCase() + '::' + normAssignment(s.assignment)] = true;
  });
  var matrixSheet = getSheet('作業繳交狀況');
  matrixSheet.clear();
  matrixSheet.appendRow(['座號', '姓名'].concat(assignments));
  if (roster.length) {
    matrixSheet.getRange(2, 1, roster.length, 2 + assignments.length).setValues(roster.map(function (s) {
      var row = [s.seat, s.name];
      assignments.forEach(function (a) {
        row.push(submittedSet[String(s.seat || '').trim().toLowerCase() + '::' + normAssignment(a)] ? '已交' : '未交');
      });
      return row;
    }));
  }

  // 評分總表：跟網頁上「評分總表」欄位順序一致，依座號排序方便對照點名冊、算成績。
  var gradeRows = buildRows(cfg, submissions, dropTestReviews(reviews)).slice().sort(function (a, b) {
    return bySeatAsc({ seat: a.studentSeat }, { seat: b.studentSeat });
  });
  var gradeSheet = getSheet('評分總表');
  gradeSheet.clear();
  var reportCriteria=uniqueCriteria(gradeRows.map(function(r){return r.criteria;}));
  var criteriaLabels = reportCriteria.map(function (c) { return c.label; });
  gradeSheet.appendRow(['座號', '姓名', '作業', '作品名稱'].concat(criteriaLabels).concat(['總平均', '份數', '裝置代碼', '教師加分', '退回原因', '榮譽加分', '含榮譽總分']));
  if (gradeRows.length) {
    gradeSheet.getRange(2, 1, gradeRows.length, 10 + criteriaLabels.length).setValues(gradeRows.map(function (r) {
      var critVals = reportCriteria.map(function (c) { var v = r.perCrit[c.key]; return v == null ? '' : v; });
      // 總平均已經包含老師加分（最高 5 分）；最後一欄另外列出加了幾分，方便對照。
      return [r.studentSeat, r.studentName, r.assignment || '', r.title || ''].concat(critVals).concat([r.overall == null ? '' : r.overall, r.count, r.deviceId || '', r.bonus || '', r.returned ? ('已退回（扣 ' + RETURN_PENALTY + '）：' + r.returned.reason) : '', r.award ? (r.awardKind + ' ＋' + r.award) : '', r.overall == null ? '' : Math.round((r.overall + (r.award || 0)) * 1e6) / 1e6]);
    }));
  }

  // 加分彙整：每位同學一列，把「額外加在總分上」的分數集中在一起（點名加分＋榮譽榜加分），方便直接加進學期成績。
  var extraBySeat = {};
  function extraFor(seat, name) {
    var key = canonSeatServer(seat);
    if (!extraBySeat[key]) extraBySeat[key] = { seat: padSeat(seat), name: name || '', att: 0, award: 0, awardKinds: [] };
    if (name && !extraBySeat[key].name) extraBySeat[key].name = name;
    return extraBySeat[key];
  }
  var attMap = cfg.attendanceBonus || {};
  Object.keys(attMap).forEach(function (k) { extraFor(attMap[k].seat || k, attMap[k].name).att = (attMap[k].points || 0) * ATTENDANCE_BONUS; });
  var awardMap = cfg.awardsByWork || {};
  Object.keys(awardMap).forEach(function (code) {
    var a = awardMap[code];
    var e = extraFor(a.seat, a.name);
    e.award += Number(a.bonus) || 0;
    e.awardKinds.push(a.assignment + ' ' + a.kind);
  });
  var extraRows = Object.keys(extraBySeat).map(function (k) { return extraBySeat[k]; }).sort(function (a, b) { return bySeatAsc({ seat: a.seat }, { seat: b.seat }); });
  var extraSheet = getSheet('加分彙整');
  extraSheet.clear();
  extraSheet.appendRow(['座號', '姓名', '點名加分', '榮譽榜加分', '額外加分合計', '榮譽榜獲獎']);
  if (extraRows.length) {
    extraSheet.getRange(2, 1, extraRows.length, 1).setNumberFormat('@');
    extraSheet.getRange(2, 1, extraRows.length, 6).setValues(extraRows.map(function (e) {
      return [e.seat, e.name, e.att, e.award, e.att + e.award, e.awardKinds.join('、')];
    }));
  }

  // 評分進度：每份作業、每個人一列，跟網頁「評分進度」表一致，多加一欄「作業」方便在同一個分頁裡看全部作業。
  var progressRows = [];
  assignments.forEach(function (a) {
    (reviewersByAssignment[a] || []).forEach(function (rv) {
      progressRows.push({ assignment: a, seat: rv.seat, name: rv.name, submitted: rv.submitted, reviewCount: rv.reviewCount, met: rv.met });
    });
  });
  progressRows.sort(function (x, y) {
    if (x.assignment !== y.assignment) return String(x.assignment).localeCompare(String(y.assignment));
    return bySeatAsc(x, y);
  });
  var progressSheet = getSheet('評分進度');
  progressSheet.clear();
  progressSheet.appendRow(['作業', '座號', '姓名', '已交作品', '已評分數', '是否達標']);
  if (progressRows.length) {
    progressSheet.getRange(2, 1, progressRows.length, 6).setValues(progressRows.map(function (r) {
      return [r.assignment, r.seat, r.name, r.submitted ? '已交' : '未交', r.reviewCount, r.met ? '已達標' : ('還差 ' + (cfg.reviewsPerStudent - r.reviewCount) + ' 件')];
    }));
  }

  // 器材輪借報表：組長、最後輪值的人（通常是最後拿著器材的人）、每項器材目前的狀態。
  var eq = readEquipment();
  var itemTypes = eq.itemTypes || [];
  var equipRows = (eq.groups || []).map(function (g) {
    var leaderSeat = leaderSeatForServer(g);
    var presentOrder = presentOrderForServer(g);
    var lastUser = presentOrder.length ? presentOrder[presentOrder.length - 1] : null;
    var statuses = g.equipmentStatus || [];
    var itemCells = itemTypes.map(function (name, i) {
      var st = statuses[i] || {};
      if (st.confirmed) return name + '：已確認';
      if (st.reported) return name + '：已回報（' + st.reportedBySeat + '號 ' + st.reportedByName + '）';
      return name + '：尚未回報';
    });
    return { name: g.name, leaderSeat: leaderSeat, lastUser: lastUser, itemCells: itemCells };
  });
  // 分頁名稱改成「器材輪借報表」：舊的「器材借用報表」分頁如果還在，直接改名沿用（Google 試算表會自動更新
  // 引用到它的公式），不會留下一個過時的舊分頁、也不會讓老師原本的總覽公式找不到資料。
  var ssForRename = SpreadsheetApp.getActiveSpreadsheet();
  var oldEquipSheet = CURRENT_CLASS ? null : ssForRename.getSheetByName('器材借用報表'); // 舊名稱只有最早的那一班可能有
  if (oldEquipSheet && !ssForRename.getSheetByName('器材輪借報表')) oldEquipSheet.setName('器材輪借報表');
  var equipSheet = getSheet('器材輪借報表');
  equipSheet.clear();
  var maxItemCols = itemTypes.length;
  var equipHeader = ['組名', '組長', '最後輪值'];
  // 同款器材（例如記憶卡x2）名稱會重複，欄位標題加上序號區分，不然兩欄都叫「記憶卡」搞不清楚是哪一份。
  var itemNameCount = {};
  itemTypes.forEach(function (name) {
    itemNameCount[name] = (itemNameCount[name] || 0) + 1;
    equipHeader.push(itemNameCount[name] > 1 ? (name + itemNameCount[name]) : name);
  });
  equipSheet.appendRow(equipHeader);
  if (equipRows.length) {
    equipSheet.getRange(2, 1, equipRows.length, equipHeader.length).setValues(equipRows.map(function (r) {
      var row = [r.name, r.leaderSeat ? (r.leaderSeat + ' 號') : '', r.lastUser ? (r.lastUser.seat + ' 號 ' + r.lastUser.name) : ''];
      for (var ei2 = 0; ei2 < maxItemCols; ei2++) row.push(r.itemCells[ei2] || '');
      return row;
    }));
  }

  // 器材放棄紀錄：誰在哪一組、什麼時候放棄了這次使用機會，依時間排序方便看時間軸。
  var waiverRows = [];
  (eq.groups || []).forEach(function (g) {
    (g.waiverLog || []).forEach(function (w) {
      waiverRows.push({ groupName: g.name, seat: w.seat, name: w.name, at: w.at });
    });
  });
  waiverRows.sort(function (x, y) { return (x.at || 0) - (y.at || 0); });
  var waiverSheet = getSheet('器材放棄紀錄');
  waiverSheet.clear();
  waiverSheet.appendRow(['組名', '座號', '姓名', '放棄時間']);
  if (waiverRows.length) {
    waiverSheet.getRange(2, 1, waiverRows.length, 4).setValues(waiverRows.map(function (r) {
      return [r.groupName, r.seat, r.name, r.at ? new Date(r.at) : ''];
    }));
  }

  return {
    ok: true, shortfallCount: shortfallRows.length, rosterCount: roster.length,
    gradeCount: gradeRows.length, progressCount: progressRows.length,
    equipCount: equipRows.length, waiverCount: waiverRows.length
  };
}

function handleTeacherFetch(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var cfg = readConfig();
  var submissions = readSubmissions();
  var reviews = readReviews();
  // 老師自己（99 號測試帳號）交的作品還是列在總表裡讓老師看得到，但它評的分不算進任何人的分數
  var rows = buildRows(cfg, submissions, dropTestReviews(reviews));
  var reviewersByAssignment = buildReviewersByAssignment(cfg, submissions, reviews);
  return { ok: true, config: cfg, rows: rows, reviewersByAssignment: reviewersByAssignment, submissionCount: submissions.length, reviewCount: reviews.length };
}

/**
 * 老師在作品牆播放時直接幫某件作品加分（0＝取消加分，1~3＝加幾分）。
 * 加分存在 Config 的 bonusByWork 裡（key 是作品編號），不動任何同學的評分紀錄，
 * 之後隨時可以改回 0 取消；實際的總平均在 buildRows 裡算，最高 5 分。
 */
function handleTeacherSetBonus(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var code = String(payload.workCode || '');
  if (!code) return { ok: false, error: 'no_work' };
  var n = parseInt(payload.bonus, 10);
  if (isNaN(n) || n < 0 || n > 3) return { ok: false, error: 'bad_bonus' };
  var reason = String(payload.reason || '').trim();
  if (!reason) return { ok: false, error: 'reason_required' }; // 加分、取消加分都要寫原因
  var cfg = readConfig(true);
  var map = cfg.bonusByWork || {};
  var notes = cfg.bonusNotes || {};
  var prev = bonusFor(cfg, code);
  if (n === 0) { delete map[code]; delete notes[code]; } else { map[code] = n; notes[code] = { reason: reason, at: Date.now() }; }
  cfg.bonusByWork = map;
  cfg.bonusNotes = notes;
  writeConfig(cfg);
  if (n !== prev) {
    var sub = submissionByCode(code) || {};
    logScoreChanges([{ type: '作品加分（總平均）', seat: sub.studentSeat, name: sub.studentName, assignment: sub.assignment, workCode: code, delta: n - prev, reason: (n === 0 ? '取消加分：' : '') + reason }]);
  }
  return { ok: true, workCode: code, bonus: n };
}

/**
 * 老師在作品牆播放時看「這件作品每位評分者各打幾分」的明細，用來檢查有沒有認真評分：
 * 每位評分者的座號姓名、各題分數（照題目順序）、他的平均、和其他評分者平均的差距，
 * 以及有沒有已經登記過「評分異常」。需要教師密語；老師自己的 99 號測試帳號的評分不算。
 * 整份都是同一個分數（例如全部 5 分）會標示 allSame，是最常見的隨便按的特徵。
 */
/** 一件作品的評分明細列（共用給「單件」和「全部作品一次載入」兩個動作）。revs 是這件作品收到的評分。 */
function workReviewRows(cfg, revs, nameByNorm, seatByNorm, flagged, code, criteria) {
  var list = revs.map(function (r) {
    var norm = r.reviewerNorm;
    var i = norm.indexOf('::');
    var vals = (criteria || cfg.criteria || []).map(function (c) { return typeof r.scores[c.key] === 'number' ? r.scores[c.key] : null; });
    var nums = vals.filter(function (v) { return v != null; });
    var at = r.createdAt instanceof Date ? r.createdAt.getTime() : (Number(r.createdAt) || null);
    return {
      seat: seatByNorm[norm] || (i < 0 ? '' : padSeat(norm.substring(0, i))),
      name: nameByNorm[norm] || (i < 0 ? norm : norm.substring(i + 2)),
      scores: vals,
      mean: nums.length ? nums.reduce(function (a, b) { return a + b; }, 0) / nums.length : null,
      allSame: nums.length >= 2 && nums.every(function (v) { return v === nums[0]; }),
      flagged: !!(flagged[norm] && flagged[norm].items && flagged[norm].items[code]),
      at: at
    };
  });
  var withMean = list.filter(function (x) { return x.mean != null; });
  var total = withMean.reduce(function (a, x) { return a + x.mean; }, 0);
  list.forEach(function (x) {
    x.gap = (x.mean != null && withMean.length > 1) ? round2(x.mean - (total - x.mean) / (withMean.length - 1)) : null;
    x.mean = round2(x.mean);
  });
  list.sort(function (a, b) { return Math.abs(b.gap || 0) - Math.abs(a.gap || 0); });
  return list;
}

function handleTeacherWorkReviews(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var code = String(payload.workCode || '');
  var cfg = readConfig();
  var subs = readSubmissions();
  var work = null;
  subs.forEach(function (s) { if (String(s.workCode) === code) work = s; });
  if (!work) return { ok: false, error: 'no_work' };
  var nameByNorm = {}, seatByNorm = {};
  subs.forEach(function (s) { nameByNorm[s.studentNorm] = s.studentName; seatByNorm[s.studentNorm] = s.studentSeat; });
  var revs = dropTestReviews(readReviews()).filter(function (r) { return r.workId === work.id; });
  var criteria=criteriaForWork(cfg,work,revs);
  return { ok: true, workCode: code, criteria: criteria.map(function(c){return c.label;}), reviews: workReviewRows(cfg, revs, nameByNorm, seatByNorm, cfg.flaggedReviewers || {}, code, criteria) };
}

/**
 * 全部作品的評分明細一次回傳：教師進入作品牆時在背景先載好，之後在播放中打開評分明細就不用再等一次來回。
 * 回傳 { works: { 作品編號: 明細列[] } }，內容和單件版（teacherWorkReviews）完全一樣。
 */
function handleTeacherAllReviews(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var cfg = readConfig();
  var subs = readSubmissions();
  var nameByNorm = {}, seatByNorm = {};
  subs.forEach(function (s) { nameByNorm[s.studentNorm] = s.studentName; seatByNorm[s.studentNorm] = s.studentSeat; });
  var byWork = {};
  dropTestReviews(readReviews()).forEach(function (r) { (byWork[r.workId] = byWork[r.workId] || []).push(r); });
  var flagged = cfg.flaggedReviewers || {};
  var works = {},criteriaByWork={};
  subs.forEach(function(s){
    var revs=byWork[s.id]||[],criteria=criteriaForWork(cfg,s,revs),code=String(s.workCode);
    criteriaByWork[code]=criteria.map(function(c){return c.label;});
    works[code]=workReviewRows(cfg,revs,nameByNorm,seatByNorm,flagged,code,criteria);
  });
  return {ok:true,criteriaByWork:criteriaByWork,works:works};
}

// ---- 評分不合理的同學 ----
// 老師覺得某件作品被評的分數不合理時，可以在作品牆播放中把「評分和其他人差距過大」或「這件作品全部的評分同學」登記下來：
// 寫進試算表「評分異常紀錄」分頁（誰、對哪件作品、平均幾分、和其他人差多少、什麼時候），
// 並且這些同學之後的評分畫面會一直看到「請認真評分」的提醒（見 withReviewWarning）。
var OUTLIER_DEFAULT_GAP = 1.5;

function reviewMean(criteria, scores) {
  var vals = (criteria || []).map(function (c) { return scores ? scores[c.key] : null; }).filter(function (v) { return typeof v === 'number'; });
  return vals.length ? vals.reduce(function (a, b) { return a + b; }, 0) / vals.length : null;
}

function round2(x) { return x == null ? null : Math.round(x * 100) / 100; }

function flaggedCount(cfg, norm) {
  var f = (cfg.flaggedReviewers || {})[norm];
  return f && f.items ? Object.keys(f.items).length : 0;
}

/** 在評分畫面用的回應上附上提醒：被老師登記過的同學會多一個 warning 欄位。 */
function withReviewWarning(res, payload) {
  try {
    var n = flaggedCount(readConfig(), payload.reviewerNorm);
    if (res && n > 0) res.warning = { count: n };
  } catch (e) {}
  return res;
}

function handleTeacherFlagReviews(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var code = String(payload.workCode || '');
  var mode = payload.mode;
  if (!code || (mode !== 'outliers' && mode !== 'all' && mode !== 'clear')) return { ok: false, error: 'bad_request' };
  var cfg = readConfig(true);
  var subs = readSubmissions();
  var work = null;
  subs.forEach(function (s) { if (String(s.workCode) === code) work = s; });
  if (!work) return { ok: false, error: 'no_work' };
  var nameByNorm = {}, seatByNorm = {};
  subs.forEach(function (s) { nameByNorm[s.studentNorm] = s.studentName; seatByNorm[s.studentNorm] = s.studentSeat; });
  function who(norm) {
    var i = norm.indexOf('::');
    return { seat: seatByNorm[norm] || (i < 0 ? '' : padSeat(norm.substring(0, i))), name: nameByNorm[norm] || (i < 0 ? norm : norm.substring(i + 2)) };
  }

  var relevantReviews=dropTestReviews(readReviews()).filter(function(r){return r.workId===work.id;}),workCriteria=criteriaForWork(cfg,work,relevantReviews);
  var list = relevantReviews.map(function (r) {
    return { norm: r.reviewerNorm, mean: reviewMean(workCriteria, r.scores) };
  }).filter(function (x) { return x.mean != null; });
  var total = list.reduce(function (a, x) { return a + x.mean; }, 0);
  list.forEach(function (x) {
    x.othersMean = list.length > 1 ? (total - x.mean) / (list.length - 1) : null;
    x.gap = x.othersMean == null ? null : x.mean - x.othersMean;
  });

  var flagged = cfg.flaggedReviewers || {};
  var candidates;
  if (mode === 'clear') {
    candidates = Object.keys(flagged).filter(function (norm) { return flagged[norm].items && flagged[norm].items[code]; }).map(function (norm) {
      var it = flagged[norm].items[code];
      return { norm: norm, mean: it.mean, othersMean: it.othersMean, gap: it.gap };
    });
  } else if (mode === 'all') {
    candidates = list;
  } else {
    if (list.length < 3) return { ok: false, error: 'too_few', count: list.length };
    var threshold = parseFloat(payload.gap);
    if (isNaN(threshold) || threshold <= 0) threshold = OUTLIER_DEFAULT_GAP;
    payload.gap = threshold;
    candidates = list.filter(function (x) { return Math.abs(x.gap) >= threshold; });
  }
  var out = candidates.map(function (x) {
    var w = who(x.norm);
    var already = mode === 'clear' ? true : !!(flagged[x.norm] && flagged[x.norm].items && flagged[x.norm].items[code]);
    return { seat: w.seat, name: w.name, mean: round2(x.mean), othersMean: round2(x.othersMean), gap: round2(x.gap), already: already };
  });
  if (payload.dryRun) return { ok: true, dryRun: true, candidates: out, reviewerCount: list.length };

  var at = Date.now();
  var log = getSheet('評分異常紀錄');
  if (log.getLastRow() === 0) log.appendRow(['時間', '動作', '作品編號', '作者座號', '作者姓名', '作業', '評分者座號', '評分者姓名', '他給這件的平均', '其他人給的平均', '差距', '登記方式', '備註']);
  var added = 0;
  candidates.forEach(function (x, i) {
    var w = who(x.norm);
    if (mode === 'clear') {
      var f = flagged[x.norm];
      delete f.items[code];
      if (!Object.keys(f.items).length) delete flagged[x.norm];
      log.appendRow([new Date(at), '撤銷登記', code, work.studentSeat, work.studentName, '', w.seat, w.name, round2(x.mean), round2(x.othersMean), round2(x.gap), '', '']);
    } else {
      var existing = flagged[x.norm] && flagged[x.norm].items && flagged[x.norm].items[code];
      if (existing) return; // 這件作品已經登記過這位同學，不重複登記
      flagged[x.norm] = flagged[x.norm] || { items: {} };
      flagged[x.norm].items[code] = { at: at, mean: round2(x.mean), othersMean: round2(x.othersMean), gap: round2(x.gap), mode: mode };
      added++;
      log.appendRow([new Date(at), '登記', code, work.studentSeat, work.studentName, '', w.seat, w.name, round2(x.mean), round2(x.othersMean), round2(x.gap),
        mode === 'all' ? '這件作品全部評分者' : '與其他人差距過大', mode === 'all' ? '' : ('差距門檻 ' + payload.gap + ' 分')]);
    }
  });
  // 作業名稱放在第 6 欄，用純文字寫進這次新增的列（像 "01" 這種名稱不要被試算表吃掉前面的 0）
  var rowsAdded = mode === 'clear' ? candidates.length : added;
  if (rowsAdded > 0) {
    var aRange = log.getRange(log.getLastRow() - rowsAdded + 1, 6, rowsAdded, 1);
    aRange.setNumberFormat('@');
    var vals = [];
    for (var k = 0; k < rowsAdded; k++) vals.push([work.assignment || '']);
    aRange.setValues(vals);
  }
  cfg.flaggedReviewers = flagged;
  writeConfig(cfg);
  return { ok: true, candidates: out, added: mode === 'clear' ? 0 : added, cleared: mode === 'clear' ? candidates.length : 0, reviewerCount: list.length };
}

/**
 * 登記某份作業的榮譽榜加分：目前榜上的得獎作者（前三名＋人氣獎）每人額外加 AWARD_BONUS 分總分，
 * 寫進試算表「榮譽榜加分」分頁（時間、動作、作業、獎項、座號、姓名、作品編號、加分、依據）。
 * 只能登記已經公布（解鎖）的作業。重新登記同一份作業會先取消之前登記的人再登記現在榜上的人
 * （兩邊都有留紀錄），所以不會重複加分；payload.undo 為 true 是整份作業的登記都取消。
 * 榮譽榜加分「不會」改變作品牆的總平均和名次（不然加分又會反過來改變誰得獎），
 * 它另外存起來，在教師總表／匯出裡以「榮譽加分」「含榮譽總分」兩欄呈現。
 */
function handleTeacherRegisterAwards(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var target = normAssignment(payload.assignment);
  var cfg = readConfig(true);
  var map = cfg.awardsByWork || {};
  var undo = !!payload.undo;
  var winners = [];
  if (!undo) {
    var honors = computeHonors();
    var item = null;
    honors.items.forEach(function (it) { if (normAssignment(it.assignment) === target) item = it; });
    if (!item || !item.unlocked) return { ok: false, error: 'not_unlocked' };
    (item.top || []).forEach(function (t) {
      winners.push({ kind: '第' + t.rank + '名', seat: t.seat, name: t.name, workCode: t.workCode, basis: '總平均 ' + t.score });
    });
    if (item.popular) winners.push({ kind: '人氣獎', seat: item.popular.seat, name: item.popular.name, workCode: item.popular.workCode, basis: '「' + item.popular.label + '」平均 ' + item.popular.score });
    if (!winners.length) return { ok: false, error: 'no_winners' };
  }
  var previous = Object.keys(map).filter(function (code) { return normAssignment(map[code].assignment) === target; });
  if (payload.dryRun) {
    return { ok: true, dryRun: true, winners: winners, previous: previous.map(function (c) { return { kind: map[c].kind, seat: map[c].seat, name: map[c].name }; }) };
  }
  var at = Date.now();
  var log = getSheet('榮譽榜加分');
  if (log.getLastRow() === 0) log.appendRow(['時間', '動作', '作業', '獎項', '座號', '姓名', '作品編號', '加分', '依據']);
  var logRows = [];
  previous.forEach(function (code) {
    var old = map[code];
    logRows.push([new Date(at), '取消', String(payload.assignment), old.kind, old.seat, old.name, code, 0, undo ? '取消整份作業的登記' : '重新登記，先取消舊的']);
    delete map[code];
  });
  winners.forEach(function (w) {
    map[String(w.workCode)] = { kind: w.kind, assignment: String(payload.assignment), at: at, bonus: AWARD_BONUS, seat: w.seat, name: w.name };
    logRows.push([new Date(at), '登記', String(payload.assignment), w.kind, w.seat, w.name, w.workCode, AWARD_BONUS, w.basis]);
  });
  var ledger = [];
  previous.forEach(function (code) {
    var old = map[code] || {};
    ledger.push({ type: '榮譽榜加分', seat: old.seat, name: old.name, assignment: payload.assignment, workCode: code, delta: -AWARD_BONUS, reason: '取消榮譽榜加分（' + old.kind + '）：' + (undo ? '取消整份作業的登記' : '重新登記，先取消舊的') });
  });
  winners.forEach(function (w) {
    ledger.push({ type: '榮譽榜加分', seat: w.seat, name: w.name, assignment: payload.assignment, workCode: w.workCode, delta: AWARD_BONUS, reason: '榮譽榜 ' + w.kind + '（' + w.basis + '）' });
  });
  logScoreChanges(ledger);
  if (logRows.length) {
    var start = log.getLastRow() + 1;
    var range = log.getRange(start, 1, logRows.length, 9);
    range.setNumberFormat('@'); // 作業名稱、座號像 "01" 的話不要被試算表吃掉前面的 0；時間和加分欄位下一行再改回日期／數字格式
    log.getRange(start, 1, logRows.length, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
    log.getRange(start, 8, logRows.length, 1).setNumberFormat('0.0');
    range.setValues(logRows.map(function (r) { return r.map(function (v, i) { return (i === 0 || i === 7) ? v : String(v); }); }));
  }
  cfg.awardsByWork = map;
  writeConfig(cfg);
  return { ok: true, winners: winners, cancelled: previous.length, bonus: AWARD_BONUS };
}

/**
 * 點名區幫同學加分或扣分：每按一次「加分」加 ATTENDANCE_BONUS 分總分（delta 為 1），按「扣分」扣同樣的分數（delta 為 -1），
 * 累計可以是負的。每一次都必須填原因：累計次數存在 Config 的 attendanceBonus（key 是座號），
 * 每一次動作另外記一列在試算表「點名加分」分頁（時間、座號、姓名、動作、這次分數、累計分數、原因），
 * 同時寫進「加扣分紀錄」總帳。回傳目前的累計次數，畫面以它為準。
 */
function handleTeacherAttendanceBonus(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var seat = canonSeatServer(payload.seat);
  if (!seat) return { ok: false, error: 'no_seat' };
  var reason = String(payload.reason || '').trim();
  if (!reason) return { ok: false, error: 'reason_required' };
  var delta = parseInt(payload.delta, 10) === -1 ? -1 : 1;
  var cfg = readConfig(true);
  var map = cfg.attendanceBonus || {};
  var rec = map[seat] || { seat: padSeat(payload.seat), name: String(payload.name || ''), points: 0 };
  rec.points = (rec.points || 0) + delta;
  if (payload.name) rec.name = String(payload.name);
  if (rec.points !== 0) map[seat] = rec; else delete map[seat];
  cfg.attendanceBonus = map;
  writeConfig(cfg);
  var log = getSheet('點名加分');
  if (log.getLastRow() === 0) log.appendRow(['時間', '座號', '姓名', '動作', '這次分數', '累計分數', '原因']);
  log.appendRow([new Date(), padSeat(payload.seat), rec.name, delta > 0 ? '加分' : '扣分', delta * ATTENDANCE_BONUS, rec.points * ATTENDANCE_BONUS, reason]);
  var seatCell = log.getRange(log.getLastRow(), 2);
  seatCell.setNumberFormat('@'); seatCell.setValue(padSeat(payload.seat)); // 座號 "01" 不要被試算表吃掉前面的 0
  logScoreChanges([{ type: delta > 0 ? '點名加分' : '點名扣分', seat: padSeat(payload.seat), name: rec.name, assignment: '', workCode: '', delta: delta * ATTENDANCE_BONUS, reason: reason }]);
  return { ok: true, points: rec.points, total: rec.points * ATTENDANCE_BONUS };
}

// ---- 從雲端硬碟匯入班級名單 ----
// 支援：Google 試算表（最好）、CSV／TXT 文字檔、Google 文件的純文字。
// Excel（.xlsx）請先在雲端硬碟用 Google 試算表開啟、另存成 Google 試算表。
// 欄位自動辨認：標題列有「座號／姓名／組別」就照標題；沒有標題就當作第一欄座號、第二欄姓名、第三欄組別（有的話）。
function extractDriveId(text) {
  var t = String(text || '').trim();
  var m = t.match(/\/d\/([a-zA-Z0-9_-]{15,})/) || t.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
  if (m) return m[1];
  return /^[a-zA-Z0-9_-]{25,}$/.test(t) ? t : '';
}
function findRosterFile(ref) {
  var id = extractDriveId(ref);
  if (id) return DriveApp.getFileById(id);
  var name = String(ref || '').trim();
  if (!name) return null;
  var it = DriveApp.getFilesByName(name), best = null;
  while (it.hasNext()) { var f = it.next(); if (!best || f.getLastUpdated() > best.getLastUpdated()) best = f; }
  return best; // 同名的有好幾個就取最近更新的；畫面上會顯示實際讀到的檔名讓老師確認
}
function decodeRosterBlob(blob) {
  var bytes = blob.getBytes();
  var text = Utilities.newBlob(bytes).getDataAsString('UTF-8');
  if (text.indexOf(' ') >= 0) { try { text = Utilities.newBlob(bytes).getDataAsString('Big5'); } catch (e) {} } // Excel 在台灣常存成 Big5 的 CSV
  return text.replace(/^﻿/, '');
}
function rosterTextToRows(text) {
  var delim = text.indexOf('\t') >= 0 ? 'tab' : (/[,，]/.test(text) ? 'comma' : 'space');
  return String(text).split(/\r?\n/).map(function (line) {
    line = line.trim();
    if (!line) return [];
    if (delim === 'tab') return line.split('\t');
    if (delim === 'comma') return line.split(/[,，]/);
    return line.split(/\s+/);
  });
}
function readRosterRows(file) {
  var mime = file.getMimeType(), name = file.getName();
  if (mime === 'application/vnd.google-apps.spreadsheet') {
    var sheets = SpreadsheetApp.openById(file.getId()).getSheets();
    var pick = sheets[0];
    for (var i = 0; i < sheets.length; i++) { if (/名單|座號|roster/i.test(sheets[i].getName())) { pick = sheets[i]; break; } }
    return { type: 'Google 試算表', rows: pick.getDataRange().getDisplayValues() };
  }
  if (mime === 'application/vnd.google-apps.document') {
    return { type: 'Google 文件', rows: rosterTextToRows(file.getAs('text/plain').getDataAsString().replace(/^﻿/, '')) };
  }
  if (/^text\//.test(mime) || /\.(csv|txt|tsv)$/i.test(name)) {
    return { type: '文字檔', rows: rosterTextToRows(decodeRosterBlob(file.getBlob())) };
  }
  throw new Error('unsupported_type:' + mime);
}
function rosterSeat(v) {
  var t = toHalfWidthServer(String(v == null ? '' : v)).replace(/\s+/g, '').replace(/號$/, '');
  if (!t) return '';
  return /^\d+$/.test(t) ? padSeat(String(parseInt(t, 10))) : t;
}
function rosterFromRows(rows) {
  rows = rows.map(function (r) { return r.map(function (c) { return String(c == null ? '' : c).trim(); }); }).filter(function (r) { return r.some(function (c) { return c; }); });
  var SEAT_RE = /座號|座位|號碼|學號|^號$|^no\.?$/i, NAME_RE = /姓名|名字|學生|^name$/i, GROUP_RE = /組別|分組|小組|^組$|^group$/i;
  var seatCol = -1, nameCol = -1, groupCol = -1, start = 0;
  for (var i = 0; i < Math.min(6, rows.length) && seatCol < 0; i++) {
    var sc = -1, nc = -1, gc = -1;
    rows[i].forEach(function (c, j) { if (sc < 0 && SEAT_RE.test(c)) sc = j; else if (nc < 0 && NAME_RE.test(c)) nc = j; else if (gc < 0 && GROUP_RE.test(c)) gc = j; });
    if (sc >= 0 && nc >= 0) { seatCol = sc; nameCol = nc; groupCol = gc; start = i + 1; }
  }
  if (seatCol < 0) {
    seatCol = 0; nameCol = 1;
    var withThird = rows.filter(function (r) { return r[2]; }).length;
    groupCol = (rows.length && withThird >= rows.length / 2) ? 2 : -1;
  }
  var groups = [], byName = {}, seen = {}, skipped = 0, warnings = [];
  rows.slice(start).forEach(function (r) {
    var seat = rosterSeat(r[seatCol]), name = String(r[nameCol] || '').trim();
    if (!seat || !name) { skipped++; return; }
    var key = canonSeatServer(seat);
    if (seen[key]) { warnings.push('座號 ' + seat + ' 重複（' + seen[key] + '、' + name + '），只保留前面那一位'); return; }
    seen[key] = name;
    var gname = groupCol >= 0 ? (r[groupCol] || '未分組') : '全班';
    if (!byName[gname]) { byName[gname] = { name: gname, members: [] }; groups.push(byName[gname]); }
    byName[gname].members.push({ seat: seat, name: name });
  });
  groups.forEach(function (g) { g.members.sort(function (a, b) { return (parseInt(a.seat, 10) || 0) - (parseInt(b.seat, 10) || 0); }); });
  if (skipped) warnings.push(skipped + ' 列缺少座號或姓名，已略過');
  if (groupCol < 0) warnings.push('檔案裡沒有「組別」欄位，全部放在「全班」一組；要分組請之後自己在名單裡調整');
  var count = groups.reduce(function (a, g) { return a + g.members.length; }, 0);
  var groupsText = groups.map(function (g) { return g.name + '\n' + g.members.map(function (m) { return m.seat + ' ' + m.name; }).join('\n'); }).join('\n\n');
  return { groups: groups, count: count, warnings: warnings, groupsText: groupsText };
}
/** 讀取檔案、整理成名單。錯誤用 { ok:false, error } 回傳，不丟例外。 */
function parseRosterRef(ref) {
  if (!String(ref || '').trim()) return { ok: false, error: 'no_file' };
  var file;
  try { file = findRosterFile(ref); } catch (e) { return { ok: false, error: 'file_not_found', detail: String(e && e.message || e) }; }
  if (!file) return { ok: false, error: 'file_not_found' };
  var read;
  try { read = readRosterRows(file); } catch (e2) {
    var msg = String(e2 && e2.message || e2);
    if (msg.indexOf('unsupported_type') === 0) return { ok: false, error: 'unsupported_type', detail: msg.substring(17), fileName: file.getName() };
    return { ok: false, error: 'read_failed', detail: msg, fileName: file.getName() };
  }
  var r = rosterFromRows(read.rows);
  if (!r.count) return { ok: false, error: 'no_students', fileName: file.getName() };
  r.ok = true; r.fileName = file.getName(); r.fileType = read.type;
  return r;
}
function handleRosterFromFile(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var r = parseRosterRef(payload.file);
  if (r.ok) r.groupCount = r.groups.length;
  return r;
}

/** 手機端帶來的分組名單：只留下合理長度的座號與姓名，擋掉亂傳的資料。 */
function cleanRosterGroups(groups) {
  var out = [], count = 0, seen = {};
  (groups || []).slice(0, 40).forEach(function (g) {
    var members = [];
    (g.members || []).slice(0, 200).forEach(function (m) {
      var seat = String(m && m.seat || '').trim().slice(0, 12), name = String(m && m.name || '').trim().slice(0, 30);
      var key = canonSeatServer(seat);
      if (!seat || !name || seen[key]) return;
      seen[key] = true; members.push({ seat: seat, name: name }); count++;
    });
    if (members.length) out.push({ name: String(g.name || '全班').trim().slice(0, 30) || '全班', members: members });
  });
  return { groups: out, count: count };
}

/**
 * 一鍵新增班級：在同一個後端裡開一個全新的班（所有分頁加上班級代號前綴，資料和原本的班完全分開）。
 * 可以沿用目前這個班的作業清單、評分題目與公開門檻；教師密語一定要另外設一個新的；
 * 名單可以順便從雲端硬碟檔案匯入。班級代號只能是英數字或中文，最多 12 字。
 */
function handleCreateClass(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var id = cleanClassId(payload.classId);
  if (!id) return { ok: false, error: 'bad_class_id' };
  if (classExists(id, true)) return { ok: false, error: 'class_exists' };
  var pass = String(payload.newPasscode || '').trim();
  if (pass.length < 4) return { ok: false, error: 'passcode_too_short' };
  var roster = null;
  if (payload.rosterGroups && payload.rosterGroups.length) {
    // 名單是老師在手機上選的檔案、已經整理成分組格式直接帶過來的：只做格式檢查，不從雲端硬碟讀
    var clean = cleanRosterGroups(payload.rosterGroups);
    if (!clean.count) return { ok: false, error: 'roster_failed', rosterError: 'no_students' };
    roster = { ok: true, groups: clean.groups, count: clean.count, fileName: String(payload.rosterFileName || '手機上的檔案').slice(0, 60) };
  } else if (String(payload.rosterFile || '').trim()) {
    roster = parseRosterRef(payload.rosterFile);
    if (!roster.ok) return { ok: false, error: 'roster_failed', rosterError: roster.error, detail: roster.detail, fileName: roster.fileName }; // 名單讀不到就不要開班，免得開了一半
  }
  var source = readConfig(true);
  var previous = CURRENT_CLASS;
  CURRENT_CLASS = id;
  try {
    ensureHeadersUncached(); // 建立這個班需要的所有分頁（標題列、預設設定）
    var cfg = defaultConfig();
    cfg.className = String(payload.className || id).trim().slice(0, 20);
    cfg.teacherPasscode = pass;
    if (payload.copySettings) {
      cfg.criteria = source.criteria;
      cfg.assignmentCriteria=source.assignmentCriteria||{};
      cfg.criteriaArchive=source.criteriaArchive||{};
      cfg.classSchedule=source.classSchedule||defaultClassSchedule();
      cfg.reviewsPerStudent = source.reviewsPerStudent;
      cfg.assignments = source.assignments;
      cfg.revealMinReviewers = source.revealMinReviewers;
      cfg.revealMinReviews = source.revealMinReviews;
    }
    writeConfig(cfg);
    if (roster) {
      var res = handleEquipmentSetGroups({ passcode: pass, groups: roster.groups });
      if (!res || res.ok === false) throw new Error('set_groups_failed');
    }
    bumpDataVersion();
    CacheService.getScriptCache().put('clsok:' + id, '1', 600);
  } finally {
    CURRENT_CLASS = previous;
  }
  return { ok: true, classId: id, className: String(payload.className || id).trim().slice(0, 20), imported: roster ? roster.count : 0, fileName: roster ? roster.fileName : '' };
}

/** 這個後端目前有哪些班（代號與班級名稱），教師才看得到；原本的那一班代號是空的。 */
function handleListClasses(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var current = CURRENT_CLASS;
  var out = [];
  CURRENT_CLASS = '';
  try { out.push({ id: '', name: String(readConfig(true).className || '') }); } finally { CURRENT_CLASS = current; }
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sh) {
    var m = sh.getName().match(/^(.+)_Config$/);
    if (!m || cleanClassId(m[1]) !== m[1]) return;
    var nm = '';
    try { nm = String(JSON.parse(sh.getRange(2, 1).getValue()).className || ''); } catch (e) {}
    out.push({ id: m[1], name: nm });
  });
  return { ok: true, current: current, classes: out };
}

/**
 * 老師在作品牆退回一件作品（例如學生拿別人的作業上傳）：總平均扣 0.1 分，作品不再公開、不再抽給同學評，
 * 同學自己的「作業繳交狀況」會看到被退回和原因（而且不能自己刪掉，要另外再交一件）。
 * 每一次退回／撤銷都會在試算表的「退回紀錄」分頁多一列：時間、動作、座號、姓名、作業、作品編號、扣分、原因。
 * payload.undo 為 true 代表撤銷退回（扣分一併還原）。
 */
function handleTeacherReturnWork(payload) {
  if (!checkPasscode(payload)) return { ok: false, error: 'unauthorized' };
  var code = String(payload.workCode || '');
  if (!code) return { ok: false, error: 'no_work' };
  var sub = null;
  readSubmissions().forEach(function (s) { if (String(s.workCode) === code) sub = s; });
  if (!sub) return { ok: false, error: 'no_work' };
  var cfg = readConfig(true);
  var map = cfg.returnedByWork || {};
  var undo = !!payload.undo;
  var reason = String(payload.reason || '').trim();
  var at = Date.now();
  if (undo) {
    if (!map[code]) return { ok: true, workCode: code, returned: null };
    if (!reason) return { ok: false, error: 'reason_required' }; // 撤銷退回（把扣的分還回去）也要寫原因
    delete map[code];
  } else {
    if (!reason) return { ok: false, error: 'reason_required' };
    map[code] = { reason: reason, at: at };
  }
  cfg.returnedByWork = map;
  writeConfig(cfg);
  var log = getSheet('退回紀錄');
  if (log.getLastRow() === 0) log.appendRow(['時間', '動作', '座號', '姓名', '作業', '作品編號', '扣分', '原因']);
  log.appendRow([new Date(at), undo ? '撤銷退回' : '退回', sub.studentSeat, sub.studentName, sub.assignment || '', code, undo ? 0 : RETURN_PENALTY, reason]);
  var aRange = log.getRange(log.getLastRow(), 5);
  aRange.setNumberFormat('@'); aRange.setValue(sub.assignment || ''); // 作業名稱像 "01" 的話，不要被試算表吃掉前面的 0
  logScoreChanges([{ type: undo ? '撤銷退回' : '退回', seat: sub.studentSeat, name: sub.studentName, assignment: sub.assignment, workCode: code, delta: undo ? RETURN_PENALTY : -RETURN_PENALTY, reason: reason }]);
  return { ok: true, workCode: code, returned: undo ? null : { reason: reason, at: at } };
}

function handleTeacherUpdateConfig(payload) {
  if (!checkPasscode(payload)) return {ok:false,error:'unauthorized'};
  var previous=readConfig(true);
  if(!validCriteria(payload.criteria)) return {ok:false,error:'invalid_criteria'};
  var byAssignment=payload.assignmentCriteria===undefined ? (previous.assignmentCriteria||{}) : payload.assignmentCriteria;
  if(!byAssignment||typeof byAssignment!=='object'||Array.isArray(byAssignment)||!Object.keys(byAssignment).every(function(k){return validCriteria(byAssignment[k]);})) return {ok:false,error:'invalid_criteria'};
  var schedule=payload.classSchedule===undefined ? previous.classSchedule : payload.classSchedule;
  if(!validClassSchedule(schedule)) return {ok:false,error:'invalid_schedule'};
  var archive=previous.criteriaArchive||{};
  (previous.assignments||[]).forEach(function(a){archive[a]=uniqueCriteria([archive[a],criteriaForAssignment(previous,a)]);});
  var eq=readEquipment(); upgradeEquipmentTimer(eq);
  settleEquipmentTimer(eq,Date.now()); eq.classSchedule=schedule; writeEquipment(eq);
  var newCfg = {
    // 沒帶班級名稱的舊版畫面送來的儲存請求，要保留原本的名稱，不能清掉
    className: payload.className === undefined ? (readConfig(true).className || '') : String(payload.className || '').trim().slice(0, 20),
    criteria: payload.criteria,
    assignmentCriteria: byAssignment, criteriaArchive: archive, classSchedule: schedule,
    reviewsPerStudent: payload.reviewsPerStudent,
    teacherPasscode: payload.teacherPasscode,
    namesRevealedByAssignment: (payload.namesRevealedByAssignment && typeof payload.namesRevealedByAssignment === 'object') ? payload.namesRevealedByAssignment : {},
    revealMinReviewers: (payload.revealMinReviewers != null && !isNaN(payload.revealMinReviewers)) ? Math.max(0, parseInt(payload.revealMinReviewers, 10)) : 15,
    revealMinReviews: (payload.revealMinReviews != null && !isNaN(payload.revealMinReviews)) ? Math.max(1, parseInt(payload.revealMinReviews, 10)) : 5,
    assignmentDescriptions: (payload.assignmentDescriptions && typeof payload.assignmentDescriptions === 'object') ? payload.assignmentDescriptions : {},
    bonusByWork: readConfig(true).bonusByWork || {}, // 老師加分不在設定表單裡改，儲存設定時要原樣保留
    returnedByWork: readConfig(true).returnedByWork || {}, // 退回的作品也一樣
    flaggedReviewers: readConfig(true).flaggedReviewers || {}, // 登記過評分不合理的同學也一樣
    awardsByWork: readConfig(true).awardsByWork || {}, // 榮譽榜加分也一樣
    attendanceBonus: readConfig(true).attendanceBonus || {}, // 點名加分也一樣
    bonusNotes: readConfig(true).bonusNotes || {}, // 作品加分的原因也一樣
    assignments: (payload.assignments && payload.assignments.length) ? payload.assignments : ['第一次作業']
  };
  writeConfig(newCfg);
  return { ok: true };
}

function handleHome() {
  // 只是要「總共幾筆」，不用把整張表讀出來逐列解析，直接看最後一列是第幾列就好。
  var cfg = readConfig();
  return {
    submissionCount: Math.max(0, getSheet('Submissions').getLastRow() - 1),
    reviewCount: Math.max(0, getSheet('Reviews').getLastRow() - 1),
    reviewsPerStudent: cfg.reviewsPerStudent,
    className: cfg.className || '',
    assignments: cfg.assignments || [],
    assignmentDescriptions: cfg.assignmentDescriptions || {}
  };
}

/**
 * 打包下載某個學生交過的所有作品圖檔。只處理存在我們雲端硬碟資料夾裡的
 * 圖片（直接上傳的），或任何 driveLink 指到的、老師帳號至少能檢視的
 * 雲端硬碟檔案；純外部連結（例如貼 imgur 網址）沒辦法在伺服器端讀到
 * 檔案內容，會被跳過，不會讓整個打包失敗。
 */
function handleDownloadBundle(payload) {
  var norm = payload.studentNorm;
  var mine = readSubmissions().filter(function (s) { return s.studentNorm === norm; });
  var blobs = [];
  mine.forEach(function (s) {
    var link = String(s.driveLink || '');
    var m = link.match(/[?&]id=([a-zA-Z0-9_-]+)/) || link.match(/\/d\/([a-zA-Z0-9_-]+)/);
    var blob = null;
    if (m) {
      try { blob = DriveApp.getFileById(m[1]).getBlob(); } catch (e) {}
    }
    if (!blob && link) {
      try { blob = UrlFetchApp.fetch(link, { muteHttpExceptions: true }).getBlob(); } catch (e) {}
    }
    if (blob) {
      var ext = (blob.getContentType() && blob.getContentType().indexOf('png') !== -1) ? '.png' : '.jpg';
      blob.setName(sanitizeFilenamePart(s.assignment) + '_' + sanitizeFilenamePart(s.workCode) + '_' + sanitizeFilenamePart(s.title || '') + ext);
      blobs.push(blob);
    }
  });
  if (!blobs.length) return { ok: false, error: 'no_files' };
  var zipBlob = Utilities.zip(blobs, sanitizeFilenamePart(payload.studentName) + '_作品.zip');
  return { ok: true, zipBase64: Utilities.base64Encode(zipBlob.getBytes()), filename: zipBlob.getName() };
}

function dispatch(action, payload) {
  switch (action) {
    case 'home': return handleHome();
    case 'getMySubmissions': return handleGetMySubmissions(payload);
    case 'addSubmission': return handleAddSubmission(payload);
    case 'deleteSubmission': return handleDeleteSubmission(payload);
    case 'getNextWork': return withReviewWarning(handleGetNextWork(payload), payload);
    case 'getMyReviewProgress': return withReviewWarning(handleGetMyReviewProgress(payload), payload);
    case 'addReview': return handleAddReview(payload);
    case 'gallery': return handleGallery(payload);
    case 'honors': return handleHonors();
    case 'teacherWorkReviews': return handleTeacherWorkReviews(payload);
    case 'teacherAllReviews': return handleTeacherAllReviews(payload);
    case 'downloadBundle': return handleDownloadBundle(payload);
    case 'equipmentGet': return handleEquipmentGet();
    case 'equipmentSetGroups': return handleEquipmentSetGroups(payload);
    case 'equipmentSetAttendance': return handleEquipmentSetAttendance(payload);
    case 'equipmentBuildOrder': return handleEquipmentBuildOrder(payload);
    case 'equipmentSetDuration': return handleEquipmentSetDuration(payload);
    case 'equipmentSetDurationRange': return handleEquipmentSetDurationRange(payload);
    case 'equipmentSetItemTypes': return handleEquipmentSetItemTypes(payload);
    case 'equipmentReportReturned': return handleEquipmentReportReturned(payload);
    case 'equipmentConfirmReturned': return handleEquipmentConfirmReturned(payload);
    case 'equipmentWaive': return handleEquipmentWaive(payload);
    case 'equipmentResetReturned': return handleEquipmentResetReturned(payload);
    case 'equipmentStop': return handleEquipmentStop(payload);
    case 'equipmentResetTimer': return handleEquipmentResetTimer(payload);
    case 'equipmentStart': return handleEquipmentStart(payload);
    case 'equipmentTogglePause': return handleEquipmentTogglePause(payload);
    case 'renameUploads': return handleRenameUploads(payload);
    case 'verifyTeacher': return handleVerifyTeacher(payload);
    case 'accountStatus': return handleAccountStatus(payload);
    case 'login': return handleLogin(payload);
    case 'register': return handleRegister(payload);
    case 'resetPassword': return handleResetPassword(payload);
    case 'teacherFetch': return handleTeacherFetch(payload);
    case 'teacherUpdateConfig': return handleTeacherUpdateConfig(payload);
    case 'teacherSetBonus': return handleTeacherSetBonus(payload);
    case 'teacherReturnWork': return handleTeacherReturnWork(payload);
    case 'teacherFlagReviews': return handleTeacherFlagReviews(payload);
    case 'teacherRegisterAwards': return handleTeacherRegisterAwards(payload);
    case 'createClass': return handleCreateClass(payload);
    case 'listClasses': return handleListClasses(payload);
    case 'rosterFromFile': return handleRosterFromFile(payload);
    case 'teacherAttendanceBonus': return handleTeacherAttendanceBonus(payload);
    case 'resetAllData': return resetAllData(payload);
    default: return { ok: false, error: 'unknown_action' };
  }
}

var READ_ONLY_ACTIONS = {
  home: true, getMySubmissions: true, getNextWork: true, getMyReviewProgress: true,
  gallery: true, honors: true, teacherWorkReviews: true, teacherAllReviews: true, listClasses: true, rosterFromFile: true, downloadBundle: true, equipmentGet: true, teacherFetch: true,
  renameUploads: true, // 只改雲端硬碟檔名、不寫資料，而且可能跑好幾分鐘，不能拿鎖卡住全班（見 handleRenameUploads）
  verifyTeacher: true, accountStatus: true, login: true // 只讀帳號資料、不寫試算表（登入失敗次數記在快取裡），不用拿鎖也不用換資料版本
};

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 注意：讀取跟寫入的動作都走 doGet（用網址參數帶資料），這是刻意的。
 * 瀏覽器對 Apps Script 的 POST 有已知的重新導向/跨網域問題，會讓 fetch()
 * 出現「Failed to fetch」，改用 GET 就不會有這個問題。
 * doPost 只保留給「上傳圖片」這一種資料量太大、塞不進網址的情況使用
 * （前端用 no-cors 模式送出，不讀回應，送完後再用 doGet 去確認有沒有存成功）。
 */
function doGet(e) {
  var noClass = enterClass(e);
  if (noClass) return noClass;
  ensureHeaders();
  var action = (e.parameter && e.parameter.action) || 'home';
  var payload = {};
  if (e.parameter && e.parameter.payload) {
    try { payload = JSON.parse(e.parameter.payload); } catch (ex) { payload = {}; }
  }
  return jsonOut(runAction(action, payload));
}

/**
 * 純讀取的動作不用排隊拿鎖：以前所有請求（包含全班每幾秒一次的器材輪詢）都擠在同一把鎖後面一個一個處理，
 * 人一多佇列就塞住，連交作品、評分都要跟著等。只有會寫入資料的動作才需要鎖。
 * 寫入動作做完（不管成功失敗）就換新的資料版本，讓上面那些快取全部作廢；器材相關動作不影響
 * 作品／評分／設定，不用換。
 * 評分送出之後，順便在鎖外面算好「下一件要評的作品」一起回給前端，省掉學生每評完一件還要多等一次來回。
 */
/**
 * 交作品最慢的一步是把圖存進雲端硬碟（1～3 秒）。以前整段都包在全域鎖裡，全班同時交的話
 * 一個一個排隊，排到後面的等鎖超過 30 秒就失敗、作品也不見了。
 * 現在先在鎖「外面」把圖存好（多個人可以同時存），再進鎖裡只做寫試算表那一小步（約 0.3 秒）。
 * 如果這次其實不會新增（同一個編號重送、內容一模一樣已經交過），先用快取的資料檢查，就不存圖；
 * 萬一還是撞到（兩個重送同時進來）或逾時失敗，事後會把多存的圖丟進垃圾桶（見 trashSavedFile）。
 */
function prepareSubmissionImage(payload) {
  var mine = 0, same = false, dup = false;
  readSubmissions().forEach(function (s) {
    if (s.id === payload.id) same = true;
    if (canonNorm(s.studentNorm) === payload.studentNorm) {
      mine++;
      if (payload.imageHash && s.imageHash === payload.imageHash && normAssignment(s.assignment) === normAssignment(payload.assignment)) dup = true;
    }
  });
  if (same || dup) { payload.skipImage = true; delete payload.imageBase64; return; }
  var link = saveUploadedImage(uploadFilename(padSeat(payload.studentSeat), payload.studentName, payload.assignment, mine + 1), payload.imageBase64, payload.imageMime);
  payload.savedLink = link;
  var m = String(link).match(/[?&]id=([a-zA-Z0-9_-]+)/);
  payload.savedFileId = m ? m[1] : '';
  delete payload.imageBase64;
}
function trashSavedFile(payload) {
  try { if (payload && payload.savedFileId) DriveApp.getFileById(payload.savedFileId).setTrashed(true); } catch (e) {}
}

function runAction(action, payload) {
  // 學生識別碼（座號::姓名）先統一成標準形式，不管是 1／01／全形，都會比對到同一個人。
  if (payload && payload.studentNorm) payload.studentNorm = canonNorm(payload.studentNorm);
  if (payload && payload.reviewerNorm) payload.reviewerNorm = canonNorm(payload.reviewerNorm);
  if (authFailed(action, payload)) return { ok: false, error: 'auth' };
  if (READ_ONLY_ACTIONS[action]) return dispatch(action, payload);
  if (action === 'addSubmission' && payload && payload.imageBase64) {
    try { prepareSubmissionImage(payload); } catch (ex) { return { ok: false, error: 'image_save_failed' }; }
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (ex) {
    // 排隊逾時：交作品的話回傳「忙碌中」並清掉多存的圖（學生端會自動重送），其他動作維持原本的行為（直接丟出錯誤）
    if (action === 'addSubmission') { trashSavedFile(payload); return { ok: false, error: 'busy' }; }
    throw ex;
  }
  var result;
  try {
    result = dispatch(action, payload);
  } catch (ex2) {
    if (action === 'addSubmission') trashSavedFile(payload);
    throw ex2;
  } finally {
    if (action.indexOf('equipment') !== 0) bumpDataVersion();
    lock.releaseLock();
  }
  if (action === 'addSubmission' && (!result || result.ok === false || result.reused)) trashSavedFile(payload);
  if (action === 'addReview' && result && result.ok) {
    try { result.next = withReviewWarning(handleGetNextWork(payload), payload); } catch (ex) {}
  }
  return result;
}

/**
 * 前端上傳圖片是用隱藏表單送出（不是用 fetch 的 JSON body），
 * 所以這裡跟 doGet 一樣，從 e.parameter 讀 action/payload，
 * 而不是從 e.postData 讀——這樣可以完全避開瀏覽器對 fetch() + POST
 * 在 Apps Script 重新導向機制上的已知問題。
 */
function doPost(e) {
  var noClass = enterClass(e);
  if (noClass) return noClass;
  ensureHeaders();
  var action = (e.parameter && e.parameter.action) || 'home';
  var payload = {};
  if (e.parameter && e.parameter.payload) {
    try { payload = JSON.parse(e.parameter.payload); } catch (ex) { payload = {}; }
  }
  return jsonOut(runAction(action, payload));
}
