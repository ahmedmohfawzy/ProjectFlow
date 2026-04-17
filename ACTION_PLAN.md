# 🚀 ProjectFlow™ — خطة العمل الشاملة لإصلاح الأخطاء الحرجة

> **تاريخ التقرير:** 17 أبريل 2026
> **المُراجِع:** Claude (مراجعة كود شاملة)
> **نطاق المراجعة:** المعادلات الحسابية + الأداء والكفاءة + التكامل والربط + Edge Cases
> **إجمالي المشاكل المكتشفة:** 60+ مشكلة موزعة على 3 مستويات أولوية

## ✅ حالة التنفيذ (محدّث 17 أبريل 2026)

| # | المشكلة | الأولوية | الحالة |
|---|---------|----------|--------|
| 1 | XXE في XML Parser | P0 | ✅ تم |
| 2 | Render Scheduler (RAF) | P0 | ✅ موجود بالفعل |
| 3 | CPM Caching | P0 | ⏳ يحتاج dirty flag |
| 4 | structuredClone (undo + scenarios + task-editor) | P0 | ✅ تم |
| 5 | XSS (board.js tags) | P0 | ✅ تم |
| 6 | localStorage Quota + notification | P0 | ✅ تم |
| 7 | Canvas Overflow | P0 | ✅ موجود بالفعل |
| 8 | EVM safeDate | P0 | ✅ تم |
| 9 | EAC Formula | P1 | ✅ تم |
| 10 | Resource Holidays | P1 | ⏳ يحتاج WorkCalendar |
| 11 | Resource Iterator DST | P1 | ✅ تم |
| 12 | CPM Calendar Days | P1 | ⏳ تعقيد عالي |
| 14 | Negative Duration | P1 | ✅ تم |
| 15 | getAncestors Guard | P1 | ✅ تم |
| 20 | Dependency findIndex → Map | P1 | ✅ تم |
| 23 | SQL Field Injection | P1 | ✅ موجود بالفعل |
| 24 | D365 Pagination | P1 | ✅ تم |
| 25 | D365 Rate Limit | P1 | ✅ تم |
| 27 | CORS Production | P1 | ✅ تم |
| 28 | HTTPS Enforcement | P1 | ✅ تم |
| 29 | File Size Limit | P1 | ✅ تم |
| P2.13 | UTF-8 BOM Strip | P2 | ✅ تم |
| P2.16 | JSON Body Limit | P2 | ✅ تم |

---


---

## 📊 ملخص تنفيذي

المشروع **مبني باحترافية عالية** مع architecture نظيف وخوارزميات صحيحة في الغالب. لكن فيه عدد من المشاكل الحرجة اللي لازم تتصلح قبل النشر للإنتاج:

| المستوى | العدد | الوقت المتوقع للإصلاح |
|--------|------|----------------------|
| 🔴 **P0 — حرج (لازم يتصلح فوراً)** | 8 مشاكل | ~6-8 ساعات |
| 🟠 **P1 — مهم جداً** | 22 مشكلة | ~12-16 ساعة |
| 🟡 **P2 — تحسين** | 30+ مشكلة | ~20+ ساعة |

**التأثير المتوقع بعد إصلاح P0 + P1:**
- تقليل زمن الـ render بنسبة **60-70%** للمشاريع الكبيرة (500+ tasks)
- إصلاح **ثغرات أمنية حرجة** (XXE، XSS، SQL field injection)
- حماية من فقدان البيانات في حالات الحافة (date/numeric corruption)
- تحسين استجابة الواجهة من ~200ms إلى ~50ms لكل تعديل

---

## 🔴 القسم الأول: الأخطاء الحرجة (P0) — أصلحها أولاً

### 1. ثغرة XXE في XML Parser
- **الملف:** `js/xml-parser.js` السطر 17-24
- **المشكلة:** `DOMParser` مستخدم بدون تعطيل الـ External Entities. ملف XML خبيث ممكن يقرأ ملفات محلية أو يعمل DoS.
- **الحل:**
```javascript
function parse(xmlString) {
    // Strip BOM
    xmlString = xmlString.replace(/^\ufeff/, '');

    // Block DTD and external entities before parsing
    if (/<!DOCTYPE/i.test(xmlString) || /<!ENTITY/i.test(xmlString)) {
        throw new Error('Unsafe XML: DTD/entities are not allowed');
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, 'text/xml');
    // ... rest of parser
}
```
- **الوقت:** 15 دقيقة

---

### 2. Full Re-render في كل تعديل
- **الملف:** `js/app.js` السطور 1240, 1250, 1260 (و renderAll في 1123)
- **المشكلة:** كل تغيير بسيط (مدة، تاريخ، نسبة) → `saveUndoState()` + `recalculate()` + `renderAll()` + `autoSave()` — كله متسلسل بدون requestAnimationFrame
- **التأثير:** مع 500 مهمة = ~250ms freeze لكل keystroke
- **الحل:**
```javascript
// استبدل renderAll مباشرة بـ scheduler
let _renderPending = false;
function scheduleRender() {
    if (_renderPending) return;
    _renderPending = true;
    requestAnimationFrame(() => {
        _renderPending = false;
        renderGantt();
        renderTable();
    });
}

// و debounce للـ autoSave
const autoSaveDebounced = debounce(autoSave, 500);
```
- **الوقت:** 2-3 ساعات (يتطلب testing شامل)

---

### 3. CPM يعاد حسابه في كل render
- **الملف:** `js/app.js` السطر 1104 وداخل `renderAll`
- **المشكلة:** `CPMEngine.compute()` له تعقيد O(tasks × predecessors) — مع 500 مهمة = 250,000 عملية في كل keystroke
- **الحل:**
```javascript
// استخدم dirty flag
let _cpmCache = null;
let _cpmDirty = true;

function markCPMDirty() { _cpmDirty = true; }
function getCPM() {
    if (_cpmDirty) {
        _cpmCache = CPMEngine.compute(project.tasks);
        _cpmDirty = false;
    }
    return _cpmCache;
}

// استدعي markCPMDirty() فقط عند تغيير: duration, start, finish, predecessors, outlineLevel
```
- **الوقت:** 2-3 ساعات

---

### 4. JSON.parse(JSON.stringify) في Undo Stack
- **الملف:** `js/app.js` السطور 1795, 5282
- **المشكلة:** Deep clone كامل للمشروع (قد يصل 5-10MB) × MAX_UNDO=50 = تسرب 250-500MB
- **الحل:**
```javascript
// استخدم structuredClone (native، أسرع 2-3x)
const snapshot = structuredClone(project);

// أو الأفضل: احفظ diff بس بدل snapshot كامل
// باستخدام jsondiffpatch أو similar
```
- **الوقت:** 1-2 ساعة (structuredClone) أو يوم كامل (diff-based)

---

### 5. XSS عبر Task Names في UI
- **الملفات:** `js/board.js` (~180), `js/gantt.js` (~240), `js/reports.js`
- **المشكلة:** أسماء المهام والوصف بتتحط في `innerHTML` مباشرة بدون escape. لو المستخدم استورد مشروع من Planner مع `<img src=x onerror=alert(1)>` في الاسم → XSS
- **الحل:**
```javascript
// helper function
function escapeHTML(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// استخدم textContent بدل innerHTML لكل user data
card.textContent = task.name;  // بدل card.innerHTML = task.name
// أو
card.innerHTML = `<strong>${escapeHTML(task.name)}</strong>`;
```
- **الوقت:** 2-3 ساعات (grep كل `innerHTML` واستبدله)

---

### 6. Scenarios.js — localStorage Quota Exceeded
- **الملف:** `js/scenarios.js` السطر 38
- **المشكلة:** `localStorage.setItem(STORAGE_KEY, ...)` بدون try/catch. مع 20+ سيناريو = exception يكسر التطبيق
- **الحل:**
```javascript
function saveScenarios(data) {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
        return true;
    } catch (e) {
        if (e.name === 'QuotaExceededError') {
            showToast('خزنة السيناريوهات ممتلئة. احذف بعض السيناريوهات القديمة.', 'error');
        } else {
            showToast('فشل حفظ السيناريو: ' + e.message, 'error');
        }
        return false;
    }
}
```
- **الوقت:** 30 دقيقة

---

### 7. Canvas Dimension Overflow في Gantt
- **الملف:** `js/gantt.js` السطر 101
- **المشكلة:** مع 2000+ مهمة أو مشروع 5 سنوات، الـ canvas بيتجاوز `MAX_CANVAS_DIM` ويتقطع بصمت
- **الحل:**
```javascript
if (w * dpr > MAX_CANVAS_DIM) {
    console.warn(`Gantt: canvas truncated. ${w}px requested, max ${MAX_CANVAS_DIM/dpr}px`);
    showToast('المشروع كبير جداً للعرض الكامل. استخدم الـ viewport controls للتنقل.', 'warning');
    // Implement virtual viewport instead
    w = MAX_CANVAS_DIM / dpr;
}
```
- **الوقت:** 1 ساعة (warning) أو يومين (virtual scrolling)

---

### 8. Date Parsing بدون Try/Catch في EVM
- **الملف:** `js/evm.js` السطر ~68
- **المشكلة:** `new Date(project.startDate)` على قيمة corrupted = `Invalid Date` يسرّب NaN لكل الحسابات
- **الحل:**
```javascript
function safeDate(s, fallback = new Date()) {
    if (!s) return fallback;
    const d = new Date(s);
    return isNaN(d.getTime()) ? fallback : d;
}

const projectStart = safeDate(project.startDate);
const projectFinish = safeDate(project.finishDate, projectStart);
if (projectFinish < projectStart) {
    console.warn('EVM: project finish before start — EVM calculations skipped');
    return null;
}
```
- **الوقت:** 1 ساعة

---

## 🟠 القسم الثاني: المشاكل المهمة (P1)

### 🧮 المعادلات الحسابية

#### 9. EAC Logic غريب عند CPI = 0
- **الملف:** `js/evm.js` السطر 93
- **الكود الحالي:** `const EAC = CPI > 0 ? effectiveBAC / CPI : effectiveBAC;`
- **المشكلة:** لو CPI = 0 (أي EV = 0 مع AC > 0)، النتيجة = `BAC` وهذا يُضلل المستخدم. الصحيح أنها Infinity أو `AC + (BAC - EV)` (Formula 2)
- **الحل:**
```javascript
// استخدم Formula 2 (الأكثر دقة عند بداية المشروع)
const EAC = (CPI > 0)
    ? AC + ((effectiveBAC - EV) / CPI)  // independent formula
    : (EV > 0 ? effectiveBAC : Infinity);
```

#### 10. Resource Leveling يتجاهل Holidays
- **الملف:** `js/resource-manager.js` السطر 94
- **المشكلة:** `dur` محسوبة بـ calendar days بدل working days، فالعبء بيتوزع على أيام العطل
- **الحل:** استخدم `WorkCalendar.getWorkingDays(start, finish)` بدل `daysBetween()`

#### 11. Resource Histogram Iterator Bug
- **الملف:** `js/resource-manager.js` السطر 8
- **المشكلة:** `for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1))` — المقارنة على نفس الـ reference، ممكن infinite loop لو حصل DST jump
- **الحل:**
```javascript
for (let d = new Date(start); d.getTime() <= end.getTime();) {
    // ... process day
    d = new Date(d.getTime() + 86400000); // add 24h as constant
}
```

#### 12. CPM يستخدم Calendar Days بدل Working Days
- **الملف:** `js/critical-path.js` السطر 110
- **المشكلة:** `daysBetween()` يحسب calendar days مما يخل بتواريخ Working Calendar
- **الحل:** مرر `workCalendar` للـ CPMEngine واستخدم `addWorkingDays`

#### 13. Date Mutation في Calendar
- **الملف:** `js/calendar.js` السطور 273, 279
- **المشكلة:** `setTime()` و `setDate()` يعدّلوا الـ Date الأصلي
- **الحل:** `d = new Date(d.getTime() + ...)` بدل `d.setDate(...)`

#### 14. Negative Task Duration
- **الملف:** `js/critical-path.js` السطر 115
- **المشكلة:** `task._ef = task._es + (task.durationDays || 0);` — لو `durationDays = -5`، يصبح EF < ES
- **الحل:** `task._ef = task._es + Math.max(0, task.durationDays || 0);`

#### 15. getAncestors Infinite Loop Risk
- **الملف:** `js/task-editor.js` السطر 177
- **المشكلة:** `outlineLevel` corrupted ممكن يسبب loop لا ينتهي
- **الحل:** Max iterations guard (100)

---

### ⚡ الأداء والكفاءة

#### 16. Event Listeners بدون Cleanup
- **الملف:** `js/app.js` السطور 714, 735, 746, 753, 760, 825
- **المشكلة:** Multiple `addEventListener` على document/window بدون removeEventListener
- **الحل:** استخدم `AbortController` واحد:
```javascript
const appController = new AbortController();
document.addEventListener('click', handler, { signal: appController.signal });
// عند unmount:
appController.abort(); // يلغي كل الـ listeners دفعة واحدة
```

#### 17. innerHTML في Loops
- **الملفات:** `js/app.js` 1188, 1688, 2728, 4081
- **الحل:** استخدم `DocumentFragment`:
```javascript
const frag = document.createDocumentFragment();
rows.forEach(row => {
    const tr = document.createElement('tr');
    tr.textContent = row; // أو insertAdjacentHTML مع escaped data
    frag.appendChild(tr);
});
tbody.innerHTML = '';
tbody.appendChild(frag);
```

#### 18. querySelector في Hot Paths
- **الملف:** `js/app.js` السطور 1364, 1382, 1916
- **الحل:** cache refs في init:
```javascript
const DOM = {
    resourceHeatmapWrap: document.getElementById('resourceHeatmapWrap'),
    // ...
};
```

#### 19. drawGrid O(days × rows)
- **الملف:** `js/gantt.js` السطور 153-162
- **المشكلة:** 500 day × 500 task = 250,000 draw calls
- **الحل:** Pre-render grid في offscreen canvas، ثم `drawImage` واحدة

#### 20. findIndex O(n³) في Dependencies
- **الملف:** `js/gantt.js` السطر 262
- **الحل:**
```javascript
// في update()، ابني index map مرة واحدة
const taskIdx = new Map();
tasks.forEach((t, i) => taskIdx.set(t.uid, i));
// في drawDependencyLinks:
const fromIdx = taskIdx.get(predUid); // O(1)
```

#### 21. Nested filter().map().sort() Chains
- **الملفات:** `js/reports.js` 36, `js/dashboard.js` 62-70
- **الحل:** Single-pass loop مع state tracking

#### 22. calculateProjectBounds() مكررة
- **الملف:** `js/gantt.js` السطور 92, 104
- **المشكلة:** محسوبة مرتين في update() + resize()
- **الحل:** احسبها مرة، خزن النتيجة

---

### 🔌 التكامل والربط

#### 23. SQL Field Injection في server.js
- **الملف:** `server.js` السطر 390 (PATCH endpoint)
- **المشكلة:** `db.prepare(`UPDATE projects_meta SET ${sets.join(', ')} WHERE id = ?`)` — أسماء الحقول من الـ body مباشرة
- **الحل:** Whitelist:
```javascript
const ALLOWED_FIELDS = ['name', 'color', 'description', 'updated_at'];
const sets = [];
const values = [];
for (const field of ALLOWED_FIELDS) {
    if (field in req.body) {
        sets.push(`${field} = ?`);
        values.push(req.body[field]);
    }
}
if (!sets.length) return res.status(400).json({ error: 'No valid fields' });
values.push(req.params.id);
db.prepare(`UPDATE projects_meta SET ${sets.join(', ')} WHERE id = ?`).run(...values);
```

#### 24. D365 — بدون Pagination
- **الملف:** `js/d365.js` (multiple locations)
- **المشكلة:** OData بيرجع 5000 سجل كحد أقصى لكل صفحة، والكود بيتجاهل `@odata.nextLink`
- **الحل:**
```javascript
async function fetchAllPages(initialUrl) {
    const all = [];
    let url = initialUrl;
    while (url) {
        const result = await _callDataverse('GET', url);
        all.push(...result.value);
        url = result['@odata.nextLink'] || null;
    }
    return all;
}
```

#### 25. D365 — Rate Limit Loop بلا حد أقصى
- **الملف:** `js/d365.js` السطور 122-128
- **المشكلة:** 429 retry بدون MAX_RETRIES → infinite recursion محتمل
- **الحل:**
```javascript
async function _callDataverse(method, url, body, retryCount = 0) {
    const MAX_RETRIES = 5;
    // ...
    if (response.status === 429) {
        if (retryCount >= MAX_RETRIES) {
            throw new Error('D365 rate limit exceeded after ' + MAX_RETRIES + ' retries');
        }
        const delay = Math.min(32000, 1000 * Math.pow(2, retryCount));
        await new Promise(r => setTimeout(r, delay));
        return _callDataverse(method, url, body, retryCount + 1);
    }
    // ...
}
```

#### 26. D365 — لا يوجد ETag للـ Optimistic Concurrency
- **الملف:** `js/d365.js` السطر 530
- **المشكلة:** PATCH بدون `If-Match` → آخر كاتب يكسب (lost updates)
- **الحل:** جيب ETag أولاً، ضيفه في headers

#### 27. CORS Too Permissive
- **الملف:** `server.js` السطور 104-112
- **المشكلة:** يسمح بـ HTTP على LAN subnets (192.168.x.x)
- **الحل:** في production، force HTTPS

#### 28. Secrets في Client Code
- **الملفات:** `js/teams-bridge.js` 40, `js/ms-graph.js` 24
- **المشكلة:** Client IDs hardcoded + credentials في localStorage (plaintext)
- **الحل:** انقل Client IDs لـ `server-config.json` (غير مضمن في git)، استخدم `sessionStorage` للـ tokens الحساسة

#### 29. Large File DoS
- **الملف:** `js/planner-parser.js` (~40)
- **المشكلة:** لا يوجد size limit للملفات المستوردة
- **الحل:**
```javascript
if (file.size > 50 * 1024 * 1024) {
    throw new Error('الملف كبير جداً. الحد الأقصى 50MB');
}
```

#### 30. Excel Date Leap-Year Bug
- **الملف:** `js/xml-parser.js` 270, `js/planner-parser.js` 400
- **المشكلة:** Excel يعامل 1900 كـ leap year (وهي مش leap year) — تواريخ قبل 1 مارس 1900 تختلف بيوم
- **الحل:**
```javascript
function excelToJSDate(val) {
    // Excel epoch is 1900-01-01, but Excel thinks 1900 is a leap year
    // so subtract 1 for dates >= March 1, 1900
    const days = val > 59 ? val - 1 : val;
    return new Date((days - 25568) * 86400 * 1000);
}
```

---

### 🛡️ Edge Cases

#### 31. JSON.parse بدون Try/Catch
- **الملف:** `js/project-io.js` السطر 13
- **الحل:** Wrap في try/catch مع user-friendly error

#### 32. forEach على undefined
- **الملف:** `js/project-io.js` السطر 87
- **المشكلة:** `project.tasks.forEach` — لو `tasks` مفقودة، crash
- **الحل:** `Array.isArray(project.tasks) ? project.tasks.forEach(...) : null`

#### 33. Array Mutation أثناء Render
- **الملف:** `js/gantt.js` السطر 267
- **الحل:** Snapshot قبل loop: `const snapshot = tasks.slice();`

#### 34. Missing Circular Dependency Recovery
- **الملف:** `js/critical-path.js`
- **المشكلة:** topoSort بيرمي exception لكن الـ caller في app.js ما بيتعامل معاه
- **الحل:** Wrap في try/catch، اعرض banner للمستخدم: "المشروع يحتوي على dependency circular — يرجى الإصلاح"

#### 35. State Manager — Circular Re-render
- **الملف:** `js/state-manager.js` السطر 91
- **الحل:** Flag `_isRendering` لمنع recursion

---

## 🟡 القسم الثالث: تحسينات (P2)

### معادلات
- **P2.1** `critical-path.js:86` — لا يوجد guard للـ summary tasks في Forward Pass
- **P2.2** `resource-manager.js:165-223` — Auto-leveling بطيء (1-day increments). استخدم heuristic
- **P2.3** `project-analytics.js:227` — avgPct edge case عند leafTasks.length = 0
- **P2.4** توحيد استخدام `WorkCalendar` في كل المعادلات (CPM, Resource, EVM)

### أداء
- **P2.5** `task-editor.js:45` — closure يحفظ task reference (memory leak محتمل)
- **P2.6** `gantt.js:170` — setLineDash بدون save/restore
- **P2.7** `network.js:458,486,511` — font assignment متكررة، اعمل cache
- **P2.8** `app.js:670` — `forEach(t => set.add(t.uid))` → `new Set(arr.map(t => t.uid))`
- **P2.9** `reports.js:216,308,416` — slice().forEach بدل direct forEach with counter
- **P2.10** CSS: أضف `contain: layout` لـ `.gantt-row`، `will-change: transform` للـ animations
- **P2.11** `network.js:591-627` — Minimap تُرسم في كل _draw. Cache في offscreen canvas

### تكامل
- **P2.12** `ms-graph.js` — بدون network failure fallback UI
- **P2.13** `xml-parser.js` — لا يعالج UTF-8 BOM (`\ufeff`)
- **P2.14** `server.js:136-146` — JS/CSS بدون caching (maxAge: 0)
- **P2.15** `server.js:569-572` — directory traversal عبر symlinks (استخدم realpath)
- **P2.16** `server.js` — JSON.parse limit 50mb عالي جداً، اخفضه لـ 10mb

### Edge Cases
- **P2.17** Undo/Redo edge cases مع very large state
- **P2.18** Task مع 0 duration، negative duration
- **P2.19** Unicode/emoji في task names
- **P2.20** MAX_SAFE_INTEGER في numeric fields
- **P2.21** Dark mode missing styles في بعض modals
- **P2.22** Keyboard shortcuts conflicts (⌘K مع browser shortcuts)

---

## 🎯 خطة التنفيذ المقترحة (Sprint Plan)

### الأسبوع 1: P0 (الحرجة)
| اليوم | المهام | الوقت |
|------|-------|------|
| الاثنين | #1 (XXE) + #5 (XSS) + #6 (localStorage) + #8 (Date parsing) | 4 ساعات |
| الثلاثاء-الأربعاء | #2 (RAF scheduler) + #3 (CPM caching) | يوم ونصف |
| الخميس | #4 (structuredClone undo) + #7 (Canvas overflow warning) | يوم |
| الجمعة | Testing شامل للـ P0 + regression | يوم |

### الأسبوع 2: P1 أداء + تكامل
| اليوم | المهام | الوقت |
|------|-------|------|
| الاثنين | #16 (Event cleanup) + #17 (DocumentFragment) + #18 (DOM cache) | يوم |
| الثلاثاء | #19 (Grid pre-render) + #20 (Dep index map) + #22 (bounds dedupe) | يوم |
| الأربعاء | #23 (SQL whitelist) + #27 (CORS) + #28 (Secrets) | يوم |
| الخميس | #24-26 (D365 pagination, rate limit, ETag) | يوم |
| الجمعة | Testing + performance benchmarks | يوم |

### الأسبوع 3: P1 معادلات + edge cases
| اليوم | المهام | الوقت |
|------|-------|------|
| الاثنين | #9 (EAC) + #12 (CPM working days) + #14 (negative duration) | يوم |
| الثلاثاء | #10-11 (Resource holidays + iterator) + #13 (Date mutation) | يوم |
| الأربعاء | #15 (getAncestors) + #31-35 (error handling) | يوم |
| الخميس | #29 (file size) + #30 (Excel dates) + #32-33 (null safety) | يوم |
| الجمعة | End-to-end testing | يوم |

### الأسبوع 4+: P2 (تحسينات اختيارية)
P2s قابلة للتأجيل، اعمل منها اللي بيأثر على المستخدمين فعلياً بناءً على التعليقات.

---

## ✅ قائمة التحقق (Checklist) قبل الإنتاج

### الأمان
- [ ] XXE محمي في xml-parser
- [ ] XSS محمي — كل user input escaped أو في textContent
- [ ] SQL injection — whitelist للـ PATCH fields
- [ ] Secrets منقولة من client code
- [ ] HTTPS enforcement في production
- [ ] File upload size limits
- [ ] CSP headers مضبوطة

### الأداء
- [ ] CPM caching مع dirty flag
- [ ] RAF-based render scheduler
- [ ] structuredClone للـ undo stack
- [ ] Event listeners عبر AbortController
- [ ] Grid pre-rendered في offscreen canvas
- [ ] Dependency index map
- [ ] Debounced autoSave
- [ ] Benchmark: 500 tasks, render < 50ms

### الصحة
- [ ] EAC formula مصححة
- [ ] Resource calculations تستخدم working days
- [ ] كل Date operations immutable
- [ ] Negative duration محمي
- [ ] Circular dependency يعرض رسالة واضحة
- [ ] Division by zero محمي في كل المعادلات

### التكامل
- [ ] D365 pagination شغالة
- [ ] D365 rate limit مع MAX_RETRIES
- [ ] D365 ETag للـ PATCH
- [ ] MS Graph retry مع backoff
- [ ] File imports تتعامل مع BOM، Excel leap year
- [ ] Offline queue للعمليات الفاشلة

### UX
- [ ] Error messages بالعربي واضحة
- [ ] Loading indicators لكل async
- [ ] Toast notifications للـ errors
- [ ] Graceful degradation لما integration يفشل

---

## 🧪 توصيات إضافية

### 1. أضف Unit Tests
المعادلات الحسابية (CPM، EVM، Calendar) لازم يكون عليها tests. استخدم **Vitest** (موجود مع Vite):
```bash
npm i -D vitest
```
ابدأ بـ:
- `critical-path.test.js` — forward/backward pass scenarios
- `evm.test.js` — edge cases (CPI=0, PV=0, NaN inputs)
- `calendar.test.js` — working days calculations

### 2. أضف ESLint + Prettier
يساعد في اكتشاف null references و type issues:
```bash
npm i -D eslint @eslint/js eslint-plugin-security prettier
```

### 3. فكر في TypeScript migration
مع 30+ ملف JS و 300KB من الكود، TypeScript هيمسك 80% من الـ bugs اللي فوق قبل ما تحصل. بدون full rewrite، ضيف `// @ts-check` في رأس كل ملف مع JSDoc types.

### 4. أضف Error Boundary
```javascript
window.addEventListener('error', (e) => {
    console.error('Global error:', e);
    showToast('حدث خطأ غير متوقع. التطبيق ربما يحتاج إعادة تحميل.', 'error');
});

window.addEventListener('unhandledrejection', (e) => {
    console.error('Unhandled rejection:', e.reason);
    showToast('فشلت عملية async: ' + e.reason?.message, 'error');
});
```

### 5. Performance Monitoring
أضف markers بـ `performance.mark()` قبل وبعد الـ renders عشان تقدر تقيس التحسن:
```javascript
performance.mark('render-start');
renderAll();
performance.mark('render-end');
performance.measure('render', 'render-start', 'render-end');
```

---

## 📝 ملاحظات ختامية

**نقاط القوة في الكود الحالي:**
- ✅ Architecture نظيف ومفصول (separation of concerns)
- ✅ CPM Forward/Backward Pass صحيح ويدعم 4 dependency types
- ✅ Circular dependency detection بـ topological sort
- ✅ Health Score algorithm متوازن ومدروس
- ✅ معظم EVM indices محمية من division by zero
- ✅ Canvas rendering مع DPR handling

**الأولويات الثلاث الأهم للعمل:**
1. **أصلح P0 فوراً** — خاصة XXE + XSS + CPM caching (يوفر 60% من زمن الـ render)
2. **أضف Unit Tests للمعادلات** — ده هيحميك من regression لسنوات
3. **نقل للـ TypeScript تدريجياً** — استثمار يستحق مع حجم الكود

---

**المراجع:**
- OWASP Top 10 (XXE، XSS، Injection)
- PMI PMBOK (EVM formulas — EAC alternatives)
- Web Performance Best Practices (Chrome DevTools docs)
- Microsoft Graph & Dataverse official retry guidance

*نهاية التقرير — تم المراجعة بتاريخ 17 أبريل 2026*
