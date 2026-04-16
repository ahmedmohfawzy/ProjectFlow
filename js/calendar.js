/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Work Calendar & Public Holidays
 * Manages working days, holidays, and schedule computation
 * Uses Nager.Date free API for public holidays
 * ═══════════════════════════════════════════════════════
 */


    // ─── Presets ───
    const PRESETS = {
        'western':    { name: 'Western (Mon-Fri)',      workDays: [1,2,3,4,5], weekendDays: [0,6] },
        'egypt':      { name: 'Egypt (Sun-Thu)',        workDays: [0,1,2,3,4], weekendDays: [5,6] },
        'saudi':      { name: 'Saudi Arabia (Sun-Thu)', workDays: [0,1,2,3,4], weekendDays: [5,6] },
        'uae':        { name: 'UAE (Mon-Fri)',          workDays: [1,2,3,4,5], weekendDays: [0,6] },
        'uae_old':    { name: 'UAE Old (Sun-Thu)',      workDays: [0,1,2,3,4], weekendDays: [5,6] },
        'custom':     { name: 'Custom',                 workDays: [1,2,3,4,5], weekendDays: [0,6] }
    };

    // ─── Countries for holidays API ───
    const COUNTRIES = [
        { code: 'EG', name: 'مصر / Egypt', flag: '🇪🇬' },
        { code: 'SA', name: 'السعودية / Saudi Arabia', flag: '🇸🇦' },
        { code: 'AE', name: 'الإمارات / UAE', flag: '🇦🇪' },
        { code: 'US', name: 'United States', flag: '🇺🇸' },
        { code: 'GB', name: 'United Kingdom', flag: '🇬🇧' },
        { code: 'DE', name: 'Germany', flag: '🇩🇪' },
        { code: 'FR', name: 'France', flag: '🇫🇷' },
        { code: 'CN', name: 'China', flag: '🇨🇳' },
        { code: 'IN', name: 'India', flag: '🇮🇳' },
        { code: 'JP', name: 'Japan', flag: '🇯🇵' },
        { code: 'KR', name: 'South Korea', flag: '🇰🇷' },
        { code: 'TR', name: 'Turkey', flag: '🇹🇷' },
        { code: 'BR', name: 'Brazil', flag: '🇧🇷' },
        { code: 'MX', name: 'Mexico', flag: '🇲🇽' },
        { code: 'AU', name: 'Australia', flag: '🇦🇺' },
        { code: 'CA', name: 'Canada', flag: '🇨🇦' },
        { code: 'JO', name: 'الأردن / Jordan', flag: '🇯🇴' },
        { code: 'KW', name: 'الكويت / Kuwait', flag: '🇰🇼' },
        { code: 'QA', name: 'قطر / Qatar', flag: '🇶🇦' },
        { code: 'BH', name: 'البحرين / Bahrain', flag: '🇧🇭' },
        { code: 'OM', name: 'عمان / Oman', flag: '🇴🇲' },
        { code: 'MA', name: 'المغرب / Morocco', flag: '🇲🇦' },
        { code: 'TN', name: 'تونس / Tunisia', flag: '🇹🇳' },
    ];

    const NAGER_API = 'https://date.nager.at/api/v3';
    const DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const DAY_NAMES_AR = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];

    // State
    let config = {
        preset: 'western',
        workDays: [1, 2, 3, 4, 5],
        weekendDays: [0, 6],
        hoursPerDay: 8,
        startHour: 9,
        endHour: 17,
        countries: [],           // selected country codes for holidays
        holidays: {},            // { '2026': { '2026-01-07': {name, localName, country} } }
        customHolidays: [],      // [{date, name}] user-added
    };

    let _holidayCache = {};

    // ─── Init ───
    function init(savedConfig) {
        if (savedConfig) {
            Object.assign(config, savedConfig);
        }
        return config;
    }

    function getConfig() { return { ...config }; }
    function getPresets() { return PRESETS; }
    function getCountries() { return COUNTRIES; }
    function getDayNames() { return DAY_NAMES; }

    // ─── Set Work Calendar ───
    function setPreset(presetKey) {
        const preset = PRESETS[presetKey];
        if (preset) {
            config.preset = presetKey;
            config.workDays = [...preset.workDays];
            config.weekendDays = [...preset.weekendDays];
        }
        save();
        return config;
    }

    function setWorkDays(days) {
        config.workDays = days;
        config.weekendDays = [0,1,2,3,4,5,6].filter(d => !days.includes(d));
        config.preset = 'custom';
        save();
        return config;
    }

    function setHoursPerDay(hours) {
        config.hoursPerDay = hours;
        save();
    }

    // ─── Holiday Countries ───
    function setCountries(countryCodes) {
        config.countries = countryCodes;
        save();
    }

    // ─── Fetch Holidays from API ───
    async function fetchHolidays(year) {
        if (!config.countries || config.countries.length === 0) return {};

        const cacheKey = year + '_' + config.countries.sort().join(',');
        if (_holidayCache[cacheKey]) return _holidayCache[cacheKey];

        const allHolidays = {};

        for (const countryCode of config.countries) {
            try {
                const url = `${NAGER_API}/PublicHolidays/${year}/${countryCode}`;
                const resp = await fetch(url);
                if (!resp.ok) continue;
                const data = await resp.json();

                for (const h of data) {
                    const key = h.date;
                    if (!allHolidays[key]) {
                        allHolidays[key] = {
                            date: h.date,
                            name: h.name,
                            localName: h.localName || h.name,
                            country: countryCode,
                            countries: [countryCode],
                            flag: COUNTRIES.find(c => c.code === countryCode)?.flag || ''
                        };
                    } else {
                        if (!allHolidays[key].countries.includes(countryCode)) {
                            allHolidays[key].countries.push(countryCode);
                            allHolidays[key].name += ' / ' + h.name;
                        }
                    }
                }
            } catch (err) {
                console.warn(`Failed to fetch holidays for ${countryCode}/${year}:`, err);
            }
        }

        // Add custom holidays
        for (const ch of config.customHolidays) {
            if (ch.date && ch.date.startsWith(String(year))) {
                allHolidays[ch.date] = {
                    date: ch.date, name: ch.name, localName: ch.name,
                    country: 'custom', countries: ['custom'], flag: '📌'
                };
            }
        }

        // Store
        if (!config.holidays) config.holidays = {};
        config.holidays[year] = allHolidays;
        _holidayCache[cacheKey] = allHolidays;
        save();

        return allHolidays;
    }

    /**
     * Preload holidays for project date range
     */
    async function preloadForProject(startDate, endDate) {
        const startYear = new Date(startDate).getFullYear();
        const endYear = new Date(endDate).getFullYear();
        const results = {};

        for (let y = startYear; y <= endYear; y++) {
            const h = await fetchHolidays(y);
            Object.assign(results, h);
        }

        return results;
    }

    // ─── Custom Holidays ───
    function addCustomHoliday(date, name) {
        config.customHolidays.push({ date, name });
        // Clear cache for that year
        const year = date.substring(0, 4);
        // Clear all cache keys containing this year (including country-specific keys)
        Object.keys(_holidayCache).forEach(k => { if (k.startsWith(year)) delete _holidayCache[k]; });
        if (config.holidays && config.holidays[year]) {
            config.holidays[year][date] = {
                date, name, localName: name,
                country: 'custom', countries: ['custom'], flag: '📌'
            };
        }
        save();
    }

    function removeCustomHoliday(date) {
        config.customHolidays = config.customHolidays.filter(h => h.date !== date);
        const year = date.substring(0, 4);
        delete _holidayCache[year];
        if (config.holidays?.[year]?.[date]?.country === 'custom') {
            delete config.holidays[year][date];
        }
        save();
    }

    // ─── Core Calendar Logic ───

    /**
     * Is a given date a working day?
     */
    function isWorkingDay(date) {
        const d = new Date(date);
        const dow = d.getDay();

        // Check if it's a weekend day
        if (config.weekendDays.includes(dow)) return false;

        // Check if it's a holiday
        const key = toDateKey(d);
        const year = d.getFullYear();
        if (config.holidays?.[year]?.[key]) return false;

        return true;
    }

    /**
     * Is a given date a public holiday?
     */
    function isHoliday(date) {
        const d = new Date(date);
        const key = toDateKey(d);
        const year = d.getFullYear();
        return config.holidays?.[year]?.[key] || null;
    }

    /**
     * Is a given date a weekend?
     */
    function isWeekend(date) {
        const d = new Date(date);
        return config.weekendDays.includes(d.getDay());
    }

    /**
     * Get the number of working days between two dates
     */
    function getWorkingDays(startDate, endDate) {
        let count = 0;
        const start = new Date(startDate); start.setHours(0, 0, 0, 0);
        const end = new Date(endDate); end.setHours(0, 0, 0, 0);

        for (let d = new Date(start); d < end; d.setTime(d.getTime() + 86400000)) {
            if (isWorkingDay(d)) count++;
        }

        return count;
    }

    /**
     * Add N working days to a date (skip weekends + holidays)
     */
    function addWorkingDays(startDate, numDays) {
        const d = new Date(startDate); d.setHours(0, 0, 0, 0);
        let remaining = numDays;
        const MAX_ITER = 10000; // Guard: prevent infinite loop
        let iter = 0;

        while (remaining > 0 && iter++ < MAX_ITER) {
            d.setTime(d.getTime() + 86400000);
            if (isWorkingDay(d)) remaining--;
        }
        if (iter >= MAX_ITER) console.warn('[Calendar] addWorkingDays hit max iterations');

        return d;
    }

    /**
     * Get next working day on or after the given date
     */
    function getNextWorkingDay(date) {
        const d = new Date(date); d.setHours(0, 0, 0, 0);
        const MAX_ITER = 10000; // Guard: prevent infinite loop
        let iter = 0;
        while (!isWorkingDay(d) && iter++ < MAX_ITER) {
            d.setTime(d.getTime() + 86400000);
        }
        if (iter >= MAX_ITER) console.warn('[Calendar] getNextWorkingDay hit max iterations');
        return d;
    }

    /**
     * Get all holidays + weekends in a date range (for Gantt shading)
     */
    function getNonWorkingDays(startDate, endDate) {
        const result = { weekends: [], holidays: [] };
        const start = new Date(startDate); start.setHours(0, 0, 0, 0);
        const end = new Date(endDate); end.setHours(0, 0, 0, 0);

        for (let d = new Date(start); d <= end; d.setTime(d.getTime() + 86400000)) {
            const key = toDateKey(d);
            const dow = d.getDay();

            if (config.weekendDays.includes(dow)) {
                result.weekends.push({ date: new Date(d), dayOfWeek: dow });
            }

            const year = d.getFullYear();
            const hol = config.holidays?.[year]?.[key];
            if (hol) {
                result.holidays.push({ date: new Date(d), ...hol });
            }
        }

        return result;
    }

    /**
     * Get a calendar month view with working/non-working classified
     */
    function getMonthView(year, month) {
        const days = [];
        const firstDay = new Date(year, month, 1);
        const lastDay = new Date(year, month + 1, 0);

        for (let d = 1; d <= lastDay.getDate(); d++) {
            const date = new Date(year, month, d);
            const key = toDateKey(date);
            const dow = date.getDay();
            const hol = config.holidays?.[year]?.[key];

            days.push({
                date: new Date(date),
                day: d,
                dayOfWeek: dow,
                dayName: DAY_NAMES[dow],
                isWeekend: config.weekendDays.includes(dow),
                isHoliday: !!hol,
                holidayInfo: hol || null,
                isWorking: isWorkingDay(date)
            });
        }

        return days;
    }

    // ─── Helpers ───
    function toDateKey(d) {
        const dt = new Date(d);
        const y = dt.getFullYear();
        const m = String(dt.getMonth() + 1).padStart(2, '0');
        const day = String(dt.getDate()).padStart(2, '0');
        return `${y}-${m}-${day}`;
    }

    function save() {
        try {
            localStorage.setItem('pf_calendar', JSON.stringify(config));
        } catch (e) { console.warn('[Calendar] Save failed:', e.message); }
    }

    function load() {
        try {
            const saved = localStorage.getItem('pf_calendar');
            if (saved) Object.assign(config, JSON.parse(saved));
        } catch (e) { console.warn('[Calendar] Load failed:', e.message); }
        return config;
    }

    export const WorkCalendar = {
        init, load, save, getConfig, getPresets, getCountries, getDayNames,
        setPreset, setWorkDays, setHoursPerDay,
        setCountries, fetchHolidays, preloadForProject,
        addCustomHoliday, removeCustomHoliday,
        isWorkingDay, isWeekend, isHoliday,
        getWorkingDays, addWorkingDays, getNextWorkingDay,
        getNonWorkingDays, getMonthView,
        PRESETS, COUNTRIES
    };
