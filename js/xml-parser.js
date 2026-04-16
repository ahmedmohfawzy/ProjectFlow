/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — MS Project XML (MSPDI) Parser
 * Handles import & export of Microsoft Project XML format
 * ═══════════════════════════════════════════════════════
 */


    // ——— MS Project XML Namespace ———
    const NS = 'http://schemas.microsoft.com/project';

    /**
     * Parse MS Project XML string → project data object
     * @param {string} xmlString - Raw XML content
     * @returns {object} Parsed project data
     */
    function parse(xmlString) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlString, 'text/xml');

        // Check for parse errors
        const parseError = doc.querySelector('parsererror');
        if (parseError) {
            throw new Error('Invalid XML file: ' + parseError.textContent.substring(0, 200));
        }

        // Determine if namespaced
        const root = doc.documentElement;
        const isNamespaced = root.namespaceURI === NS;

        // Helper to read element text
        const getText = (parent, tagName) => {
            let el;
            if (isNamespaced) {
                el = parent.getElementsByTagNameNS(NS, tagName)[0];
            } else {
                el = parent.getElementsByTagName(tagName)[0];
            }
            return el ? (el.textContent || '').trim() : '';
        };

        // Read project-level properties
        const project = {
            name: getText(root, 'Name') || getText(root, 'Title') || 'Untitled Project',
            manager: getText(root, 'Manager') || '',
            startDate: getText(root, 'StartDate') || new Date().toISOString(),
            finishDate: getText(root, 'FinishDate') || '',
            creationDate: getText(root, 'CreationDate') || new Date().toISOString(),
            calendarUID: getText(root, 'CalendarUID') || '1',
            minutesPerDay: parseInt(getText(root, 'MinutesPerDay')) || 480,
            minutesPerWeek: parseInt(getText(root, 'MinutesPerWeek')) || 2400,
            daysPerMonth: parseInt(getText(root, 'DaysPerMonth')) || 20,
            currencySymbol: getText(root, 'CurrencySymbol') || '$',
            tasks: [],
            resources: [],
            assignments: [],
            calendars: []
        };

        // ——— Parse Resources ———
        const resourceNodes = isNamespaced
            ? root.getElementsByTagNameNS(NS, 'Resource')
            : root.getElementsByTagName('Resource');

        for (const resNode of resourceNodes) {
            const uid = getText(resNode, 'UID');
            if (uid === '0') continue; // Skip UID 0 (unassigned)
            project.resources.push({
                uid: parseInt(uid),
                id: parseInt(getText(resNode, 'ID')) || 0,
                name: getText(resNode, 'Name'),
                type: parseInt(getText(resNode, 'Type')) || 1, // 1=Work, 0=Material
                maxUnits: parseFloat(getText(resNode, 'MaxUnits')) || 1,
                standardRate: parseFloat(getText(resNode, 'StandardRate')) || 0,
                cost: parseFloat(getText(resNode, 'Cost')) || 0
            });
        }

        // ——— Parse Tasks ———
        const taskNodes = isNamespaced
            ? root.getElementsByTagNameNS(NS, 'Task')
            : root.getElementsByTagName('Task');

        for (const taskNode of taskNodes) {
            const uid = parseInt(getText(taskNode, 'UID'));
            if (uid === 0) continue; // Skip the project summary task (UID 0)

            const task = {
                uid: uid,
                id: parseInt(getText(taskNode, 'ID')) || uid,
                name: getText(taskNode, 'Name') || 'New Task',
                wbs: getText(taskNode, 'WBS') || '',
                outlineLevel: parseInt(getText(taskNode, 'OutlineLevel')) || 1,
                outlineNumber: getText(taskNode, 'OutlineNumber') || '',
                start: parseDate(getText(taskNode, 'Start')),
                finish: parseDate(getText(taskNode, 'Finish')),
                duration: getText(taskNode, 'Duration') || 'PT8H0M0S',
                durationDays: 0,
                percentComplete: parseInt(getText(taskNode, 'PercentComplete')) || 0,
                summary: getText(taskNode, 'Summary') === '1' || getText(taskNode, 'Summary') === 'true',
                milestone: getText(taskNode, 'Milestone') === '1' || getText(taskNode, 'Milestone') === 'true',
                critical: getText(taskNode, 'Critical') === '1' || getText(taskNode, 'Critical') === 'true',
                cost: parseFloat(getText(taskNode, 'Cost')) || 0,
                notes: getText(taskNode, 'Notes') || '',
                predecessors: [],
                resourceNames: [],
                isExpanded: true,
                isVisible: true
            };

            // Parse duration to days
            task.durationDays = parseDurationToDays(task.duration, project.minutesPerDay);

            // Check milestone by duration
            if (task.durationDays === 0 && !task.summary) {
                task.milestone = true;
            }

            // Parse predecessor links
            const predNodes = isNamespaced
                ? taskNode.getElementsByTagNameNS(NS, 'PredecessorLink')
                : taskNode.getElementsByTagName('PredecessorLink');

            for (const predNode of predNodes) {
                task.predecessors.push({
                    predecessorUID: parseInt(getText(predNode, 'PredecessorUID')),
                    type: parseInt(getText(predNode, 'Type')) || 1, // 1=FS
                    lag: parseInt(getText(predNode, 'LinkLag')) || 0
                });
            }

            project.tasks.push(task);
        }

        // ——— Parse Assignments ———
        const assignNodes = isNamespaced
            ? root.getElementsByTagNameNS(NS, 'Assignment')
            : root.getElementsByTagName('Assignment');

        for (const assignNode of assignNodes) {
            const taskUID = parseInt(getText(assignNode, 'TaskUID'));
            const resUID = parseInt(getText(assignNode, 'ResourceUID'));
            if (resUID === 0) continue;

            project.assignments.push({
                taskUID: taskUID,
                resourceUID: resUID,
                units: parseFloat(getText(assignNode, 'Units')) || 1
            });

            // Map resource names to tasks
            const task = project.tasks.find(t => t.uid === taskUID);
            const resource = project.resources.find(r => r.uid === resUID);
            if (task && resource) {
                task.resourceNames.push(resource.name);
            }
        }

        // ——— Detect summary tasks ———
        detectSummaryTasks(project.tasks);

        return project;
    }

    /**
     * Detect summary tasks from outline levels
     */
    function detectSummaryTasks(tasks) {
        for (let i = 0; i < tasks.length; i++) {
            const current = tasks[i];
            const next = tasks[i + 1];
            if (next && next.outlineLevel > current.outlineLevel) {
                current.summary = true;
            }
        }
    }

    /**
     * Export project data → MS Project XML string
     * @param {object} project - Project data object
     * @returns {string} XML string
     */
    function exportXML(project) {
        const tasks = project.tasks;
        const resources = project.resources || [];

        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
        xml += '<Project xmlns="http://schemas.microsoft.com/project">\n';

        // Project properties
        xml += `  <Name>${escapeXml(project.name)}</Name>\n`;
        xml += `  <Title>${escapeXml(project.name)}</Title>\n`;
        xml += `  <Manager>${escapeXml(project.manager || '')}</Manager>\n`;
        xml += `  <CreationDate>${new Date().toISOString()}</CreationDate>\n`;
        xml += `  <StartDate>${formatDateISO(project.startDate)}</StartDate>\n`;
        xml += `  <FinishDate>${formatDateISO(project.finishDate || project.startDate)}</FinishDate>\n`;
        xml += `  <MinutesPerDay>${project.minutesPerDay || 480}</MinutesPerDay>\n`;
        xml += `  <MinutesPerWeek>${project.minutesPerWeek || 2400}</MinutesPerWeek>\n`;
        xml += `  <DaysPerMonth>${project.daysPerMonth || 20}</DaysPerMonth>\n`;
        xml += `  <CurrencySymbol>${escapeXml(project.currencySymbol || '$')}</CurrencySymbol>\n`;
        xml += `  <CalendarUID>1</CalendarUID>\n`;

        // Calendars
        xml += '  <Calendars>\n';
        xml += '    <Calendar>\n';
        xml += '      <UID>1</UID>\n';
        xml += '      <Name>Standard</Name>\n';
        xml += '      <IsBaseCalendar>1</IsBaseCalendar>\n';
        xml += '    </Calendar>\n';
        xml += '  </Calendars>\n';

        // Tasks
        xml += '  <Tasks>\n';
        // Project summary task (UID 0)
        xml += '    <Task>\n';
        xml += '      <UID>0</UID>\n';
        xml += '      <ID>0</ID>\n';
        xml += `      <Name>${escapeXml(project.name)}</Name>\n`;
        xml += '      <Type>1</Type>\n';
        xml += '      <IsNull>0</IsNull>\n';
        xml += '      <OutlineLevel>0</OutlineLevel>\n';
        xml += '      <Summary>1</Summary>\n';
        xml += '    </Task>\n';

        for (const task of tasks) {
            xml += '    <Task>\n';
            xml += `      <UID>${task.uid}</UID>\n`;
            xml += `      <ID>${task.id}</ID>\n`;
            xml += `      <Name>${escapeXml(task.name)}</Name>\n`;
            xml += `      <WBS>${escapeXml(task.wbs || '')}</WBS>\n`;
            xml += `      <OutlineLevel>${task.outlineLevel || 1}</OutlineLevel>\n`;
            xml += `      <OutlineNumber>${escapeXml(task.outlineNumber || task.wbs || '')}</OutlineNumber>\n`;
            xml += `      <Start>${formatDateISO(task.start)}</Start>\n`;
            xml += `      <Finish>${formatDateISO(task.finish)}</Finish>\n`;
            xml += `      <Duration>${daysToDuration(task.durationDays)}</Duration>\n`;
            xml += `      <DurationFormat>7</DurationFormat>\n`;
            xml += `      <PercentComplete>${task.percentComplete || 0}</PercentComplete>\n`;
            xml += `      <Summary>${task.summary ? 1 : 0}</Summary>\n`;
            xml += `      <Milestone>${task.milestone ? 1 : 0}</Milestone>\n`;
            xml += `      <Cost>${task.cost || 0}</Cost>\n`;
            if (task.notes) {
                xml += `      <Notes>${escapeXml(task.notes)}</Notes>\n`;
            }

            // Predecessors
            for (const pred of (task.predecessors || [])) {
                xml += '      <PredecessorLink>\n';
                xml += `        <PredecessorUID>${pred.predecessorUID}</PredecessorUID>\n`;
                xml += `        <Type>${pred.type || 1}</Type>\n`;
                xml += `        <LinkLag>${pred.lag || 0}</LinkLag>\n`;
                xml += '      </PredecessorLink>\n';
            }

            xml += '    </Task>\n';
        }
        xml += '  </Tasks>\n';

        // Resources
        if (resources.length > 0) {
            xml += '  <Resources>\n';
            xml += '    <Resource>\n';
            xml += '      <UID>0</UID>\n';
            xml += '      <ID>0</ID>\n';
            xml += '      <Type>1</Type>\n';
            xml += '    </Resource>\n';
            for (const res of resources) {
                xml += '    <Resource>\n';
                xml += `      <UID>${res.uid}</UID>\n`;
                xml += `      <ID>${res.id}</ID>\n`;
                xml += `      <Name>${escapeXml(res.name)}</Name>\n`;
                xml += `      <Type>${res.type || 1}</Type>\n`;
                xml += '    </Resource>\n';
            }
            xml += '  </Resources>\n';
        }

        // Assignments
        const assignments = project.assignments || [];
        if (assignments.length > 0) {
            xml += '  <Assignments>\n';
            for (let i = 0; i < assignments.length; i++) {
                const a = assignments[i];
                xml += '    <Assignment>\n';
                xml += `      <UID>${i + 1}</UID>\n`;
                xml += `      <TaskUID>${a.taskUID}</TaskUID>\n`;
                xml += `      <ResourceUID>${a.resourceUID}</ResourceUID>\n`;
                xml += `      <Units>${a.units || 1}</Units>\n`;
                xml += '    </Assignment>\n';
            }
            xml += '  </Assignments>\n';
        }

        xml += '</Project>\n';
        return xml;
    }

    // ——— HELPERS ———

    /**
     * Parse ISO duration (e.g., PT24H0M0S) to workdays
     */
    function parseDurationToDays(duration, minutesPerDay) {
        if (!duration) return 1;
        minutesPerDay = minutesPerDay || 480;

        // Handle simple day format: "5 days", "3d" etc.
        const dayMatch = duration.match(/^(\d+)\s*d(ays?)?$/i);
        if (dayMatch) return parseInt(dayMatch[1]);

        // ISO 8601 duration: PT24H0M0S or P5DT0H0M0S
        const match = duration.match(/P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
        if (!match) return 1;

        const days = parseInt(match[1]) || 0;
        const hours = parseInt(match[2]) || 0;
        const minutes = parseInt(match[3]) || 0;

        const totalMinutes = (days * minutesPerDay) + (hours * 60) + minutes;
        return Math.max(Math.round(totalMinutes / minutesPerDay), 0);
    }

    /**
     * Convert days to ISO duration string
     */
    function daysToDuration(days) {
        const hours = (days || 0) * 8;
        return `PT${hours}H0M0S`;
    }

    /**
     * Parse date string to Date object
     */
    function parseDate(dateStr) {
        if (!dateStr) return new Date();
        const d = new Date(dateStr);
        return isNaN(d.getTime()) ? new Date() : d;
    }

    /**
     * Format Date to ISO string
     */
    function formatDateISO(date) {
        if (!date) return new Date().toISOString();
        if (date instanceof Date) return date.toISOString();
        return new Date(date).toISOString();
    }

    /**
     * Escape special XML characters
     */
    function escapeXml(str) {
        if (!str) return '';
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    // Public API
    export const MSProjectXML = { parse, exportXML, parseDurationToDays, daysToDuration, parseDate };
