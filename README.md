# 🚀 ProjectFlow™

**Professional Project Management System**

> Built with ❤️ by [Ahmed M. Fawzy](https://www.linkedin.com/in/ahmed-m-fawzy)

---

## 📋 Overview

ProjectFlow™ is a comprehensive, browser-based project management system that combines the power of Microsoft Project with a modern, intuitive interface. Built entirely with vanilla JavaScript and HTML5 Canvas — no heavy frameworks, no subscription fees.

---

## ✨ Key Features

### 📊 Project Management
- **Gantt Chart** — Interactive, drag-and-drop scheduling with dependency links
- **Critical Path Method (CPM)** — Automatic forward/backward pass calculations
- **Resource Management** — Allocation, leveling, and histogram views
- **Work Calendar** — Custom work days, public holidays (20+ countries)
- **Baseline Tracking** — Set and compare against baselines

### 📈 Analytics & Reporting
- **Dashboard** — Real-time KPIs, donut charts, bar charts, S-Curve
- **Earned Value Management (EVM)** — PV, EV, AC, SPI, CPI, EAC, TCPI
- **PDF Reports** — Professional, print-ready project reports with cover page
- **Excel Export** — Multi-sheet XLSX with tasks, resources, and KPIs
- **Portfolio Reports** — Cross-project executive summaries

### 🗂️ Multi-Project Support
- **Portfolio Hub** — Manage unlimited projects from a single interface
- **Project Health Score** — Algorithmic health assessment (0-100)
- **Global Search** — Find tasks across all projects instantly
- **SQLite Backend** — Optional local database for persistent storage

### 🎨 User Experience
- **Dark / Light Mode** — Beautiful UI with smooth transitions
- **RTL Support** — Full Arabic language support
- **Keyboard Shortcuts** — Power-user productivity (⌘K, ⌘1-9, ⌘B)
- **Network Diagram** — PERT chart visualization
- **Calendar View** — Drag-and-drop task scheduling

### 📥 Import / Export
- **Microsoft Project XML** — Full import support
- **MS Planner** — Import from Excel exports
- **Portfolio Backup** — JSON export/import of all projects
- **Gantt PNG** — High-resolution image export

---

## 🛠️ Tech Stack

| Technology | Purpose |
|:-----------|:--------|
| **Vanilla JavaScript** | Core application (zero frameworks) |
| **HTML5 Canvas** | Gantt chart, dashboards, network diagram |
| **jsPDF** | PDF report generation |
| **SheetJS** | Excel file export |
| **Dexie.js** | IndexedDB wrapper for multi-project storage |
| **Node.js + Express** | Optional SQLite server backend |
| **better-sqlite3** | Local database engine |

---

## 🚀 Getting Started

### Option 1: Browser Only (No Server)
Simply open `index.html` in your browser. All data is stored in IndexedDB.

### Option 2: With SQLite Backend
```bash
# Install dependencies (first time only)
npm install

# Start the server
node server.js
# — or double-click "Start Server.command" on macOS

# Open in browser
open http://localhost:3456
```

---

## 📁 Project Structure

```
ProjectFlow/
├── index.html              # Main application
├── index.css               # Styles & design system
├── server.js               # Node.js SQLite backend
├── package.json            # Dependencies
├── LICENSE                 # Proprietary license
├── Start Server.command    # macOS quick start
├── Stop Server.command     # macOS quick stop
├── js/
│   ├── app.js              # Main application logic
│   ├── gantt.js            # Gantt chart renderer
│   ├── critical-path.js    # CPM engine
│   ├── calendar.js         # Work calendar & holidays
│   ├── reports.js          # PDF & Excel reports
│   ├── resource-manager.js # Resource leveling
│   ├── dashboard.js        # Dashboard charts
│   ├── network.js          # Network diagram
│   └── logo-data.js        # Default logo
└── data/                   # SQLite databases (auto-created)
```

---

## 📸 Screenshots

*Coming soon*

---

## 📄 License

**Proprietary Software** — © 2026 Ahmed M. Fawzy. All Rights Reserved.

This software is protected under copyright law. Unauthorized copying, modification, distribution, or use of this software is strictly prohibited. See [LICENSE](LICENSE) for details.

**ProjectFlow™** is a trademark of Ahmed M. Fawzy.

---

## 👨‍💻 Developer

**Ahmed M. Fawzy**  
13+ years in digital transformation, project management & software engineering.

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Ahmed_M._Fawzy-0077B5?style=for-the-badge&logo=linkedin)](https://www.linkedin.com/in/ahmed-m-fawzy)
[![Email](https://img.shields.io/badge/Email-Ahmed.m.fawzy-D14836?style=for-the-badge&logo=gmail)](mailto:Ahmed.m.fawzy@hotmail.com)

---

> *ProjectFlow™ — The smartest project management tool in your browser.*
