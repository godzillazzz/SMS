# AGENTS.md — Security Management System (ระบบจัดการตารางกะและการเข้า-ออกงาน)

## 📌 1. Project Overview & Environment
- **Project Name**: Security Management System (ระบบจัดการตารางกะรักษาความปลอดภัยและระบบอนุมัติการลา)
- **Current Version**: `v88` (`@88` on Google Apps Script)
- **Local Path**: `C:\Users\sermp\OneDrive - PTTPLC\04_SSO\ปี-2569\40.AI\ระบบ Security Management System`
- **Google Apps Script Web App Deployment ID**: `AKfycbz8x2w341mY4z6_S4n0yYnO85ZcQ1iK311ZfU-m-QW4S1X0E-m`
- **Sync Tools**:
  - `clasp`: For pushing code (`Code.gs`, `index.html`, etc.) to Google Apps Script cloud via `clasp push`.
  - `git`: For version control and syncing with GitHub (`git push origin main`).

---

## 🏗️ 2. Core Architecture & File Structure
- **`Code.gs`**: Backend logic running on Google Apps Script servers (`doGet`, `doPost`, `authenticateUser`, `getOptimizedScheduleData`, `exportApprovedSchedule`, `changeRequestStatus`, etc.).
- **`index.html` / `shift_scheduling.html`**: Main Single Page Application (SPA) frontend containing styles, HTML layout (`#app-shell`, `#nav-container`, `.content`), and core JS logic.
- **`leave_view.html` / `leave_logic.js`**: Modular UI and business logic for the Leave Management (`จัดการการลา`) feature.
- **`frontend/` directory**: Mirrored / modular standalone components (`frontend/index.html`, `frontend/gas_index.html`, `frontend/leave_logic.js`) which must be kept synchronized with root files.

---

## 🔐 3. User Roles & Impersonation System (`🐞 View As`)
### Roles (`roleStr`)
1. **Admin (`ผู้ดูแล`)**: Full access to all 8 navigation views and system settings. Can approve/reject any leave request.
2. **Manager (`หัวหน้า` / `Manager`)**: Can access Dashboard, Schedule Calendar, Rule Checking, Reports, Users & Roles, and Leave Management (can approve leave requests). Cannot access Master Data, Shift Setup, or System Settings.
3. **Viewer (`พนักงาน` / `Viewer`)**: Can only view Dashboard, Schedule Calendar, Reports, and submit personal Leave Requests. Cannot see approval controls, Master Data, Shift Setup, Settings, Rule Checking, or Users & Roles.

### Impersonation Mode (`View As`) Rules
- Triggered by Admin in `Users & Roles` menu.
- Stores state in `localStorage` & `sessionStorage`:
  - `shiftflow_is_impersonating` = `'true'`
  - `shiftflow.currentUser` overridden with target user details (`Role: Viewer` or `Role: Manager`).
- **Visual Warning Banner**: Displays sticky orange banner (`#f59e0b` to `#d97706`) at top with button `⬅️ กลับสู่บัญชี Admin ของคุณ`.
- **Strict Navigation Parity (`applyNavPermissions`)**: Navigation clicks (`item.addEventListener('click')`) MUST work smoothly in Impersonation mode while respecting exact visibility rules of the impersonated role (`item.hidden` and `item.style.display`).

---

## 🖨️ 4. PDF Export & Report Formatting (`Schedule Calendar & Leave`)
- **Format**: A4 Landscape (`A4 แนวนอน`).
- **Signature Box Positioning**:
  - Exactly **65% width of the right side** of the page (`width: 65% !important; margin-left: auto;`).
  - Spacing / Line height equal to `margin-bottom: 28px` between lines for comfortable real-ink signing.
- **Signature Blocks**:
  1. `ลงชื่อ...............................................................................................` และ `(...............................................................................................)` -> **พนักงานผู้จัดพิมพ์รายงาน / หัวหน้าพนักงานรักษาความปลอดภัย**
  2. `ทราบ / ลงชื่อ...............................................................................................` และ `(...............................................................................................)` -> **ผู้จัดการเขต (ผู้อนุมัติ)**
- **Implementation**: Handled both via `sheet.getRange(...).merge()` & row height in `Code.gs` (`exportApprovedSchedule`) for GAS exports, and via CSS `.signature-box` (`margin-bottom: 28px`) for HTML Print.

---

## 🛠️ 5. Standard Development Rules for AI Agents
1. **Never Break Navigation**: When modifying navigation permissions or styles in `index.html` / `shift_scheduling.html`, ensure click event listeners (`item.addEventListener('click', ...)`) inside the IIFE remain attached and functional.
2. **Sync Across All Frontend Files**: Any bugfix or UI modification in `index.html` MUST be mirrored identically across `shift_scheduling.html`, `frontend/index.html`, and `frontend/gas_index.html`.
3. **Always Push via Clasp**: After modifying code, always run `clasp push` via terminal before instructing the user to test (`Ctrl + Shift + R`).
4. **Preserve Comments**: Preserve all existing documentation and comments inside `Code.gs` and HTML script tags unless requested otherwise.
