function getActiveSessionUser() {
  const isImp = localStorage.getItem('shiftflow_is_impersonating') === 'true' || sessionStorage.getItem('shiftflow_is_impersonating') === 'true';
  if (isImp) {
    try {
      const mockUser = JSON.parse(sessionStorage.getItem('shiftflow.currentUser') || localStorage.getItem('shiftflow.currentUser') || sessionStorage.getItem('shiftflow_session_v1') || localStorage.getItem('shiftflow_session_v1') || 'null');
      if (mockUser && (mockUser.isImpersonated || mockUser.Role || mockUser.role)) {
        window.sessionUser = mockUser;
        window.shiftFlowCurrentUser = mockUser;
        return mockUser;
      }
    } catch(e) {}
  }
  const u = window.sessionUser || window.shiftFlowCurrentUser || (typeof currentSessionUser === 'function' ? currentSessionUser() : null);
  if (u) window.sessionUser = u;
  return u;
}
window.localLeaveHistoryList = [];
window.currentLeaveRemaining = { sick: 0, personal: 0, vacation: 0 };

window.apiCall = function(action, args = [], payloadObj = null) {
  return new Promise((resolve, reject) => {
    const p = payloadObj || (Array.isArray(args) && args.length > 0 && typeof args[0] === 'object' ? args[0] : {});
    const token = window.sessionUser ? window.sessionUser.token : (window.shiftFlowCurrentUser ? window.shiftFlowCurrentUser.token : '');
    if (token) {
      p.token = token;
      if (!args || args.length === 0) args = [token];
      else if (typeof args[0] !== 'string') args = [token, ...args];
    }
    const url = window.GAS_WEB_APP_URL || 'https://script.google.com/macros/s/AKfycbyNSu_HlAemXEdVjaxeNu-m15Uln5qBzv4-ZfnoyoIWKCbCuAfuLN1AnVX9s9zgxuuj/exec';
    fetch(url, {
      method: 'POST',
      body: JSON.stringify({ action: action, args: args, payload: p }),
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }
    })
    .then(res => res.json())
    .then(result => {
      if (result.error) reject(new Error(result.error));
      else if (result.status === 'error') reject(new Error(result.message || 'Error'));
      else resolve(result.data !== undefined ? result.data : result);
    })
    .catch(err => reject(err));
  });
};

window.loadLeaveView = function() {
  const sessionUser = getActiveSessionUser();
  if (!sessionUser) return;
  const adminContainer = document.getElementById('leave-admin-container');
  if (adminContainer) {
    const r = String(sessionUser.Role || sessionUser.role || '').trim().toLowerCase();
    const isAdmOrMgr = (r === 'admin' || r === 'manager' || r.includes('admin') || r.includes('manager') || r.includes('ผู้ดูแล') || r.includes('หัวหน้า'));
    adminContainer.style.display = isAdmOrMgr ? 'block' : 'none';
  }
  
  const empName = sessionUser.FullName || sessionUser.Name || sessionUser.Username || '';
  apiCall('getLeaveSummary', [], { name: empName, month: 'all', year: '' }).then(data => {
    window.currentLeaveRemaining = data.remaining || { sick: 0, personal: 0, vacation: 0 };
    const sickEl = document.getElementById('leave-quota-sick');
    const personalEl = document.getElementById('leave-quota-personal');
    const vacationEl = document.getElementById('leave-quota-vacation');
    if (sickEl) sickEl.innerText = data.remaining.sick;
    if (personalEl) personalEl.innerText = data.remaining.personal;
    if (vacationEl) vacationEl.innerText = data.remaining.vacation;
    
    window.localLeaveHistoryList = data.details || data.history || [];
    let histHtml = '';
    window.localLeaveHistoryList.forEach((h, idx) => {
      let statusBadge = `<span style="background: #f1f5f9; color: #475569; padding: 4px 10px; border-radius: 20px; font-weight: 600; font-size: 12px;">${h.status}</span>`;
      let printBtn = `<button disabled style="background: #e2e8f0; color: #94a3b8; padding: 5px 12px; border-radius: 8px; border: none; font-size: 12px; cursor: not-allowed; display: inline-flex; align-items: center; gap: 4px;">🖨️ พิมพ์ A4</button>`;
      
      if (h.status === 'อนุมัติ') {
        statusBadge = `<span style="background: #dcfce7; color: #166534; padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;">✅ อนุมัติ</span>`;
        printBtn = `<button onclick="window.printLeaveA4(${idx})" style="background: linear-gradient(135deg, #059669 0%, #10b981 100%); color: white; padding: 5px 14px; border-radius: 8px; border: none; font-weight: 600; font-size: 12px; cursor: pointer; box-shadow: 0 2px 4px rgba(16, 185, 129, 0.25); display: inline-flex; align-items: center; gap: 4px; transition: transform 0.15s;">🖨️ พิมพ์ A4</button>`;
      } else if (h.status === 'รอตรวจสอบ' || h.status === 'รออนุมัติ' || h.status.includes('รอ')) {
        statusBadge = `<span style="background: #fef9c3; color: #854d0e; padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;">⏳ รออนุมัติ</span>`;
      } else if (h.status === 'ไม่อนุมัติ') {
        statusBadge = `<span style="background: #fee2e2; color: #991b1b; padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;">❌ ไม่อนุมัติ</span>`;
      }

      histHtml += `<tr style="transition: background 0.15s; border-bottom: 1px solid #f1f5f9;">
        <td style="padding: 14px; font-weight: 500; color: #334155;">${h.date}</td>
        <td style="padding: 14px;"><span style="background: #eff6ff; color: #2563eb; padding: 3px 10px; border-radius: 6px; font-size: 12.5px; font-weight: 600;">${h.type}</span></td>
        <td style="padding: 14px; text-align: center;">${statusBadge}</td>
        <td style="padding: 14px; text-align: center;">${printBtn}</td>
      </tr>`;
    });
    const historyTable = document.getElementById('leave-history-table');
    if (historyTable) {
      historyTable.innerHTML = histHtml || '<tr><td colspan="4" style="text-align:center; padding: 36px; color: #94a3b8;">✨ ยังไม่มีประวัติการยื่นคำขอลา</td></tr>';
    }
  }).catch(err => console.error(err));

  if (sessionUser.Role === 'admin' || sessionUser.Role === 'manager') {
    apiCall('getPendingRequests', []).then(list => {
      let pendHtml = '';
      (list || []).forEach(p => {
        const fileLink = (p.fileUrl && p.fileUrl !== 'ไม่มีไฟล์แนบ') ? `<a href="${p.fileUrl}" target="_blank" style="display: inline-flex; align-items: center; gap: 4px; background: #eff6ff; color: #2563eb; padding: 4px 10px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 12px;">📎 ดูเอกสาร</a>` : `<span style="color: #94a3b8; font-size: 12px;">-</span>`;
        pendHtml += `<tr style="border-bottom: 1px solid #fef3c7; transition: background 0.15s;">
          <td style="padding: 14px; font-weight: 600; color: #1e293b;">${p.name} <div style="font-size: 11.5px; font-weight: normal; color: #64748b;">${p.department || '-'}</div></td>
          <td style="padding: 14px;"><span style="background: #fffbeb; border: 1px solid #fde68a; color: #d97706; padding: 3px 10px; border-radius: 6px; font-size: 12.5px; font-weight: 600;">${p.type}</span></td>
          <td style="padding: 14px; color: #334155; font-weight: 500;">${p.date}</td>
          <td style="padding: 14px; color: #1e293b; font-weight: 700;">${p.days}</td>
          <td style="padding: 14px; color: #475569; max-width: 200px;">${p.reason || '-'}</td>
          <td style="padding: 14px;">${fileLink}</td>
          <td style="padding: 14px; text-align: center; white-space: nowrap;">
            <button onclick="approveLeave(${p.id})" style="background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; padding: 6px 14px; border-radius: 8px; border: none; cursor: pointer; font-weight: 600; font-size: 12.5px; box-shadow: 0 2px 4px rgba(16, 185, 129, 0.2); margin-right: 6px;">✅ อนุมัติ</button>
            <button onclick="rejectLeave(${p.id})" style="background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: white; padding: 6px 14px; border-radius: 8px; border: none; cursor: pointer; font-weight: 600; font-size: 12.5px; box-shadow: 0 2px 4px rgba(239, 68, 68, 0.2);">❌ ไม่อนุมัติ</button>
          </td>
        </tr>`;
      });
      const pendingTable = document.getElementById('leave-pending-table');
      if (pendingTable) {
        pendingTable.innerHTML = pendHtml || '<tr><td colspan="7" style="text-align:center; padding: 36px; color: #b45309;">🎉 ยอดเยี่ยม! ไม่มีคำขอที่รออนุมัติค้างอยู่</td></tr>';
      }
    }).catch(err => console.error(err));

    apiCall('getAllLeaveHistory', []).then(list => {
      window.allEmployeesLeaveHistoryList = list || [];
      let allHtml = '';
      (list || []).forEach((item, idx) => {
        let statusBadge = `<span style="background: #f1f5f9; color: #475569; padding: 4px 10px; border-radius: 20px; font-weight: 600; font-size: 12px;">${item.status}</span>`;
        let printBtn = `<button disabled style="background: #e2e8f0; color: #94a3b8; padding: 5px 12px; border-radius: 8px; border: none; font-size: 12px; cursor: not-allowed; display: inline-flex; align-items: center; gap: 4px;">🖨️ พิมพ์ A4</button>`;
        
        if (item.status === 'อนุมัติ') {
          statusBadge = `<span style="background: #dcfce7; color: #166534; padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;">✅ อนุมัติ</span>`;
          printBtn = `<button onclick="window.printLeaveA4Admin(${idx})" style="background: linear-gradient(135deg, #059669 0%, #10b981 100%); color: white; padding: 5px 14px; border-radius: 8px; border: none; font-weight: 600; font-size: 12px; cursor: pointer; box-shadow: 0 2px 4px rgba(16, 185, 129, 0.25); display: inline-flex; align-items: center; gap: 4px; transition: transform 0.15s;">🖨️ พิมพ์ A4</button>`;
        } else if (item.status === 'รอตรวจสอบ' || item.status === 'รออนุมัติ' || item.status.includes('รอ')) {
          statusBadge = `<span style="background: #fef9c3; color: #854d0e; padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;">⏳ รออนุมัติ</span>`;
        } else if (item.status === 'ไม่อนุมัติ') {
          statusBadge = `<span style="background: #fee2e2; color: #991b1b; padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;">❌ ไม่อนุมัติ</span>`;
        }

        allHtml += `<tr style="transition: background 0.15s; border-bottom: 1px solid #f1f5f9;">
          <td style="padding: 14px; font-weight: 600; color: #1e293b;">${item.name}</td>
          <td style="padding: 14px;"><span style="background: #eff6ff; color: #2563eb; padding: 3px 10px; border-radius: 6px; font-size: 12.5px; font-weight: 600;">${item.type}</span></td>
          <td style="padding: 14px; color: #334155; font-weight: 500;">${item.date}</td>
          <td style="padding: 14px; color: #1e293b; font-weight: 700;">${item.days}</td>
          <td style="padding: 14px; color: #475569; max-width: 200px;">${item.reason || '-'}</td>
          <td style="padding: 14px; text-align: center;">${statusBadge}</td>
          <td style="padding: 14px; text-align: center;">${printBtn}</td>
        </tr>`;
      });
      const allTable = document.getElementById('leave-all-history-table');
      if (allTable) {
        allTable.innerHTML = allHtml || '<tr><td colspan="7" style="text-align:center; padding: 36px; color: #94a3b8;">✨ ยังไม่มีประวัติการลาในระบบ</td></tr>';
      }
    }).catch(err => console.error(err));
  }
};


/* ==========================================================
   🛡️ ระบบตรวจสอบโควตาวันลาคงเหลือก่อนยื่น (Client-Side Validation)
   ========================================================== */
window.checkLeaveQuotaOnForm = function() {
  const sessionUser = getActiveSessionUser();
  const type = document.getElementById('leave-type')?.value;
  const start = document.getElementById('leave-start')?.value;
  const end = document.getElementById('leave-end')?.value;
  const warningEl = document.getElementById('leave-quota-warning');
  const submitBtn = document.getElementById('leave-submit-btn');
  const submitText = document.getElementById('leave-submit-text');

  if (!type || !start || !end || !warningEl || !submitBtn) return;

  const startDate = new Date(start);
  const endDate = new Date(end);
  const diffTime = endDate - startDate;
  if (diffTime < 0) {
    warningEl.style.display = 'block';
    warningEl.innerText = '⚠️ วันที่สิ้นสุดต้องอยู่หลังหรือวันเดียวกับวันที่เริ่มต้น';
    submitBtn.disabled = true;
    submitBtn.style.opacity = '0.5';
    submitBtn.style.cursor = 'not-allowed';
    return;
  }

  const days = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
  let remaining = 0;
  if (type === 'ลาป่วย') remaining = window.currentLeaveRemaining.sick || 0;
  else if (type === 'ลากิจ') remaining = window.currentLeaveRemaining.personal || 0;
  else if (type === 'ลาพักร้อน') remaining = window.currentLeaveRemaining.vacation || 0;

  if (remaining < days) {
    warningEl.style.display = 'block';
    warningEl.innerText = `⚠️ คุณเลือกคำขอลาจำนวน ${days} วัน แต่โควตา ${type} คงเหลือมีเพียง ${remaining} วัน เท่านั้น`;
    submitBtn.disabled = true;
    submitBtn.style.opacity = '0.5';
    submitBtn.style.cursor = 'not-allowed';
    if (submitText) submitText.innerText = '🚫 วันลาคงเหลือไม่เพียงพอ';
  } else {
    warningEl.style.display = 'none';
    submitBtn.disabled = false;
    submitBtn.style.opacity = '1';
    submitBtn.style.cursor = 'pointer';
    if (submitText) submitText.innerText = '🚀 ยืนยันและส่งคำขอลา';
  }
};

window.printLeaveA4 = function(index) {
  const sessionUser = getActiveSessionUser();
  const item = window.localLeaveHistoryList[index];
  if (!item) return;

  const today = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });
  const empNameEl = document.getElementById('printEmpName');
  const dateEl = document.getElementById('printDate');
  const titleEl = document.getElementById('printReportTitle');
  const rowBody = document.getElementById('printSingleRowBody');
  const printContainer = document.getElementById('globalPrintReport') || document.getElementById('leave-a4-print-area');

  if (empNameEl) empNameEl.innerText = sessionUser?.FullName || sessionUser?.Name || 'พนักงาน';
  if (dateEl) dateEl.innerText = today;
  if (titleEl) titleEl.innerText = "ใบขออนุมัติลางาน";

  const cleanDate = typeof window.cleanLeaveDateStr === 'function' ? window.cleanLeaveDateStr(item.date) : item.date;
  const detailsOrReason = item.reason || item.details || item.type || '-';
  const daysText = item.days ? `${item.days} วัน` : '1 วัน';

  if (rowBody) {
    rowBody.innerHTML = `
      <tr>
        <td style="padding: 12px; border: 1px solid #000; text-align: center; font-weight: 500;">${cleanDate}</td>
        <td style="padding: 12px; border: 1px solid #000; text-align: center; font-weight: bold; color: #1d4ed8;">${item.type}</td>
        <td style="padding: 12px; border: 1px solid #000; text-align: center;">${daysText}</td>
        <td style="padding: 12px; border: 1px solid #000; text-align: left;">${detailsOrReason}</td>
      </tr>
    `;
  }

  window.print();
};

window.printLeaveA4Admin = function(index) {
  const sessionUser = getActiveSessionUser();
  const item = (window.allEmployeesLeaveHistoryList || [])[index];
  if (!item) return;

  const today = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });
  const empNameEl = document.getElementById('printEmpName');
  const dateEl = document.getElementById('printDate');
  const titleEl = document.getElementById('printReportTitle');
  const rowBody = document.getElementById('printSingleRowBody');
  const printContainer = document.getElementById('globalPrintReport') || document.getElementById('leave-a4-print-area');

  if (empNameEl) empNameEl.innerText = item.name || item.empName || 'พนักงาน';
  if (dateEl) dateEl.innerText = today;
  if (titleEl) titleEl.innerText = "ใบขออนุมัติลางาน";

  const cleanDate = typeof window.cleanLeaveDateStr === 'function' ? window.cleanLeaveDateStr(item.date) : item.date;
  const detailsOrReason = item.reason || item.details || item.type || '-';
  const daysText = item.days ? `${item.days} วัน` : '1 วัน';

  if (rowBody) {
    rowBody.innerHTML = `
      <tr>
        <td style="padding: 12px; border: 1px solid #000; text-align: center; font-weight: 500;">${cleanDate}</td>
        <td style="padding: 12px; border: 1px solid #000; text-align: center; font-weight: bold; color: #1d4ed8;">${item.type}</td>
        <td style="padding: 12px; border: 1px solid #000; text-align: center;">${daysText}</td>
        <td style="padding: 12px; border: 1px solid #000; text-align: left;">${detailsOrReason}</td>
      </tr>
    `;
  }

  window.print();
};

window.submitLeave = function(event) {
  const sessionUser = getActiveSessionUser();
  if (event && event.preventDefault) event.preventDefault();
  const btn = document.getElementById('leave-submit-btn');
  const text = document.getElementById('leave-submit-text');
  if (text) text.innerText = '⏳ กำลังส่งข้อมูล...';
  if (btn) btn.disabled = true;

  const type = document.getElementById('leave-type').value;
  const start = document.getElementById('leave-start').value;
  const end = document.getElementById('leave-end').value;
  const sub = document.getElementById('leave-substitute')?.value || '';
  const reasonText = document.getElementById('leave-reason')?.value?.trim() || '';
  const fileInput = document.getElementById('leave-file');

  const fullReason = sub ? (reasonText ? `[แทน: ${sub}] ${reasonText}` : `[แทน: ${sub}]`) : (reasonText || '-');

  const payload = {
    name: sessionUser?.FullName || sessionUser?.Name || 'พนักงาน',
    department: sessionUser?.Department || sessionUser?.Workspace || 'ศูนย์ปฏิบัติการระบบท่อเขต 11',
    leaveType: type,
    startDate: start,
    endDate: end,
    reason: fullReason
  };

  const finalize = () => {
    apiCall('saveLeaveRequest', [], payload).then(res => {
      alert(res.message || 'บันทึกสำเร็จ');
      const form = document.getElementById('leave-form');
      if (form) form.reset();
      loadLeaveView();
    }).catch(err => {
      alert('เกิดข้อผิดพลาด: ' + err.message);
    }).finally(() => {
      if (text) text.innerText = '🚀 ยืนยันและส่งคำขอลา';
      if (btn) btn.disabled = false;
    });
  };

  if (fileInput && fileInput.files && fileInput.files.length > 0) {
    const file = fileInput.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
      payload.fileBase64 = e.target.result.split(',')[1];
      payload.fileName = file.name;
      payload.mimeType = file.type;
      finalize();
    };
    reader.readAsDataURL(file);
  } else {
    finalize();
  }
};

window.approveLeave = function(id) {
  if (!confirm('ยืนยันการอนุมัติคำขอนี้? ระบบจะทำการอัปเดตตารางและหักโควตาอัตโนมัติ')) return;
  apiCall('changeRequestStatus', [], { requestId: id, status: 'อนุมัติ' }).then(() => {
    alert('อนุมัติสำเร็จ');
    loadLeaveView();
  }).catch(err => alert(err.message));
};

window.rejectLeave = function(id) {
  if (!confirm('ยืนยันการไม่อนุมัติคำขอนี้?')) return;
  apiCall('changeRequestStatus', [], { requestId: id, status: 'ไม่อนุมัติ' }).then(() => {
    alert('บันทึกผลสำเร็จ');
    loadLeaveView();
  }).catch(err => alert(err.message));
};

const originalRenderView = window.renderView;
window.renderView = function(viewId) {
  if (originalRenderView) originalRenderView(viewId);
  const leaveView = document.getElementById('view-leave');
  if (viewId === 'dashboard' || viewId !== 'leave') {
    if(leaveView) leaveView.style.display = 'none';
  }
  if (viewId === 'leave') {
    if(leaveView) leaveView.style.display = 'block';
    loadLeaveView();
  }
};



// ==========================================
// 🐞 IMPERSONATION MODE (View As User Sandbox)
// ==========================================

// ==========================================
// 🐞 IMPERSONATION MODE (View As User Sandbox) - ROBUST V84
// ==========================================

// ==========================================
// 🐞 IMPERSONATION MODE (View As User Sandbox) - ROBUST V85
// ==========================================
window.impersonateUser = function(userId) {
  if (!confirm('🐞 ยืนยันเข้าตรวจสอบระบบในมุมมองของ User ID: ' + userId + ' ใช่หรือไม่?')) return;
  
  let currentSession = window.shiftFlowCurrentUser || window.sessionUser || (typeof getActiveSessionUser === 'function' ? getActiveSessionUser() : null);
  if (!currentSession) {
    try {
      currentSession = JSON.parse(sessionStorage.getItem('shiftflow.currentUser') || localStorage.getItem('shiftflow.currentUser') || sessionStorage.getItem('shiftflow_session_v1') || localStorage.getItem('shiftflow_session_v1') || 'null');
    } catch(e) {}
  }
  if (!currentSession || (!currentSession.Email && !currentSession.email && !currentSession.EmployeeID && !currentSession.id && !currentSession['User ID'])) {
    if (window.currentData && typeof window.currentData === 'function') {
      const d = window.currentData();
      if (d && d.session && (d.session.User || d.session.Role)) {
        currentSession = {
          Email: d.session.User,
          email: d.session.User,
          Role: d.session.Role || 'Admin',
          FullName: d.session.User || 'Admin'
        };
      }
    }
  }
  
  if (!currentSession) {
    currentSession = { Email: 'admin@ptt.com', email: 'admin@ptt.com', Role: 'Admin', FullName: 'Admin' };
  }

  // Save original session first
  if (!localStorage.getItem('shiftflow.originalAdmin') && !sessionStorage.getItem('shiftflow.originalAdmin')) {
    const orig = JSON.stringify(currentSession);
    localStorage.setItem('shiftflow.originalAdmin', orig);
    sessionStorage.setItem('shiftflow.originalAdmin', orig);
  }

  let targetUser = null;
  if (window.currentData && typeof window.currentData === 'function') {
    const data = window.currentData();
    if (data && data.users) {
      targetUser = data.users.find(u => String(u['User ID']) === String(userId) || String(u.Email || '').toLowerCase() === String(userId).toLowerCase() || String(u.EmployeeID) === String(userId) || String(u.Username) === String(userId) || String(u.Name) === String(userId));
    }
    if (!targetUser && data && data.employees) {
      targetUser = data.employees.find(u => String(u.id) === String(userId) || String(u.Email || '').toLowerCase() === String(userId).toLowerCase() || String(u.EmployeeID) === String(userId) || String(u.name) === String(userId) || String(u.Name) === String(userId));
    }
  }
  
  if (!targetUser) {
    targetUser = {
      'User ID': userId,
      EmployeeID: userId,
      Name: 'พนักงานรหัส ' + userId,
      FullName: 'พนักงานรหัส ' + userId,
      Role: 'Viewer',
      Department: 'ส่วนปฏิบัติการระบบท่อเขต 11',
      Zone: 'เขต 11',
      Email: userId + '@ptt.com'
    };
  }
  
  const targetId = targetUser['User ID'] || targetUser.EmployeeID || targetUser.id || userId;
  const targetName = targetUser.Name || targetUser.FullName || targetUser.name || ('พนักงานรหัส ' + userId);
  const targetEmail = targetUser.Email || targetUser.email || (userId + '@ptt.com');
  const targetRole = targetUser.Role || targetUser.role || 'Viewer';
  const targetDept = targetUser.Department || targetUser.department || 'ส่วนปฏิบัติการระบบท่อเขต 11';

  const mockSession = {
    id: targetId,
    EmployeeID: targetId,
    'User ID': targetId,
    name: targetName,
    Name: targetName,
    FullName: targetName,
    email: targetEmail,
    Email: targetEmail,
    role: targetRole,
    Role: targetRole,
    department: targetDept,
    Department: targetDept,
    Zone: targetUser.Zone || 'เขต 11',
    token: currentSession.token || 'impersonated-token',
    Status: 'Active',
    isImpersonated: true,
    originalAdminEmail: currentSession.Email || currentSession.email || currentSession.User || 'Admin'
  };
  
  const mockStr = JSON.stringify(mockSession);
  sessionStorage.setItem('shiftflow.currentUser', mockStr);
  localStorage.setItem('shiftflow.currentUser', mockStr);
  sessionStorage.setItem('shiftflow_session_v1', mockStr);
  localStorage.setItem('shiftflow_session_v1', mockStr);
  localStorage.setItem('shiftflow_is_impersonating', 'true');
  sessionStorage.setItem('shiftflow_is_impersonating', 'true');

  window.shiftFlowCurrentUser = mockSession;
  window.sessionUser = mockSession;
  
  alert('🐞 โหมด View As User: เข้าสู่ระบบในมุมมองของ ' + targetName + ' (Role: ' + targetRole + ') เรียบร้อยแล้ว!');
  location.reload();
};

window.exitImpersonation = function() {
  if (!confirm('⬅️ ยืนยันกลับสู่บัญชี Admin ของคุณใช่หรือไม่?')) return;
  const originalAdminStr = localStorage.getItem('shiftflow.originalAdmin') || sessionStorage.getItem('shiftflow.originalAdmin') || localStorage.getItem('shiftflow_original_admin_v1') || sessionStorage.getItem('shiftflow_original_admin_v1');
  if (originalAdminStr) {
    sessionStorage.setItem('shiftflow.currentUser', originalAdminStr);
    localStorage.setItem('shiftflow.currentUser', originalAdminStr);
    sessionStorage.setItem('shiftflow_session_v1', originalAdminStr);
    localStorage.setItem('shiftflow_session_v1', originalAdminStr);
    
    localStorage.removeItem('shiftflow.originalAdmin');
    sessionStorage.removeItem('shiftflow.originalAdmin');
    localStorage.removeItem('shiftflow_original_admin_v1');
    sessionStorage.removeItem('shiftflow_original_admin_v1');
    localStorage.removeItem('shiftflow_is_impersonating');
    sessionStorage.removeItem('shiftflow_is_impersonating');
    
    alert('✅ กลับสู่บัญชี Admin ของคุณเรียบร้อยแล้ว');
    location.reload();
  } else {
    sessionStorage.removeItem('shiftflow.currentUser');
    localStorage.removeItem('shiftflow.currentUser');
    localStorage.removeItem('shiftflow_is_impersonating');
    sessionStorage.removeItem('shiftflow_is_impersonating');
    location.reload();
  }
};

window.checkImpersonationStatus = function() {
  const isImpersonating = localStorage.getItem('shiftflow_is_impersonating') === 'true' || sessionStorage.getItem('shiftflow_is_impersonating') === 'true';
  if (isImpersonating) {
    const currentSession = window.shiftFlowCurrentUser || window.sessionUser || JSON.parse(sessionStorage.getItem('shiftflow.currentUser') || localStorage.getItem('shiftflow.currentUser') || sessionStorage.getItem('shiftflow_session_v1') || '{}');
    let banner = document.getElementById('impersonation-banner');
    if (!banner) {
      banner = document.createElement('div');
      banner.id = 'impersonation-banner';
      banner.style.cssText = 'background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; padding: 10px 20px; font-weight: 600; font-size: 14px; text-align: center; position: sticky; top: 0; z-index: 999999; display: flex; justify-content: center; align-items: center; gap: 16px; box-shadow: 0 4px 12px rgba(217, 119, 6, 0.4); flex-wrap: wrap;';
      banner.innerHTML = `
        <span>🚨 <b>โหมด Debug (View As):</b> คุณกำลังใช้งานในมุมมองของ <u>${currentSession.Name || currentSession.name || 'พนักงาน'}</u> (Role: ${currentSession.Role || currentSession.role || 'Viewer'})</span>
        <button onclick="window.exitImpersonation()" style="background: white; color: #d97706; padding: 5px 14px; border-radius: 6px; border: none; font-weight: bold; cursor: pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.15); transition: transform 0.1s;">⬅️ กลับสู่บัญชี Admin ของคุณ</button>
      `;
      if (document.body) document.body.prepend(banner);
      else window.addEventListener('DOMContentLoaded', () => document.body.prepend(banner));
    }
    
    const r = String(currentSession.Role || currentSession.role || '').trim().toLowerCase();
    const isAdminRole = r === 'admin' || r.includes('admin') || r.includes('ผู้ดูแล');
    const isManagerRole = r === 'manager' || r.includes('manager') || r.includes('หัวหน้า');
    
    const adminContainer = document.getElementById('leave-admin-container');
    if (adminContainer) {
      adminContainer.style.display = (isAdminRole || isManagerRole) ? 'block' : 'none';
    }

    document.querySelectorAll('.nav-item').forEach((item) => {
      const txt = item.textContent.trim().toLowerCase();
      if (txt.includes('master data') || txt.includes('shift setup') || txt.includes('settings')) {
        item.hidden = !isAdminRole;
        item.style.display = isAdminRole ? '' : 'none';
      } else if (txt.includes('rule checking') || txt.includes('users & roles') || item.hasAttribute('data-user-approver')) {
        item.hidden = (!isAdminRole && !isManagerRole);
        item.style.display = (isAdminRole || isManagerRole) ? '' : 'none';
      } else {
        item.hidden = false;
        item.style.display = '';
      }
    });

    const activeNav = document.querySelector('.nav-item.active');
    if (activeNav && (activeNav.hidden || activeNav.style.display === 'none')) {
      const dashNav = Array.from(document.querySelectorAll('.nav-item')).find(n => n.textContent.trim().toLowerCase().includes('dashboard'));
      if (dashNav && typeof renderView === 'function') {
        document.querySelectorAll('.nav-item').forEach((nav) => nav.classList.remove('active'));
        dashNav.classList.add('active');
        renderView('dashboard');
      }
    }
  }
};
if (typeof window !== 'undefined') {
  setTimeout(() => window.checkImpersonationStatus(), 300);
}


