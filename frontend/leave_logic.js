function isUserAdminOrManager(u) {
  if (!u) return false;
  const r = String(u.Role || '').trim().toLowerCase();
  return (r === 'admin' || r === 'manager' || r.includes('admin') || r.includes('manager') || r.includes('ผู้ดูแล') || r.includes('หัวหน้า'));
}

function getActiveSessionUser() {
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
      else if (args.length === 1 && typeof args[0] === 'object') args = [args[0], token];
    }
    const url = window.GAS_WEB_APP_URL || 'https://script.google.com/macros/s/AKfycbyNSu_HlAemXEdVjaxeNu-m15Uln5qBzv4-ZfnoyoIWKCbCuAfuLN1AnVX9s9zgxuuj/exec';
    fetch(url, {
      method: 'POST',
      body: JSON.stringify({ action: action, args: args, payload: p }),
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }
    })
    .then(res => res.json())
    .then(result => {
      if (result.error) {
        if (result.error.toLowerCase().includes('session expired') || result.error.includes('เข้าสู่ระบบใหม่')) {
          alert('เซสชั่นของคุณหมดอายุ กรุณาเข้าสู่ระบบใหม่อีกครั้ง');
          if (window.shiftFlowLogout) window.shiftFlowLogout();
          else {
            localStorage.removeItem('shiftflow_user');
            sessionStorage.removeItem('shiftflow_user');
            window.location.reload();
          }
        }
        reject(new Error(result.error));
      }
      else if (result.status === 'error') reject(new Error(result.message || 'Error'));
      else resolve(result.data !== undefined ? result.data : result);
    })
    .catch(err => {
      if (err.message && (err.message.toLowerCase().includes('session expired') || err.message.includes('เข้าสู่ระบบใหม่'))) {
          alert('เซสชั่นของคุณหมดอายุ กรุณาเข้าสู่ระบบใหม่อีกครั้ง');
          if (window.shiftFlowLogout) window.shiftFlowLogout();
          else {
            localStorage.removeItem('shiftflow_user');
            sessionStorage.removeItem('shiftflow_user');
            window.location.reload();
          }
      }
      reject(err);
    });
  });
};

window.cleanLeaveDateStr = function(dateStr) {
  if (!dateStr) return '-';
  const str = String(dateStr);
  if (str.includes('GMT') || str.includes('00:00:00') || str.includes('T00:00:00')) {
    return str.split(' ถึง ').map(part => {
      const d = new Date(part);
      if (!isNaN(d.getTime())) {
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
      }
      return part.replace(/00:00:00.*$/g, '').trim();
    }).join(' ถึง ');
  }
  return str;
};

window.loadLeaveView = function() {
  const sessionUser = getActiveSessionUser();
  if (!sessionUser) return;
  const adminContainer = document.getElementById('leave-admin-container');
  if (adminContainer) {
    adminContainer.style.display = (isUserAdminOrManager(sessionUser)) ? 'block' : 'none';
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
        <td style="padding: 14px; font-weight: 500; color: #334155;">${window.cleanLeaveDateStr(h.date)}</td>
        <td style="padding: 14px;"><span style="background: #eff6ff; color: #2563eb; padding: 3px 10px; border-radius: 6px; font-size: 12.5px; font-weight: 600;">${h.type}</span></td>
        <td style="padding: 14px; text-align: center;">${statusBadge}</td>
        <td style="padding: 14px; text-align: center;">${printBtn}</td>
      </tr>`;
    });
    const historyTable = document.getElementById('leave-history-table');
    if (historyTable) {
      historyTable.innerHTML = histHtml || '<tr><td colspan="4" style="text-align:center; padding: 36px; color: #94a3b8;">✨ ยังไม่มีประวัติการยื่นคำขอลา</td></tr>';
    }

    if (isUserAdminOrManager(sessionUser)) {
      const renderPend = (list) => {
        let pendHtml = '';
        (list || []).forEach(p => {
          const fileLink = (p.fileUrl && p.fileUrl !== 'ไม่มีไฟล์แนบ') ? `<a href="${p.fileUrl}" target="_blank" style="display: inline-flex; align-items: center; gap: 4px; background: #eff6ff; color: #2563eb; padding: 4px 10px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 12px;">📎 ดูเอกสาร</a>` : `<span style="color: #94a3b8; font-size: 12px;">-</span>`;
          pendHtml += `<tr style="border-bottom: 1px solid #fef3c7; transition: background 0.15s;">
            <td style="padding: 14px; font-weight: 600; color: #1e293b;">${p.name || p.empName || '-'} <div style="font-size: 11.5px; font-weight: normal; color: #64748b;">${p.department || '-'}</div></td>
            <td style="padding: 14px;"><span style="background: #fffbeb; border: 1px solid #fde68a; color: #d97706; padding: 3px 10px; border-radius: 6px; font-size: 12.5px; font-weight: 600;">${p.type}</span></td>
            <td style="padding: 14px; color: #334155; font-weight: 500;">${window.cleanLeaveDateStr(p.date)}</td>
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
      };

      const renderAllHist = (list) => {
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
            <td style="padding: 14px; font-weight: 600; color: #1e293b;">${item.name || item.empName || '-'}</td>
            <td style="padding: 14px;"><span style="background: #eff6ff; color: #2563eb; padding: 3px 10px; border-radius: 6px; font-size: 12.5px; font-weight: 600;">${item.type}</span></td>
            <td style="padding: 14px; color: #334155; font-weight: 500;">${window.cleanLeaveDateStr(item.date)}</td>
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
      };

      if (data.pendingList && Array.isArray(data.pendingList)) {
        renderPend(data.pendingList);
      } else {
        apiCall('getPendingRequests', []).then(list => renderPend(list));
      }

      if (data.allHistory && Array.isArray(data.allHistory)) {
        renderAllHist(data.allHistory);
      } else {
        apiCall('getAllLeaveHistory', []).then(list => renderAllHist(list));
      }
    }
  }).catch(err => console.error(err));
};

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


window.checkDatabaseInfo = function() {
  const u = getActiveSessionUser();
  if (!u) return alert('กรุณาล็อกอินเข้าสู่ระบบ');
  const btn = document.getElementById('btn-check-db');
  if (btn) btn.innerText = '⏳ กำลังตรวจสอบชีตฐานข้อมูล...';
  apiCall('getDatabaseInfo', [u.token]).then(data => {
    if (btn) btn.innerText = '🗂️ ตรวจสอบชีตฐานข้อมูล และโควตาปัจจุบัน';
    let empListStr = (data.quotaEmployees || []).map(e => `• ${e.name} (ป่วย: ${e.sick}, กิจ: ${e.personal}, พักร้อน: ${e.vacation})`).join('\n');
    let msg = `📑 ชื่อไฟล์ชีตปัจจุบัน: ${data.spreadsheetName}\n` +
              `🔗 ลิงก์ Google Sheets: \n${data.spreadsheetUrl}\n\n` +
              `👥 จำนวนแถวในชีต Quota: ${data.quotaRowsCount} บรรทัด\n` +
              `📝 จำนวนคำขอในชีตลางาน: ${data.leaveRequestsCount} รายการ\n\n` +
              `📌 รายชื่อและโควตาในชีต Quota ขณะนี้:\n${empListStr || '(ยังไม่มีข้อมูล)'}`;
    alert(msg);
    if (confirm('คุณต้องการเปิดหน้า Google Sheets ฐานข้อมูลของระบบนี้ในแท็บใหม่ตอนนี้หรือไม่?')) {
      window.open(data.spreadsheetUrl, '_blank');
    }
  }).catch(err => {
    if (btn) btn.innerText = '🗂️ ตรวจสอบชีตฐานข้อมูล และโควตาปัจจุบัน';
    alert('เกิดข้อผิดพลาดในการตรวจสอบฐานข้อมูล: ' + err.message);
  });
};

window.testLineNotification = function() {
  const u = getActiveSessionUser();
  if (!u) return alert('กรุณาล็อกอินเข้าสู่ระบบ');
  const btn = document.getElementById('btn-test-line');
  if (btn) btn.innerText = '⏳ กำลังยิงข้อความทดสอบ...';
  apiCall('testLineMessage', [u.token]).then(result => {
    if (btn) btn.innerText = '🔔 ทดสอบส่งข้อความแจ้งเตือนเข้ากลุ่ม LINE';
    console.log('LINE Test Result:', result);
    if (result && result.success) {
      alert('✅ ยิงข้อความทดสอบเข้ากลุ่ม LINE สำเร็จเรียบร้อยแล้ว! กรุณาตรวจสอบในแอป LINE ของคุณ');
    } else {
      let details = JSON.stringify(result.results || result, null, 2);
      alert('⚠️ ยิงข้อความเข้า LINE ไม่สำเร็จ! รายละเอียดจากเซิร์ฟเวอร์ LINE:\n\n' + details + '\n\n💡 คำแนะนำ: ตรวจสอบ LINE Access Token และ Group ID ในตั้งค่าระบบ');
    }
  }).catch(err => {
    if (btn) btn.innerText = '🔔 ทดสอบส่งข้อความแจ้งเตือนเข้ากลุ่ม LINE';
    alert('เกิดข้อผิดพลาดในการทดสอบ LINE: ' + err.message);
  });
};
