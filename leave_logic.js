window.loadLeaveView = function() {
  if (!sessionUser) return;
  const adminContainer = document.getElementById('leave-admin-container');
  if (adminContainer) {
    adminContainer.style.display = (sessionUser.Role === 'admin' || sessionUser.Role === 'manager') ? 'block' : 'none';
  }
  
  apiCall('getLeaveSummary', []).then(data => {
    const sickEl = document.getElementById('leave-quota-sick');
    const personalEl = document.getElementById('leave-quota-personal');
    const vacationEl = document.getElementById('leave-quota-vacation');
    if (sickEl) sickEl.innerText = data.remaining.sick;
    if (personalEl) personalEl.innerText = data.remaining.personal;
    if (vacationEl) vacationEl.innerText = data.remaining.vacation;
    
    let histHtml = '';
    data.history.forEach(h => {
      let statusBadge = `<span style="background: #f1f5f9; color: #475569; padding: 4px 10px; border-radius: 20px; font-weight: 600; font-size: 12px;">${h.status}</span>`;
      if (h.status === 'อนุมัติ') {
        statusBadge = `<span style="background: #dcfce7; color: #166534; padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;">✅ อนุมัติ</span>`;
      } else if (h.status === 'รอตรวจสอบ' || h.status.includes('รอ')) {
        statusBadge = `<span style="background: #fef9c3; color: #854d0e; padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;">⏳ รอตรวจสอบ</span>`;
      } else if (h.status === 'ไม่อนุมัติ') {
        statusBadge = `<span style="background: #fee2e2; color: #991b1b; padding: 4px 12px; border-radius: 20px; font-weight: 600; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;">❌ ไม่อนุมัติ</span>`;
      }

      histHtml += `<tr style="transition: background 0.15s; border-bottom: 1px solid #f1f5f9;">
        <td style="padding: 14px; font-weight: 500; color: #334155;">${h.date}</td>
        <td style="padding: 14px;"><span style="background: #eff6ff; color: #2563eb; padding: 3px 10px; border-radius: 6px; font-size: 12.5px; font-weight: 600;">${h.type}</span></td>
        <td style="padding: 14px; text-align: center;">${statusBadge}</td>
      </tr>`;
    });
    const historyTable = document.getElementById('leave-history-table');
    if (historyTable) {
      historyTable.innerHTML = histHtml || '<tr><td colspan="3" style="text-align:center; padding: 36px; color: #94a3b8;">✨ ยังไม่มีประวัติการยื่นคำขอลา</td></tr>';
    }
  }).catch(err => console.error(err));

  if (sessionUser.Role === 'admin' || sessionUser.Role === 'manager') {
    apiCall('getPendingLeaves', []).then(list => {
      let pendHtml = '';
      list.forEach(p => {
        const fileLink = (p.fileUrl && p.fileUrl !== 'ไม่มีไฟล์แนบ') ? `<a href="${p.fileUrl}" target="_blank" style="display: inline-flex; align-items: center; gap: 4px; background: #eff6ff; color: #2563eb; padding: 4px 10px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 12px;">📎 ดูเอกสาร</a>` : `<span style="color: #94a3b8; font-size: 12px;">-</span>`;
        pendHtml += `<tr style="border-bottom: 1px solid #fef3c7; transition: background 0.15s;">
          <td style="padding: 14px; font-weight: 600; color: #1e293b;">${p.name} <div style="font-size: 11.5px; font-weight: normal; color: #64748b;">${p.department}</div></td>
          <td style="padding: 14px;"><span style="background: #fffbeb; border: 1px solid #fde68a; color: #d97706; padding: 3px 10px; border-radius: 6px; font-size: 12.5px; font-weight: 600;">${p.type}</span></td>
          <td style="padding: 14px; color: #334155; font-weight: 500;">${p.date}</td>
          <td style="padding: 14px; color: #1e293b; font-weight: 700;">${p.days}</td>
          <td style="padding: 14px; color: #475569; max-width: 200px;">${p.reason}</td>
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
  }
};

window.submitLeave = function(event) {
  event.preventDefault();
  const btn = document.getElementById('leave-submit-btn');
  btn.innerText = 'กำลังส่ง...';
  btn.disabled = true;

  const type = document.getElementById('leave-type').value;
  const start = document.getElementById('leave-start').value;
  const end = document.getElementById('leave-end').value;
  const reason = document.getElementById('leave-reason').value;
  const fileInput = document.getElementById('leave-file');

  const payload = {
    leaveType: type,
    startDate: start,
    endDate: end,
    reason: reason
  };

  const finalize = () => {
    apiCall('submitLeaveRequest', [payload]).then(res => {
      alert(res.message);
      document.getElementById('leave-form').reset();
      loadLeaveView();
    }).catch(err => {
      alert('เกิดข้อผิดพลาด: ' + err.message);
    }).finally(() => {
      btn.innerText = 'ส่งใบลา';
      btn.disabled = false;
    });
  };

  if (fileInput.files.length > 0) {
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
  if (!confirm('ยืนยันการอนุมัติคำขอนี้? ระบบจะทำการอัปเดตตารางกะเป็น AL อัตโนมัติ')) return;
  apiCall('updateLeaveStatus', [id, 'อนุมัติ']).then(() => {
    alert('อนุมัติสำเร็จ และอัปเดตตารางกะเรียบร้อยแล้ว');
    loadLeaveView();
  }).catch(err => alert(err.message));
};

window.rejectLeave = function(id) {
  if (!confirm('ยืนยันการไม่อนุมัติคำขอนี้?')) return;
  apiCall('updateLeaveStatus', [id, 'ไม่อนุมัติ']).then(() => {
    alert('บันทึกผลสำเร็จ');
    loadLeaveView();
  }).catch(err => alert(err.message));
};

// Hook into the main view rendering
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

