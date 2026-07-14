window.loadLeaveView = function() {
  if (!sessionUser) return;
  document.getElementById('leave-admin-container').style.display = (sessionUser.Role === 'admin' || sessionUser.Role === 'manager') ? 'block' : 'none';
  
  apiCall('getLeaveSummary', []).then(data => {
    document.getElementById('leave-quota-sick').innerText = data.remaining.sick;
    document.getElementById('leave-quota-personal').innerText = data.remaining.personal;
    document.getElementById('leave-quota-vacation').innerText = data.remaining.vacation;
    
    let histHtml = '';
    data.history.forEach(h => {
      histHtml += `<tr>
        <td>${h.date}</td>
        <td>${h.type}</td>
        <td>${h.status}</td>
      </tr>`;
    });
    document.getElementById('leave-history-table').innerHTML = histHtml || '<tr><td colspan="3" style="text-align:center;">ไม่มีประวัติการลา</td></tr>';
  }).catch(err => console.error(err));

  if (sessionUser.Role === 'admin' || sessionUser.Role === 'manager') {
    apiCall('getPendingLeaves', []).then(list => {
      let pendHtml = '';
      list.forEach(p => {
        const fileLink = (p.fileUrl && p.fileUrl !== 'ไม่มีไฟล์แนบ') ? `<a href="${p.fileUrl}" target="_blank" style="color:#3B82F6;text-decoration:underline;">ดูไฟล์</a>` : '-';
        pendHtml += `<tr>
          <td>${p.name} (${p.department})</td>
          <td>${p.type}</td>
          <td>${p.date}</td>
          <td>${p.days}</td>
          <td>${p.reason}</td>
          <td>${fileLink}</td>
          <td>
            <button onclick="approveLeave(${p.id})" style="background:#10B981;color:white;padding:5px 10px;border-radius:5px;border:none;cursor:pointer;margin-right:5px;font-family:inherit;">อนุมัติ</button>
            <button onclick="rejectLeave(${p.id})" style="background:#EF4444;color:white;padding:5px 10px;border-radius:5px;border:none;cursor:pointer;font-family:inherit;">ไม่อนุมัติ</button>
          </td>
        </tr>`;
      });
      document.getElementById('leave-pending-table').innerHTML = pendHtml || '<tr><td colspan="7" style="text-align:center;">ไม่มีคำขอที่รออนุมัติ</td></tr>';
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

