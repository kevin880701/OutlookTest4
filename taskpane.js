/* global Office, document */

Office.onReady(() => {
    loadItemData();
    document.getElementById("btnVerify").onclick = markAsVerified;
});

function loadItemData() {
    const item = Office.context.mailbox.item;

    Promise.all([
        new Promise(r => item.from.getAsync(x => r(x.value))),
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.bcc.getAsync(x => r(x.value || []))),
        new Promise(r => item.getAttachmentsAsync(x => r(x.value || [])))
    ]).then(([from, to, cc, bcc, attachments]) => {
        
        renderSingleItem("from-list", from);
        renderList("to-list", to);
        renderList("cc-list", cc);
        renderList("bcc-list", bcc);
        renderAttachments("attachments-list", attachments);

        checkAllChecked();

    }).catch(err => {
        console.error(err);
        document.body.innerHTML = "<h3 style='color:red'>讀取錯誤</h3>" + err.message;
    });
}

function renderSingleItem(containerId, data) {
    const container = document.getElementById(containerId);
    container.innerHTML = "";

    if (!data) {
        container.innerHTML = "<div class='empty-msg'>(未知)</div>";
        return;
    }

    const div = document.createElement("div");
    div.className = "item-row";
    div.innerHTML = `
        <input type='checkbox' class='verify-check' id='chk_${containerId}' onchange='checkAllChecked()'>
        <label for='chk_${containerId}'>
            ${data.displayName || data.emailAddress} 
            <span class="email-sub">&lt;${data.emailAddress}&gt;</span>
        </label>
    `;
    container.appendChild(div);
}

function renderList(containerId, dataArray) {
    const container = document.getElementById(containerId);
    container.innerHTML = "";

    if (!dataArray || dataArray.length === 0) {
        container.innerHTML = "<div class='empty-msg'>(無)</div>";
        return;
    }

    dataArray.forEach((p, i) => {
        const uniqueId = `${containerId}_${i}`;
        const div = document.createElement("div");
        div.className = "item-row";
        div.innerHTML = `
            <input type='checkbox' class='verify-check' id='${uniqueId}' onchange='checkAllChecked()'>
            <label for='${uniqueId}'>
                ${p.displayName || p.emailAddress}
            </label>
        `;
        container.appendChild(div);
    });
}

function renderAttachments(containerId, dataArray) {
    const container = document.getElementById(containerId);
    container.innerHTML = "";

    if (!dataArray || dataArray.length === 0) {
        container.innerHTML = "<div class='empty-msg'>(無附件)</div>";
        return;
    }

    dataArray.forEach((a, i) => {
        const uniqueId = `att_${i}`;
        const div = document.createElement("div");
        div.className = "item-row";
        div.innerHTML = `
            <input type='checkbox' class='verify-check' id='${uniqueId}' onchange='checkAllChecked()'>
            <label for='${uniqueId}'>📎 ${a.name}</label>
        `;
        container.appendChild(div);
    });
}

window.checkAllChecked = function() {
    const all = document.querySelectorAll(".verify-check");
    let pass = true;
    
    if (all.length === 0) pass = true;
    else {
        all.forEach(c => { 
            if(!c.checked) pass = false; 
        });
    }
    
    if (pass) enableButton();
    else disableButton();
};

// --- 修改重點：按鈕文字設定 ---
function enableButton() {
    const btn = document.getElementById("btnVerify");
    btn.disabled = false;
    btn.classList.add("active");
    // 這裡移除了 unicode icon，並更新了文字
    btn.innerText = "確認完成並送出";
}

function disableButton() {
    const btn = document.getElementById("btnVerify");
    btn.disabled = true;
    btn.classList.remove("active");
    btn.innerText = "請勾選所有項目...";
}
// ----------------------------

function markAsVerified() {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        props.set("isVerified", true);
        
        props.saveAsync((saveResult) => {
            if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                document.getElementById("btnVerify").style.display = "none";
                document.getElementById("status-msg").style.display = "block";
            } else {
                document.getElementById("btnVerify").innerText = "❌ 儲存失敗，請重試";
            }
        });
    });
}