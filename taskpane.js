/* global Office, document */

Office.onReady(() => {
    loadItemData();
    document.getElementById("btnVerify").onclick = markAsVerified;
});

function loadItemData() {
    const item = Office.context.mailbox.item;

    // 同時讀取所有需要的欄位
    Promise.all([
        new Promise(r => item.from.getAsync(x => r(x.value))),       // 寄件人 (通常是物件)
        new Promise(r => item.to.getAsync(x => r(x.value || []))),   // 收件人 (陣列)
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),   // 副本 (陣列)
        new Promise(r => item.bcc.getAsync(x => r(x.value || []))),  // 密件副本 (陣列)
        new Promise(r => item.getAttachmentsAsync(x => r(x.value || []))) // 附件 (陣列)
    ]).then(([from, to, cc, bcc, attachments]) => {
        
        // 渲染各個區塊
        renderSingleItem("from-list", from);
        renderList("to-list", to);
        renderList("cc-list", cc);
        renderList("bcc-list", bcc);
        renderAttachments("attachments-list", attachments);

        // 如果所有欄位都是空的 (極端情況)，也要檢查一下按鈕狀態
        checkAllChecked();

    }).catch(err => {
        console.error(err);
        document.body.innerHTML = "<h3 style='color:red'>讀取錯誤</h3>" + err.message;
    });
}

/**
 * 渲染單一項目 (用於 From)
 */
function renderSingleItem(containerId, data) {
    const container = document.getElementById(containerId);
    container.innerHTML = "";

    if (!data) {
        container.innerHTML = "<div class='empty-msg'>(未知)</div>";
        return;
    }

    // 建立 Checkbox
    const div = document.createElement("div");
    div.className = "item-row";
    div.innerHTML = `
        <input type='checkbox' class='verify-check' id='chk_${containerId}' onchange='checkAllChecked()'>
        <label for='chk_${containerId}'>
            ${data.displayName || data.emailAddress} <br>
            <span style="font-size:11px; color:#666;">&lt;${data.emailAddress}&gt;</span>
        </label>
    `;
    container.appendChild(div);
}

/**
 * 渲染人員列表 (用於 To, Cc, Bcc)
 */
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

/**
 * 渲染附件列表 (邏輯類似，但顯示名稱欄位不同)
 */
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

// 檢查是否全部勾選 (這個邏輯不用變，它會自動抓頁面上所有的 .verify-check)
window.checkAllChecked = function() {
    const all = document.querySelectorAll(".verify-check");
    let pass = true;
    
    if (all.length === 0) pass = true; // 如果完全沒有任何需要檢查的東西
    else {
        all.forEach(c => { 
            if(!c.checked) pass = false; 
        });
    }
    
    if (pass) enableButton();
    else disableButton();
};

function enableButton() {
    const btn = document.getElementById("btnVerify");
    btn.disabled = false;
    btn.classList.add("active");
    btn.innerText = "✅ 確認無誤 (解除鎖定)";
}

function disableButton() {
    const btn = document.getElementById("btnVerify");
    btn.disabled = true;
    btn.classList.remove("active");
    btn.innerText = "請勾選所有項目...";
}

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