/* global Office, document */

Office.onReady((info) => {
    // 確保 DOM 載入後才執行
    if (info.host === Office.HostType.Outlook) {
        // 使用 try-catch 確保即使初始化失敗也能顯示錯誤
        try {
            loadItemData();
            document.getElementById("btnVerify").onclick = markAsVerified;
        } catch (e) {
            logError("Init Error: " + e.message);
        }
    }
});

// 錯誤顯示 helper
function logError(msg) {
    const el = document.getElementById("error-log");
    el.style.display = "block";
    el.innerText += "❌ " + msg + "\n";
    console.error(msg);
}

// 取得 Email 的網域 (強化防呆)
function getDomain(email) {
    if (!email || typeof email !== 'string') return "unknown";
    if (!email.includes("@")) return "unknown";
    return email.split("@")[1].toLowerCase().trim();
}

function loadItemData() {
    const item = Office.context.mailbox.item;

    if (!item) {
        logError("無法讀取郵件物件 (Item is null)");
        return;
    }

    // 定義一個安全的 Promise wrapper，避免單一失敗導致全部卡住
    const safeGet = (apiCall) => new Promise(resolve => {
        try {
            apiCall(result => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    // 即使失敗也 resolve null，不要 reject 導致全部停住
                    console.warn("API Failed:", result.error);
                    resolve(null);
                }
            });
        } catch (e) {
            console.error("API Call Error:", e);
            resolve(null);
        }
    });

    Promise.all([
        safeGet(cb => item.from.getAsync(cb)),
        safeGet(cb => item.to.getAsync(cb)),
        safeGet(cb => item.cc.getAsync(cb)),
        safeGet(cb => item.bcc.getAsync(cb)),
        safeGet(cb => item.getAttachmentsAsync(cb))
    ]).then(([from, to, cc, bcc, attachments]) => {
        
        // 確保陣列不為 null (Fallback to empty array)
        to = to || [];
        cc = cc || [];
        bcc = bcc || [];
        attachments = attachments || [];

        // 1. 獲取寄件人網域
        // 注意：新草稿有時 from 為 null，預設為空字串，這會導致所有人都變成 External (這是安全的做法)
        const senderEmail = (from && from.emailAddress) ? from.emailAddress : "";
        const senderDomain = getDomain(senderEmail);
        
        // 渲染寄件人
        renderSender("from-container", from);

        // 2. 渲染列表
        renderGroupedList("to-list", to, senderDomain);
        renderGroupedList("cc-list", cc, senderDomain);
        renderGroupedList("bcc-list", bcc, senderDomain);
        
        renderAttachments("attachments-list", attachments);

        checkAllChecked();

    }).catch(err => {
        logError("Load Data Error: " + err.message);
    });
}

function renderSender(containerId, data) {
    const container = document.getElementById(containerId);
    if (!data) {
        // 如果抓不到寄件者，顯示提示但不報錯
        container.innerHTML = "<div class='empty-msg'>寄件者資訊讀取中或未設定</div>";
        return;
    }
    container.innerHTML = `
        <div class="safe-icon">👤</div>
        <div class="item-content">
            <div class="name">${data.displayName || data.emailAddress}</div>
            <div class="email">${data.emailAddress}</div>
        </div>
    `;
}

function renderGroupedList(containerId, dataArray, senderDomain) {
    const container = document.getElementById(containerId);
    container.innerHTML = "";

    if (!dataArray || dataArray.length === 0) {
        container.innerHTML = "<div class='empty-msg'>(無)</div>";
        return;
    }

    // 分組邏輯
    const groups = {};
    dataArray.forEach(p => {
        const domain = getDomain(p.emailAddress);
        if (!groups[domain]) groups[domain] = [];
        groups[domain].push(p);
    });

    // 排序：External 在前
    const sortedDomains = Object.keys(groups).sort((a, b) => {
        const aIsExt = a !== senderDomain;
        const bIsExt = b !== senderDomain;
        return bIsExt - aIsExt; 
    });

    sortedDomains.forEach(domain => {
        const isExternal = domain !== senderDomain; // 如果 senderDomain 是空字串，這裡會全變成 true (安全)
        const recipients = groups[domain];

        const groupDiv = document.createElement("div");
        groupDiv.className = "domain-group";

        const headerDiv = document.createElement("div");
        headerDiv.className = "domain-header";
        
        const tagHtml = isExternal 
            ? `<span class="tag external">External</span>` 
            : `<span class="tag internal">Internal</span>`;
        
        headerDiv.innerHTML = `<span>@${domain}</span> ${tagHtml}`;
        groupDiv.appendChild(headerDiv);

        recipients.forEach((p, i) => {
            const rowDiv = document.createElement("div");
            rowDiv.className = "item-row";
            
            // 只有 External 才有 Checkbox
            let controlHtml = "";
            if (isExternal) {
                controlHtml = `<input type='checkbox' class='verify-check' onchange='checkAllChecked()'>`;
            } else {
                controlHtml = `<span class="safe-icon">🛡️</span>`;
            }

            rowDiv.innerHTML = `
                ${controlHtml}
                <div class="item-content">
                    <div class="name">${p.displayName || p.emailAddress}</div>
                    <div class="email">${p.emailAddress}</div>
                </div>
            `;
            groupDiv.appendChild(rowDiv);
        });

        container.appendChild(groupDiv);
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
        const div = document.createElement("div");
        div.className = "item-row";
        div.innerHTML = `
            <input type='checkbox' class='verify-check' id='att_${i}' onchange='checkAllChecked()'>
            <div class="item-content">
                <label for='att_${i}' style="cursor:pointer" class="name">📎 ${a.name}</label>
            </div>
        `;
        container.appendChild(div);
    });
}

window.checkAllChecked = function() {
    const allCheckboxes = document.querySelectorAll(".verify-check");
    let pass = true;
    
    if (allCheckboxes.length === 0) {
        pass = true;
    } else {
        allCheckboxes.forEach(c => { 
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
    btn.innerText = "確認完成並送出";
}

function disableButton() {
    const btn = document.getElementById("btnVerify");
    btn.disabled = true;
    btn.classList.remove("active");
    
    const all = document.querySelectorAll(".verify-check");
    let uncheckCount = 0;
    all.forEach(c => { if(!c.checked) uncheckCount++; });
    
    btn.innerText = uncheckCount > 0 ? `請檢查外部收件人 (${uncheckCount})` : "請勾選所有項目...";
}

function markAsVerified() {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        props.set("isVerified", true);
        props.saveAsync((saveResult) => {
            if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                document.getElementById("btn-area").style.display = "none";
                document.getElementById("status-msg").style.display = "block";
            } else {
                logError("儲存失敗: " + saveResult.error.message);
            }
        });
    });
}