/* global Office, document */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        try {
            loadItemData();
            document.getElementById("btnVerify").onclick = markAsVerified;
        } catch (e) {
            logError("Init Error: " + e.message);
        }
    }
});

function logError(msg) {
    const el = document.getElementById("error-log");
    el.style.display = "block";
    el.innerText += "❌ " + msg + "\n";
    console.error(msg);
}

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

    const safeGet = (apiCall) => new Promise(resolve => {
        try {
            apiCall(result => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
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
        
        to = to || [];
        cc = cc || [];
        bcc = bcc || [];
        attachments = attachments || [];

        const senderEmail = (from && from.emailAddress) ? from.emailAddress : "";
        const senderDomain = getDomain(senderEmail);
        
        renderSender("from-container", from);

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

    const groups = {};
    dataArray.forEach(p => {
        const domain = getDomain(p.emailAddress);
        if (!groups[domain]) groups[domain] = [];
        groups[domain].push(p);
    });

    // 排序：External 排前面
    const sortedDomains = Object.keys(groups).sort((a, b) => {
        const aIsExt = a !== senderDomain;
        const bIsExt = b !== senderDomain;
        return bIsExt - aIsExt; 
    });

    sortedDomains.forEach(domain => {
        const isExternal = domain !== senderDomain;
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
            
            // --- 核心修改 ---
            // 1. 所有人都有 checkbox
            // 2. 如果是 External -> 預設不勾 ("")
            // 3. 如果是 Internal -> 預設勾選 ("checked")
            const checkedState = isExternal ? "" : "checked";
            
            rowDiv.innerHTML = `
                <input type='checkbox' class='verify-check' ${checkedState} onchange='checkAllChecked()'>
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

// 檢查邏輯：現在 Internal 也有 checkbox，
// 所以如果使用者手動把 Internal 取消勾選，這裡會回傳 false，按鈕會變回 Disabled (符合預期)
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
    
    // 這裡的文字顯示 "未勾選項目"，因為現在內部也可以被取消勾選
    btn.innerText = uncheckCount > 0 ? `尚有 ${uncheckCount} 個項目未確認` : "請勾選所有項目...";
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