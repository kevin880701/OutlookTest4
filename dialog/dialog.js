/* global Office, document */

// 0. 定義除錯工具 (把訊息印在黑色框框)
function log(msg) {
    const consoleDiv = document.getElementById("debug-console");
    if (consoleDiv) {
        const time = new Date().toLocaleTimeString();
        consoleDiv.innerHTML += `[${time}] ${msg}<br>`;
        consoleDiv.scrollTop = consoleDiv.scrollHeight; // 自動捲動到底部
    }
}

log("JS File Loaded. Waiting for Office.onReady...");

Office.onReady(() => {
    log("Office.onReady triggered! (環境載入成功)");

    // 1. 註冊接收器
    try {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onParentMessageReceived
        );
        log("Handler registered. Waiting for Parent to broadcast data...");
    } catch (e) {
        log("❌ Error registering handler: " + e.message);
    }

    // 綁定按鈕
    document.getElementById("btnSend").onclick = () => {
        log("User clicked Send");
        Office.context.ui.messageParent("VERIFIED_PASS");
    };
    document.getElementById("btnCancel").onclick = () => {
        log("User clicked Cancel");
        Office.context.ui.messageParent("CANCEL");
    };
});

// 當收到 Parent 傳來的資料時
function onParentMessageReceived(arg) {
    // log("Received message from Parent!"); // 避免洗版，先註解掉
    try {
        const message = arg.message;
        // log("Raw message length: " + message.length);

        const data = JSON.parse(message);
        
        // 確保資料有效才渲染
        if (data && data.recipients) {
             // 為了避免重複渲染導致閃爍，可以加個檢查
             // 這裡直接渲染並記錄
             renderData(data);
             
             // 回報給 Parent 說收到了 (選用)
             Office.context.ui.messageParent("DATA_RECEIVED");
        }
    } catch (e) {
        log("❌ Data parse error: " + e.message);
    }
}

let isRendered = false; // 防止重複渲染洗版 Log

function renderData(data) {
    if(!isRendered) {
        log("✅ Rendering Data...");
        log(`Recipients: ${data.recipients.length}, Attachments: ${data.attachments.length}`);
        isRendered = true; // 鎖定，避免一直印 Log
    }

    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    const userDomain = "outlook.com"; 

    if (data.recipients && data.recipients.length > 0) {
        data.recipients.forEach((person, index) => {
            const row = document.createElement("div");
            row.className = "item-row";
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.className = "verify-check";
            checkbox.id = `recip_${index}`;
            checkbox.onchange = checkAllChecked;

            const email = person.emailAddress || "";
            let personDomain = "";
            if (email.includes("@")) personDomain = email.split('@')[1];
            const isExternal = personDomain && personDomain !== userDomain;

            let html = `<b>${person.displayName || "Unknown"}</b> <br><small>${email}</small>`;
            if (isExternal) {
                html += ` <span class="external-tag">External</span>`;
                checkbox.checked = false; 
            } else {
                checkbox.checked = true; 
            }
            const label = document.createElement("label");
            label.htmlFor = `recip_${index}`;
            label.innerHTML = html;
            row.appendChild(checkbox);
            row.appendChild(label);
            container.appendChild(row);
        });
    } else {
        container.innerHTML = "無收件人";
    }
    
    // 附件
    const attContainer = document.getElementById("attachments-list");
    attContainer.innerHTML = "";
    if (data.attachments && data.attachments.length > 0) {
        data.attachments.forEach((att, index) => {
             const row = document.createElement("div");
             row.className = "item-row";
             const checkbox = document.createElement("input");
             checkbox.type = "checkbox";
             checkbox.className = "verify-check";
             checkbox.id = `att_${index}`;
             checkbox.onchange = checkAllChecked;
             const label = document.createElement("label");
             label.htmlFor = `att_${index}`;
             label.innerText = `📎 ${att.name}`;
             row.appendChild(checkbox);
             row.appendChild(label);
             attContainer.appendChild(row);
        });
    } else {
        attContainer.innerText = "無附件";
    }
    checkAllChecked(); 
}

function checkAllChecked() {
    const all = document.querySelectorAll(".verify-check");
    let pass = true;
    all.forEach(c => { if(!c.checked) pass = false; });
    
    const btn = document.getElementById("btnSend");
    if (all.length === 0) pass = true;

    btn.disabled = !pass;
    if (pass) {
        btn.style.opacity = "1";
        btn.classList.add("active");
    } else {
        btn.style.opacity = "0.5";
        btn.classList.remove("active");
    }
}