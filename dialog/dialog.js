/* global Office, document, window */

// 1. 定義除錯工具 (一定要放在最上面)
function log(msg) {
    const consoleDiv = document.getElementById("debug-console");
    if (consoleDiv) {
        // 加上時間戳記
        const time = new Date().toLocaleTimeString();
        consoleDiv.innerHTML += `[${time}] ${msg}<br>`;
        consoleDiv.scrollTop = consoleDiv.scrollHeight; // 自動捲動到底部
    }
}

log("JS File Loaded. Waiting for Office.onReady...");

Office.onReady(() => {
    log("Office.onReady triggered! (Office環境載入成功)");

    // 綁定按鈕
    try {
        document.getElementById("btnSend").onclick = () => {
            log("User clicked Send");
            Office.context.ui.messageParent("VERIFIED_PASS");
        };
        document.getElementById("btnCancel").onclick = () => {
            log("User clicked Cancel");
            Office.context.ui.messageParent("CANCEL");
        };
        log("Buttons event listeners attached.");
    } catch (e) {
        log("Error attaching buttons: " + e.message);
    }

    // 開始讀取資料
    try {
        log("Current URL: " + window.location.href);
        
        const urlParams = new URLSearchParams(window.location.search);
        const dataString = urlParams.get('data');

        if (dataString) {
            log("Found 'data' param length: " + dataString.length);
            
            // 嘗試解碼
            const decoded = decodeURIComponent(dataString);
            log("Data decoded successfully.");
            
            // 嘗試解析 JSON
            const data = JSON.parse(decoded);
            log("JSON parsed successfully.");
            log("Recipients count: " + (data.recipients ? data.recipients.length : 0));

            // 開始繪圖
            renderData(data);
            log("renderData finished.");
            
        } else {
            log("❌ ERROR: 'data' parameter is MISSING in URL.");
            document.getElementById("recipients-list").innerText = "錯誤：網址沒有參數";
        }

    } catch (e) {
        log("❌ CRITICAL ERROR: " + e.message);
        document.getElementById("recipients-list").innerText = "程式崩潰：" + e.message;
    }
});

function renderData(data) {
    log("Starting renderData...");
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    
    // 收件人
    const userDomain = "outlook.com"; 
    if (data.recipients && data.recipients.length > 0) {
        data.recipients.forEach((person, index) => {
            const row = document.createElement("div");
            row.className = "item-row";
            
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.className = "verify-check";
            checkbox.id = `recip_${index}`;
            checkbox.onchange = checkAllChecked; // 綁定勾選事件
            
            // 預設勾選內部信箱
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
        log("Recipients rendered.");
    } else {
        container.innerHTML = "無收件人";
        log("No recipients found.");
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
        log("Attachments rendered.");
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