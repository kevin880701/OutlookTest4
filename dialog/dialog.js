/* global Office, document, window */

// 除錯工具
function log(msg) {
    const consoleDiv = document.getElementById("debug-console");
    if (consoleDiv) {
        const time = new Date().toLocaleTimeString();
        consoleDiv.innerHTML += `[${time}] ${msg}<br>`;
        consoleDiv.scrollTop = consoleDiv.scrollHeight;
    }
}

log("JS Loaded. Initializing...");

Office.onReady(() => {
    log("Office.onReady triggered.");

    // 按鈕綁定
    document.getElementById("btnSend").onclick = () => {
        log("Sending VERIFIED_PASS...");
        Office.context.ui.messageParent("VERIFIED_PASS");
    };
    document.getElementById("btnCancel").onclick = () => {
        log("Sending CANCEL...");
        Office.context.ui.messageParent("CANCEL");
    };

    // 【關鍵】從 URL 解析資料
    try {
        log("Checking URL parameters...");
        const urlParams = new URLSearchParams(window.location.search);
        const dataString = urlParams.get('data');

        if (dataString) {
            log("Data found in URL! Length: " + dataString.length);
            
            const decoded = decodeURIComponent(dataString);
            const data = JSON.parse(decoded);
            
            log("JSON parsed. Recipients: " + (data.recipients ? data.recipients.length : 0));
            renderData(data); // 畫出介面
            
        } else {
            log("❌ No data found in URL. (Did commands.js send it?)");
            document.getElementById("recipients-list").innerText = "錯誤：網址沒有資料";
        }
    } catch (e) {
        log("❌ Error parsing data: " + e.message);
        document.getElementById("recipients-list").innerText = "資料解析失敗";
    }
});

// 渲染函式 (維持不變)
function renderData(data) {
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