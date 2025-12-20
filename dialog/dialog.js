/* global Office, document */

function log(msg) {
    const consoleDiv = document.getElementById("debug-console");
    if (consoleDiv) {
        const time = new Date().toLocaleTimeString();
        consoleDiv.innerHTML += `[${time}] ${msg}<br>`;
        consoleDiv.scrollTop = consoleDiv.scrollHeight;
    }
}

let pollInterval; // 輪詢計時器

Office.onReady(() => {
    log("Dialog Opened. Start Polling for Data...");

    // 1. 啟動輪詢：每 1000ms (1秒) 檢查一次資料
    pollInterval = setInterval(checkBridgeData, 1000);
    
    // 先立刻檢查一次
    checkBridgeData();

    // 按鈕綁定
    document.getElementById("btnSend").onclick = () => {
        log("Saving verification...");
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true);
            props.saveAsync(() => {
                Office.context.ui.messageParent("VERIFIED_PASS");
            });
        });
    };

    document.getElementById("btnCancel").onclick = () => {
        Office.context.ui.messageParent("CANCEL");
    };
});

// 檢查橋接資料函式
function checkBridgeData() {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            log("❌ Load props failed: " + result.error.message);
            return;
        }

        const props = result.value;
        const dataString = props.get("bridge_data");

        if (dataString) {
            log("✅ Data Found! Stopping poll.");
            
            // 讀到了！停止輪詢
            clearInterval(pollInterval);
            
            try {
                const data = JSON.parse(dataString);
                renderData(data);
            } catch (e) {
                log("❌ JSON Parse Error: " + e.message);
            }
        } else {
            log("⏳ Waiting for data... (Commands.js is saving)");
        }
    });
}

// 渲染函式 (維持不變)
function renderData(data) {
    log("Rendering UI...");
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
            checkbox.onchange = checkAllChecked;
            checkbox.checked = true; // 預設全選
            
            // 判斷外部信箱
            const email = person.emailAddress || "";
            let personDomain = "";
            if (email.includes("@")) personDomain = email.split('@')[1];
            const isExternal = personDomain && personDomain !== userDomain;

            let html = `<b>${person.displayName || "Unknown"}</b> <br><small>${email}</small>`;
            if (isExternal) {
                html += ` <span class="external-tag">External</span>`;
                checkbox.checked = false; 
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
             checkbox.checked = true;

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