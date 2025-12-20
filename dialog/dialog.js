/* global Office, document */

// 簡單的 Log 工具
function log(msg) {
    const consoleDiv = document.getElementById("debug-console");
    if (consoleDiv) {
        const time = new Date().toLocaleTimeString();
        consoleDiv.innerHTML += `[${time}] ${msg}<br>`;
        consoleDiv.scrollTop = consoleDiv.scrollHeight;
    }
}

Office.onReady(() => {
    log("Dialog Ready. Reading Data from Bridge...");

    // 1. 【關鍵】從 CustomProperties 撈資料
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            log("❌ Error loading props: " + result.error.message);
            return;
        }

        const props = result.value;
        const dataString = props.get("bridge_data"); // 取出資料

        if (dataString) {
            log("✅ Data found in Bridge!");
            try {
                const data = JSON.parse(dataString);
                renderData(data);
                
                // (選用) 讀完後可以清除，這裡先保留方便除錯
            } catch (e) {
                log("❌ JSON Parse Error: " + e.message);
            }
        } else {
            log("⚠️ Bridge is empty. (Commands.js didn't save it?)");
            document.getElementById("recipients-list").innerText = "讀取不到資料 (請稍後重試)";
        }
    });

    // 按鈕綁定
    document.getElementById("btnSend").onclick = () => {
        log("Saving verification...");
        // 直接在這裡寫入驗證通過，不依賴 Parent
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true);
            props.saveAsync(() => {
                // 通知 Parent 關閉
                Office.context.ui.messageParent("VERIFIED_PASS");
            });
        });
    };

    document.getElementById("btnCancel").onclick = () => {
        Office.context.ui.messageParent("CANCEL");
    };
});

// --- 渲染邏輯 (保留您原本的樣式) ---
function renderData(data) {
    log("Rendering Data...");
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

            // 預設全選
            checkbox.checked = true;
            
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
            label.innerText = person.displayName || person.emailAddress; // Fallback
            label.innerHTML = html; // Use HTML version
            label.htmlFor = `recip_${index}`;

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
             
             // 預設全選
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