/* global Office, document */

function log(msg) {
    const consoleDiv = document.getElementById("debug-console");
    if (consoleDiv) {
        const time = new Date().toLocaleTimeString();
        consoleDiv.innerHTML += `[${time}] ${msg}<br>`;
        consoleDiv.scrollTop = consoleDiv.scrollHeight;
    }
}

Office.onReady(() => {
    log("Dialog Ready. Loading Data Bridge...");

    // 1. 【關鍵】從 CustomProperties 讀取資料
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            log("❌ Error loading props: " + result.error.message);
            return;
        }

        const props = result.value;
        const dataString = props.get("temp_data"); // 取出膠囊

        if (dataString) {
            log("✅ Data Bridge found!");
            try {
                const data = JSON.parse(dataString);
                renderData(data);
                
                // (選用) 讀完後可以清除，但非必要，下次會被覆蓋
                // props.remove("temp_data");
                // props.saveAsync();
            } catch (e) {
                log("❌ JSON Parse Error: " + e.message);
            }
        } else {
            log("⚠️ Bridge is empty. (Commands.js didn't save it?)");
            document.getElementById("recipients-list").innerText = "錯誤：讀不到信件暫存資料";
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

// 渲染函式 (維持不變)
function renderData(data) {
    log("Rendering Data...");
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    
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

            const label = document.createElement("label");
            label.innerText = person.displayName || person.emailAddress;
            label.htmlFor = `recip_${index}`;

            row.appendChild(checkbox);
            row.appendChild(label);
            container.appendChild(row);
        });
    } else {
        container.innerHTML = "無收件人";
    }

    const attContainer = document.getElementById("attachments-list");
    attContainer.innerHTML = (data.attachments && data.attachments.length > 0) 
        ? `${data.attachments.length} 個附件` 
        : "無附件";
    
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
        btn.style.cursor = "pointer";
    } else {
        btn.style.opacity = "0.5";
        btn.style.cursor = "not-allowed";
    }
}