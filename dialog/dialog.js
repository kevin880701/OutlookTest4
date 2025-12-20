/* global Office, document */

// 定義 log 函式方便除錯 (會顯示在黑色框框)
function log(msg) {
    const consoleDiv = document.getElementById("debug-console");
    if (consoleDiv) {
        const time = new Date().toLocaleTimeString();
        consoleDiv.innerHTML += `[${time}] ${msg}<br>`;
        consoleDiv.scrollTop = consoleDiv.scrollHeight;
    }
}

Office.onReady(() => {
    log("Ready! Waiting for Broadcast...");

    // 1. 【關鍵修正】註冊接收器，準備被資料「砸中」
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onParentMessageReceived
    );

    // 按鈕綁定
    document.getElementById("btnSend").onclick = () => {
        if (!document.getElementById("btnSend").disabled) {
            Office.context.ui.messageParent("VERIFIED_PASS");
        }
    };
    document.getElementById("btnCancel").onclick = () => {
        Office.context.ui.messageParent("CANCEL");
    };
});

// 當收到 Parent 廣播來的資料時
function onParentMessageReceived(arg) {
    try {
        const message = arg.message;
        const data = JSON.parse(message); // 解析資料
        
        // 確保資料有效
        if (data && data.recipients) {
             log("Data Received! Rendering...");
             renderData(data); // 渲染畫面
             
             // 禮貌性回覆：我收到了，別再廣播了
             Office.context.ui.messageParent("DATA_RECEIVED");
        }
    } catch (e) {
        log("Error: " + e.message);
    }
}

// --- 您的渲染邏輯 (維持不變) ---
function renderData(data) {
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    
    // 這裡我簡化了顯示邏輯，請替換回您完整的 renderData 代碼
    // 重點是確認 data 進來了
    
    if (data.recipients && data.recipients.length > 0) {
        data.recipients.forEach((person, index) => {
            const row = document.createElement("div");
            row.className = "item-row";
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.className = "verify-check";
            checkbox.id = `recip_${index}`;
            checkbox.onchange = checkAllChecked;

            const label = document.createElement("label");
            label.innerText = person.displayName || person.emailAddress;
            label.htmlFor = `recip_${index}`;
            checkbox.checked = true; // 預設勾選

            row.appendChild(checkbox);
            row.appendChild(label);
            container.appendChild(row);
        });
    } else {
        container.innerHTML = "無收件人";
    }

    // 附件 (略，維持您的代碼)
    const attContainer = document.getElementById("attachments-list");
    attContainer.innerHTML = "無附件"; 
    
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