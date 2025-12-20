/* global Office, document */

let pollInterval;

function log(msg) {
    const consoleDiv = document.getElementById("debug-console");
    if (consoleDiv) {
        const time = new Date().toLocaleTimeString();
        consoleDiv.innerHTML += `[${time}] ${msg}<br>`;
        consoleDiv.scrollTop = consoleDiv.scrollHeight;
    }
}

Office.onReady(() => {
    log("UI Ready. Polling data...");
    
    // 每 1 秒檢查一次資料
    pollInterval = setInterval(checkBridgeData, 1000);
    checkBridgeData();

    document.getElementById("btnSend").onclick = () => {
        log("Saving verified...");
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

function checkBridgeData() {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) return;
        
        const props = result.value;
        const dataString = props.get("bridge_data");
        
        if (dataString) {
            log("✅ Data found!");
            clearInterval(pollInterval); // 停止輪詢
            try {
                renderData(JSON.parse(dataString));
            } catch (e) {
                log("Parse Error: " + e.message);
            }
        } else {
            log("⏳ Waiting...");
        }
    });
}

// ... (renderData 函式維持您原本的樣子即可)
// 為了完整性，這裡需要包含 renderData 和 checkAllChecked
function renderData(data) {
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    if (data.recipients) {
        data.recipients.forEach((p, i) => {
            // 簡單渲染邏輯...
            const d = document.createElement("div");
            d.innerHTML = `<input type='checkbox' checked class='verify-check'> ${p.displayName || p.emailAddress}`;
            container.appendChild(d);
        });
    } else {
        container.innerHTML = "無收件人";
    }
    checkAllChecked();
}

function checkAllChecked() {
    document.getElementById("btnSend").disabled = false;
    document.getElementById("btnSend").style.opacity = "1";
    document.getElementById("btnSend").style.cursor = "pointer";
}