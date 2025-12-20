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
    log("UI Ready. Waiting for Broadcast...");

    // 1. 註冊接收器
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onParentMessageReceived
    );

    document.getElementById("btnSend").onclick = () => {
        log("Sending VERIFIED_PASS...");
        // 這裡不能寫入屬性(會崩潰)，直接通知 Parent 去寫
        Office.context.ui.messageParent("VERIFIED_PASS");
    };
    
    document.getElementById("btnCancel").onclick = () => {
        Office.context.ui.messageParent("CANCEL");
    };
});

// 當收到 Parent 廣播來的資料時
function onParentMessageReceived(arg) {
    try {
        const message = arg.message;
        const data = JSON.parse(message); 
        
        if (data && data.recipients) {
             log("✅ Data Received! Rendering...");
             renderData(data);
             
             // 告訴 Parent 別再廣播了
             Office.context.ui.messageParent("DATA_RECEIVED");
        }
    } catch (e) {
        log("Error: " + e.message);
    }
}

function renderData(data) {
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    
    if (data.recipients && data.recipients.length > 0) {
        data.recipients.forEach((p, i) => {
            const d = document.createElement("div");
            d.className = "item-row";
            d.innerHTML = `
                <input type='checkbox' checked class='verify-check' id='r_${i}' onchange='checkAllChecked()'>
                <label for='r_${i}'>${p.displayName || p.emailAddress}</label>
            `;
            container.appendChild(d);
        });
    } else {
        container.innerHTML = "無收件人";
    }
    
    // 附件
    const attContainer = document.getElementById("attachments-list");
    attContainer.innerHTML = "";
    if (data.attachments && data.attachments.length > 0) {
        data.attachments.forEach((a, i) => {
            const d = document.createElement("div");
            d.className = "item-row";
            d.innerHTML = `
                <input type='checkbox' checked class='verify-check' id='a_${i}' onchange='checkAllChecked()'>
                <label for='a_${i}'>📎 ${a.name}</label>
            `;
            attContainer.appendChild(d);
        });
    } else {
        attContainer.innerText = "無附件";
    }

    checkAllChecked();
}

// 將 checkAllChecked 綁定到 window 以便 HTML 字串中的 onchange 可以呼叫
window.checkAllChecked = function() {
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
};