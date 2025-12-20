/* global Office, document */

let handshakeInterval; // 用來存定時器的變數

Office.onReady(() => {
    // 1. 註冊接收器
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onParentMessageReceived
    );

    // 2. 【關鍵修改】啟動「奪命連環 Call」
    // 每 1000 毫秒 (1秒) 喊一次 DIALOG_READY，確保 Parent 一定聽得到
    handshakeInterval = setInterval(() => {
        try {
            Office.context.ui.messageParent("DIALOG_READY");
            console.log("Sent: DIALOG_READY");
        } catch (e) {
            console.error("Connection not ready yet...");
        }
    }, 1000);

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

// 當收到 Parent 傳來的資料時
function onParentMessageReceived(arg) {
    // 3. 【關鍵修改】收到資料了，停止喊話
    if (handshakeInterval) {
        clearInterval(handshakeInterval);
        handshakeInterval = null;
    }

    try {
        const message = arg.message;
        const data = JSON.parse(message); // 解析資料
        
        // 簡單檢查資料是否正確 (避免收到空字串)
        if (data && data.subject !== undefined) {
             renderData(data); // 渲染畫面
        }
    } catch (e) {
        document.getElementById("recipients-list").innerText = "資料錯誤: " + e.message;
    }
}

// 渲染函式 (不用改，照舊)
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

            let html = `<b>[${person.type}]</b> ${person.displayName || "Unknown"} <br><small>${email}</small>`;
            if (isExternal) {
                html += ` <span class="external-tag" style="color:red; border:1px solid red; font-size:10px; margin-left:5px;">External</span>`;
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
    
    // 附件渲染
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
        btn.style.cursor = "pointer";
        btn.innerText = "確認完畢，允許發送";
    } else {
        btn.style.opacity = "0.5";
        btn.style.cursor = "not-allowed";
        btn.innerText = "請勾選所有項目";
    }
}