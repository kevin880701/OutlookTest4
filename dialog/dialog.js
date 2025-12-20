/* global Office, document */

let handshakeInterval; // 用來存定時器的變數

Office.onReady(() => {
    // 1. 註冊接收器：準備接收來自 Parent 的資料
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onParentMessageReceived
    );

    // 2. 【關鍵修正】啟動「奪命連環 Call」
    // 每 1000 毫秒 (1秒) 喊一次 DIALOG_READY，確保 Parent 一定聽得到
    // 這是解決 "一直 Loading" 的核心關鍵
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
    // 3. 【關鍵修正】收到資料了，停止喊話
    if (handshakeInterval) {
        clearInterval(handshakeInterval);
        handshakeInterval = null;
    }

    try {
        const message = arg.message;
        const data = JSON.parse(message); // 解析資料
        
        // 簡單檢查資料是否正確
        if (data && data.subject !== undefined) {
             renderData(data); // 渲染畫面
        }
    } catch (e) {
        document.getElementById("recipients-list").innerText = "資料錯誤: " + e.message;
    }
}

// 渲染函式 (維持不變，請保留您原本的這段代碼)
function renderData(data) {
    // ... 請保留您原本的 renderData 內容 ...
    // (為了版面簡潔，這裡省略，請直接使用您原本寫好的渲染邏輯)
    
    // 這裡幫您補上開頭幾行，避免您複製貼上時漏掉
    document.getElementById("subject").innerText = data.subject || "(無主旨)";
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    // ... 接續您的渲染代碼 ...
    
    // 記得這一行要在 renderData 裡：
    // renderAttachments(data.attachments);
    // checkAllChecked();
    
    // 為了讓您方便測試，我直接把簡單版渲染邏輯附在下面，您可以選擇是否覆蓋：
    const recipientContainer = document.getElementById("recipients-list");
    recipientContainer.innerHTML = "";
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
            recipientContainer.appendChild(row);
        });
    } else {
        recipientContainer.innerHTML = "無收件人";
    }

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