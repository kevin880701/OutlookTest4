/* global Office */

Office.onReady(() => {
    // 綁定按鈕
    document.getElementById("btnSend").onclick = () => sendMessageToParent("SEND_MAIL");
    document.getElementById("btnCancel").onclick = () => sendMessageToParent("CANCEL");

    // 從 localStorage 讀取剛剛 commands.js 存好的資料
    try {
        const dataJson = localStorage.getItem("emailCheckData");
        if (dataJson) {
            const data = JSON.parse(dataJson);
            renderData(data);
        } else {
            document.getElementById("loading").innerText = "無法讀取郵件資料 (Storage Empty)";
        }
    } catch (e) {
        document.getElementById("loading").innerText = "發生錯誤: " + e.message;
    }
});

function sendMessageToParent(message) {
    Office.context.ui.messageParent(message);
}

function renderData(data) {
    // 1. 顯示主旨
    document.getElementById("subject").innerText = data.subject || "(無主旨)";

    // 2. 顯示收件人
    const recipientContainer = document.getElementById("recipients");
    recipientContainer.innerHTML = "";
    
    if (data.recipients && data.recipients.length > 0) {
        data.recipients.forEach(person => {
            const div = document.createElement("div");
            div.style.marginBottom = "5px";
            const email = person.emailAddress;
            const domain = email.split('@')[1] || "unknown";
            const name = person.displayName;
            div.innerHTML = `<span class="domain-tag">${domain}</span> <b>${name}</b> <br/><small>&lt;${email}&gt;</small>`;
            recipientContainer.appendChild(div);
        });
    } else {
        recipientContainer.innerText = "無收件人";
    }

    // 3. 顯示附件
    const attContainer = document.getElementById("attachments");
    attContainer.innerHTML = "";

    if (data.attachments && data.attachments.length > 0) {
        data.attachments.forEach(att => {
            const div = document.createElement("div");
            div.innerText = `📎 ${att.name}`;
            attContainer.appendChild(div);
        });
    } else {
        attContainer.innerText = "無附件";
    }

    // 隱藏 Loading，顯示內容
    document.getElementById("loading").style.display = "none";
    document.getElementById("content").style.display = "block";
}