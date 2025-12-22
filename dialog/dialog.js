/* global Office */

Office.onReady(() => {
    document.getElementById("btnSend").onclick = () => sendMessageToParent("SEND_MAIL");
    document.getElementById("btnCancel").onclick = () => sendMessageToParent("CANCEL");

    try {
        // 1. 【修改重點】從網址列 (URL) 抓取參數
        const urlParams = new URLSearchParams(window.location.search);
        const dataString = urlParams.get('data'); // 抓取 ?data= 後面的東西

        if (dataString) {
            // 2. 解碼並還原成物件
            const data = JSON.parse(decodeURIComponent(dataString));
            renderData(data);
        } else {
            document.getElementById("loading").innerText = "網址內沒有資料";
        }
    } catch (e) {
        document.getElementById("loading").innerText = "解析錯誤: " + e.message;
    }
});

function sendMessageToParent(message) {
    Office.context.ui.messageParent(message);
}

function renderData(data) {
    // 顯示主旨
    document.getElementById("subject").innerText = data.subject || "(無主旨)";

    // 顯示收件人
    const recipientContainer = document.getElementById("recipients");
    recipientContainer.innerHTML = "";
    
    if (data.recipients && data.recipients.length > 0) {
        data.recipients.forEach(person => {
            const div = document.createElement("div");
            div.style.marginBottom = "5px";
            const email = person.emailAddress;
            const name = person.displayName;
            div.innerHTML = `<b>${name}</b> <br/><small>&lt;${email}&gt;</small>`;
            recipientContainer.appendChild(div);
        });
    } else {
        recipientContainer.innerText = "無收件人";
    }

    // 附件 (因為我們剛剛沒傳附件詳情，這裡先寫死或顯示數量)
    document.getElementById("attachments").innerText = "附件檢查暫略";

    // 隱藏 Loading
    document.getElementById("loading").style.display = "none";
    document.getElementById("content").style.display = "block";
}