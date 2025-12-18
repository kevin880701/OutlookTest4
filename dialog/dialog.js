/* global Office */

Office.onReady(() => {
    // 當 Office 環境準備好，開始讀取資料
    loadItemDetails();

    // 綁定按鈕事件
    document.getElementById("btnSend").onclick = () => sendMessageToParent("SEND_MAIL");
    document.getElementById("btnCancel").onclick = () => sendMessageToParent("CANCEL");
});

function sendMessageToParent(message) {
    // 將結果傳回 commands.js
    Office.context.ui.messageParent(message);
}

function loadItemDetails() {
    // 取得目前的郵件項目
    const item = Office.context.mailbox.item;

    // 1. 顯示主旨
    document.getElementById("subject").innerText = item.subject || "(無主旨)";

    // 2. 處理收件人 (To, Cc)
    let allRecipients = [];
    item.to.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            allRecipients = allRecipients.concat(result.value);
            
            // 接著讀取 CC
            item.cc.getAsync((ccResult) => {
                if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
                    allRecipients = allRecipients.concat(ccResult.value);
                    displayRecipients(allRecipients);
                }
            });
        }
    });

    // 3. 處理附件
    item.attachments.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            displayAttachments(result.value);
        }
        // 資料讀取完畢，顯示畫面
        document.getElementById("loading").style.display = "none";
        document.getElementById("content").style.display = "block";
    });
}

function displayRecipients(recipients) {
    const container = document.getElementById("recipients");
    container.innerHTML = "";

    if (recipients.length === 0) {
        container.innerText = "無收件人";
        return;
    }

    recipients.forEach(person => {
        const div = document.createElement("div");
        div.style.marginBottom = "5px";
        
        // 解析 Email 網域
        const email = person.emailAddress;
        const domain = email.split('@')[1] || "unknown";
        const name = person.displayName;

        // 顯示格式：[gmail.com] 劉浩然 (liu@gmail.com)
        div.innerHTML = `<span class="domain-tag">${domain}</span> <b>${name}</b> <br/><small>&lt;${email}&gt;</small>`;
        container.appendChild(div);
    });
}

function displayAttachments(attachments) {
    const container = document.getElementById("attachments");
    if (attachments.length === 0) {
        container.innerText = "無附件";
        return;
    }
    
    container.innerHTML = "";
    attachments.forEach(att => {
        const div = document.createElement("div");
        // 顯示附件名稱與大小 (如果是檔案)
        div.innerText = `📎 ${att.name} (${att.attachmentType})`;
        container.appendChild(div);
    });
}