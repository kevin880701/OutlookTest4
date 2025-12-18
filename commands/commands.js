/* global Office */

let sendEvent;

Office.onReady();

function validateSend(event) {
    sendEvent = event;
    
    // 1. 先讀取所有資料 (使用 Promise 包裝以確保讀完才開視窗)
    const item = Office.context.mailbox.item;
    
    const pSubject = new Promise((resolve) => {
        item.subject.getAsync((r) => resolve(r.status === 'succeeded' ? r.value : ''));
    });

    const pTo = new Promise((resolve) => {
        item.to.getAsync((r) => resolve(r.status === 'succeeded' ? r.value : []));
    });

    const pCc = new Promise((resolve) => {
        item.cc.getAsync((r) => resolve(r.status === 'succeeded' ? r.value : []));
    });

    const pAttachments = new Promise((resolve) => {
        item.attachments.getAsync((r) => resolve(r.status === 'succeeded' ? r.value : []));
    });

    // 2. 等全部讀完
    Promise.all([pSubject, pTo, pCc, pAttachments]).then((values) => {
        const [subject, to, cc, attachments] = values;

        // 3. 把資料打包，存入 localStorage (瀏覽器暫存)
        const emailData = {
            subject: subject,
            recipients: [...to, ...cc], // 合併收件人與副本
            attachments: attachments
        };
        
        // 【關鍵】存入暫存，讓 dialog.js 讀取
        localStorage.setItem("emailCheckData", JSON.stringify(emailData));

        // 4. 資料準備好了，現在才打開視窗
        openDialog();
    });
}

function openDialog() {
    const dialogUrl = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html';
    
    Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 50, width: 30, displayInIframe: true },
        dialogCallback
    );
}

function dialogCallback(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // 如果視窗開失敗，直接放行
        sendEvent.completed({ allowEvent: true });
    } else {
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            dialog.close();
            // 清除暫存資料
            localStorage.removeItem("emailCheckData");

            if (arg.message === "SEND_MAIL") {
                sendEvent.completed({ allowEvent: true });
            } else {
                sendEvent.completed({ allowEvent: false });
            }
        });
    }
}