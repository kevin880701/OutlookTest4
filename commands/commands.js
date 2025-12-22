/* global Office */

let sendEvent;

Office.onReady();

function validateSend(event) {
    sendEvent = event;

    // 1. 讀取資料
    const item = Office.context.mailbox.item;
    
    // 使用 Promise 確保讀取完成
    const pSubject = new Promise((resolve) => {
        item.subject.getAsync((r) => resolve(r.status === 'succeeded' ? r.value : ''));
    });

    const pTo = new Promise((resolve) => {
        item.to.getAsync((r) => resolve(r.status === 'succeeded' ? r.value : []));
    });

    // 為了簡化 URL 長度，我們先只測試「主旨」和「收件人」
    // (如果資料太多，URL 會爆掉，那是進階課題)
    Promise.all([pSubject, pTo])
        .then((values) => {
            const [subject, to] = values;

            // 2. 【修改重點】不存 storage，改打包成 JSON 字串
            const dataPackage = {
                subject: subject,
                recipients: to, // 只傳收件人陣列
                attachmentCount: 0 // 暫時省略附件細節
            };

            // 3. 轉成字串並編碼 (因為要放在網址裡，不能有特殊符號)
            const jsonString = encodeURIComponent(JSON.stringify(dataPackage));

            // 4. 把資料串在網址後面 (?data=...)
            const baseUrl = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html';
            const fullUrl = `${baseUrl}?data=${jsonString}`;

            // 5. 打開視窗
            Office.context.ui.displayDialogAsync(
                fullUrl,
                { height: 50, width: 30, displayInIframe: true },
                dialogCallback
            );
        })
        .catch((error) => {
            // 【救命繩】如果上面發生任何錯誤，這裡會接住，並允許發信
            // 這樣就不會發生「無限轉圈圈」的慘劇
            console.error("發生錯誤:", error);
            // 發生錯誤時，選擇放行或擋下 (這裡設為 true 放行以免卡住)
            sendEvent.completed({ allowEvent: true });
        });
}

function dialogCallback(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // 開視窗失敗，放行
        sendEvent.completed({ allowEvent: true });
    } else {
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            dialog.close();
            // 接收視窗傳回來的指令
            if (arg.message === "SEND_MAIL") {
                sendEvent.completed({ allowEvent: true });
            } else {
                sendEvent.completed({ allowEvent: false });
            }
        });
    }
}