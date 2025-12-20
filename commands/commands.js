/* global Office, console */

let dialog;
let currentEvent;

Office.onReady(() => {
  // Init
});

// 1. 攔截器
function validateSend(event) {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        const isVerified = props.get("isVerified");

        if (isVerified === true) {
            props.remove("isVerified");
            props.saveAsync(() => event.completed({ allowEvent: true }));
        } else {
            event.completed({ 
                allowEvent: false, 
                errorMessage: "⚠️ 發送中止：請按「不傳送」，然後點擊上方工具列的【LaunchEvent Test】按鈕進行檢查。" 
            });
        }
    });
}

// 2. 開啟視窗 (URL 傳參版)
function openDialog(event) {
    currentEvent = event;

    // A. 8秒強制止血
    setTimeout(() => {
        if (currentEvent) {
            currentEvent.completed();
            currentEvent = null;
        }
    }, 8000);

    // B. 先抓資料
    const item = Office.context.mailbox.item;
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.attachments.getAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
        
        // 整理資料
        const payload = {
            subject: item.subject || "(無主旨)",
            recipients: [
                ...to.map(r => ({...r, type: 'To'})),
                ...cc.map(r => ({...r, type: 'Cc'}))
            ],
            attachments: attachments
        };

        // 【關鍵】把資料轉成字串，並進行 URL 編碼
        // 雖然有長度限制，但這是最穩定的方法
        const jsonString = JSON.stringify(payload);
        const encodedData = encodeURIComponent(jsonString);
        
        // 組合網址 (請確認您的路徑是否正確)
        const url = `https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html?data=${encodedData}`;
        
        // C. 開啟視窗
        Office.context.ui.displayDialogAsync(
            url, 
            { height: 60, width: 50, displayInIframe: true },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Dialog Failed:", asyncResult.error.message);
                    if (currentEvent) { currentEvent.completed(); currentEvent = null; }
                } else {
                    dialog = asyncResult.value;
                    // 這裡只需要監聽「結果」
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                }
            }
        );
    });
}

// 3. 處理回傳 (只處理驗證結果)
function processMessage(arg) {
    const message = arg.message;

    if (message === "VERIFIED_PASS") {
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            props.saveAsync(() => {
                dialog.close();
                if (currentEvent) { currentEvent.completed(); currentEvent = null; }
            });
        });
    } else if (message === "CANCEL") {
        dialog.close();
        if (currentEvent) { currentEvent.completed(); currentEvent = null; }
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;