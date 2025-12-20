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

// 2. 開啟視窗 (URL 傳值版)
function openDialog(event) {
    currentEvent = event;

    // A. 8秒後強制停止轉圈圈 (安全保險)
    setTimeout(() => {
        if (currentEvent) {
            currentEvent.completed();
            currentEvent = null;
        }
    }, 8000);

    // B. 抓取資料
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

        // 【關鍵】將資料編碼放入 URL
        let jsonString = JSON.stringify(payload);
        let encodedData = encodeURIComponent(jsonString);
        
        // 檢查網址長度 (Outlook Mac 限制約 2048 chars)
        // 如果太長，我們就只傳部分資訊或錯誤提示，避免視窗打不開
        if (encodedData.length > 1500) {
            console.warn("Data too long, truncating...");
            payload.recipients = []; // 清空收件人，改為顯示提示
            payload.subject = "⚠️ 資料量過大，請手動檢查收件人";
            jsonString = JSON.stringify(payload);
            encodedData = encodeURIComponent(jsonString);
        }
        
        // 組合網址 (請確認您的路徑是否正確)
        const url = `https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html?data=${encodedData}`;
        
        // C. 開啟視窗
        Office.context.ui.displayDialogAsync(
            url, 
            { height: 60, width: 50, displayInIframe: true },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Dialog Failed:", asyncResult.error.message);
                } else {
                    dialog = asyncResult.value;
                    // 這裡只監聽「驗證通過」或「取消」的訊號
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                }
            }
        );
    });
}

// 3. 處理回傳 (只處理結果)
function processMessage(arg) {
    const message = arg.message;

    if (message === "VERIFIED_PASS") {
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            props.saveAsync(() => {
                dialog.close();
                // 收到確認後，結束轉圈圈 (使用者需要再按一次傳送)
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