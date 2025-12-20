/* global Office, console */

let dialog;
let currentEvent;
let fetchedData = null;
let isDataSent = false; // 【關鍵修正】防止重複發送的旗標

Office.onReady(() => {
  // Init
});

// 1. 攔截器 (維持不變)
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

// 2. 開啟視窗
function openDialog(event) {
    currentEvent = event;
    isDataSent = false; // 重置旗標
    fetchedData = null; // 重置資料

    // A. 8秒強制止血 (防止轉圈圈卡死)
    setTimeout(() => {
        if (currentEvent) {
            currentEvent.completed();
            currentEvent = null;
        }
    }, 8000);

    // B. 開啟視窗
    // 【路徑確認】根據您的截圖，正確路徑應為 /dialog/dialog.html
    // 如果您已確認現在能開視窗，請維持您目前能運作的網址即可
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html'; 
    
    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog Failed:", asyncResult.error.message);
                if (currentEvent) { currentEvent.completed(); currentEvent = null; }
            } else {
                dialog = asyncResult.value;
                // 監聽視窗傳來的訊號
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            }
        }
    );

    // C. 抓資料 (平行處理)
    const item = Office.context.mailbox.item;
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.attachments.getAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
        fetchedData = {
            subject: item.subject,
            recipients: [
                ...to.map(r => ({...r, type: 'To'})),
                ...cc.map(r => ({...r, type: 'Cc'}))
            ],
            attachments: attachments
        };
    });
}

// 3. 處理訊息
function processMessage(arg) {
    const message = arg.message;

    // A. 視窗說它打開了
    if (message === "DIALOG_READY") {
        // 如果資料已經送過了，就忽略後續的 READY 訊號 (防止重複發送)
        if (isDataSent) return;

        // 使用輪詢確保資料抓好了
        const waitForData = setInterval(() => {
            if (fetchedData) {
                clearInterval(waitForData);
                
                // 【關鍵】送出資料給視窗
                dialog.messageChild(JSON.stringify(fetchedData));
                isDataSent = true; // 標記為已發送
                
                // 資料送達後，才停止 Ribbon 上的轉圈圈
                if (currentEvent) {
                    currentEvent.completed();
                    currentEvent = null;
                }
            }
        }, 100);
    }

    // B. 驗證通過
    else if (message === "VERIFIED_PASS") {
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            props.saveAsync(() => dialog.close());
        });
    } 
    // C. 取消
    else if (message === "CANCEL") {
        dialog.close();
    }
}

// 綁定
if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;