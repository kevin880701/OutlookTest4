/* global Office, console */

let dialog;
let currentEvent; // 暫存 event 用
let fetchedData = null; // 暫存抓到的資料

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

// 2. 開啟視窗 (核心修改：不存 Storage，改用即時傳遞)
function openDialog(event) {
    currentEvent = event; // 把 event 存起來，等資料送出後再結束

    // A. 設定 8 秒強制止血機制 (防止轉圈圈卡死)
    setTimeout(() => {
        if (currentEvent) {
            currentEvent.completed();
            currentEvent = null;
        }
    }, 8000);

    // B. 立刻開啟視窗
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html';
    
    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog Failed:", asyncResult.error.message);
                // 開不起來就直接結束
                if (currentEvent) { currentEvent.completed(); currentEvent = null; }
            } else {
                dialog = asyncResult.value;
                // 監聽視窗傳來的訊號
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            }
        }
    );

    // C. 同時開始抓資料 (平行處理)
    const item = Office.context.mailbox.item;
    
    // 使用 Promise 確保資料都抓回來
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.attachments.getAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
        // 資料打包
        fetchedData = {
            subject: item.subject,
            recipients: [
                ...to.map(r => ({...r, type: 'To'})),
                ...cc.map(r => ({...r, type: 'Cc'}))
            ],
            attachments: attachments
        };
        // 如果 Dialog 已經開好在等了 (很少見)，這裡可以補發，但主要交給 processMessage 處理
    });
}

// 3. 處理訊息 (交握邏輯)
function processMessage(arg) {
    const message = arg.message;

    // A. 視窗說它打開了，跟我們要資料
    if (message === "DIALOG_READY") {
        // 使用輪詢確保資料已經抓好 (因為 openDialog 裡的 Promise 可能比視窗開啟慢)
        const waitForData = setInterval(() => {
            if (fetchedData) {
                clearInterval(waitForData);
                
                // 【關鍵】把資料推給視窗
                dialog.messageChild(JSON.stringify(fetchedData));
                
                // 資料送達後，才停止 Ribbon 上的轉圈圈
                if (currentEvent) {
                    currentEvent.completed();
                    currentEvent = null;
                }
            }
        }, 100); // 每 0.1 秒檢查一次
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