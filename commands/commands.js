/* global Office, console */

let dialog;
let currentEvent;
let broadcastTimer; // 廣播計時器

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

// 2. 開啟視窗 (廣播模式)
function openDialog(event) {
    currentEvent = event;
    
    // A. 3秒後強制結束 LaunchEvent 轉圈圈 (這是解決 Loading 的關鍵保險)
    // 就算資料還沒傳完，也要先讓 Outlook 知道我們活著，避免被系統殺掉
    setTimeout(() => {
        if (currentEvent) {
            currentEvent.completed(); 
            currentEvent = null;
        }
    }, 3000);

    // B. 開啟視窗
    // 請確認路徑是否正確，您之前的截圖顯示是 /dialog/dialog.html
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html'; 
    
    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
            } else {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                
                // 視窗開啟成功，開始準備資料並廣播
                startBroadcasting();
            }
        }
    );
}

// 核心：抓資料並持續廣播
function startBroadcasting() {
    const item = Office.context.mailbox.item;
    
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.attachments.getAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
        
        const payload = {
            subject: item.subject,
            recipients: [
                ...to.map(r => ({...r, type: 'To'})),
                ...cc.map(r => ({...r, type: 'Cc'}))
            ],
            attachments: attachments
        };
        
        const jsonString = JSON.stringify(payload);

        // 【關鍵】啟動廣播：每 500ms 發送一次資料
        // 不管 Dialog 有沒有回應，一直送就對了
        broadcastTimer = setInterval(() => {
            if (dialog) {
                try {
                    dialog.messageChild(jsonString);
                    console.log("Broadcasting data...");
                } catch (e) {
                    console.error("Broadcast failed:", e);
                }
            }
        }, 500);
    });
}

// 3. 處理訊息 (只處理關閉與驗證)
function processMessage(arg) {
    const message = arg.message;

    // 如果視窗收到資料並回傳 ACK，我們可以停止廣播 (選擇性)
    if (message === "DATA_RECEIVED") {
        if (broadcastTimer) clearInterval(broadcastTimer);
    }
    else if (message === "VERIFIED_PASS") {
        if (broadcastTimer) clearInterval(broadcastTimer);
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            props.saveAsync(() => dialog.close());
        });
    } 
    else if (message === "CANCEL") {
        if (broadcastTimer) clearInterval(broadcastTimer);
        dialog.close();
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;