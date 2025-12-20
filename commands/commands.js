/* global Office, console */

let dialog;
let currentEvent;
let broadcastTimer; 

Office.onReady(() => {
  // Init
});

function validateSend(event) {
    // 簡單的攔截邏輯
    event.completed({ 
        allowEvent: false, 
        errorMessage: "⚠️ 請點擊工具列按鈕進行檢查。" 
    });
}

// 2. 開啟視窗 (安全版)
function openDialog(event) {
    currentEvent = event;
    
    // 【安全機制】5秒後強制停止轉圈圈，防止程式崩潰導致 Outlook 卡死
    setTimeout(() => {
        if (currentEvent) {
            currentEvent.completed();
            currentEvent = null;
        }
    }, 5000);

    try {
        // 1. 先開啟視窗 (不帶資料，避免 URL 過長崩潰)
        // 請確認路徑是否正確
        const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html'; 
        
        Office.context.ui.displayDialogAsync(
            url, 
            { height: 60, width: 50, displayInIframe: true },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Dialog Open Failed:", asyncResult.error.message);
                    // 如果開視窗失敗，也要記得結束轉圈
                    if (currentEvent) { currentEvent.completed(); currentEvent = null; }
                } else {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                    
                    // 視窗開好了，開始準備資料並廣播
                    startBroadcasting();
                }
            }
        );
    } catch (e) {
        console.error("Critical Error:", e);
        // 發生意外錯誤時，確保轉圈圈停止
        if (currentEvent) { currentEvent.completed(); currentEvent = null; }
    }
}

// 3. 抓資料並廣播
function startBroadcasting() {
    const item = Office.context.mailbox.item;
    
    // 使用 Promise 抓取資料
    Promise.all([
        new Promise(resolve => item.to.getAsync(r => resolve(r.status === 'succeeded' ? r.value : []))),
        new Promise(resolve => item.cc.getAsync(r => resolve(r.status === 'succeeded' ? r.value : []))),
        new Promise(resolve => item.attachments.getAsync(r => resolve(r.status === 'succeeded' ? r.value : [])))
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

        // 【廣播模式】每 500ms 發送一次資料給 Dialog
        // 這樣 Dialog 一打開就會收到，不用依賴 URL
        broadcastTimer = setInterval(() => {
            if (dialog) {
                try {
                    dialog.messageChild(jsonString);
                } catch (e) {
                    // 視窗可能關閉了，忽略錯誤
                }
            }
        }, 500);
    }).catch(e => {
        console.error("Fetch Data Failed:", e);
    });
}

function processMessage(arg) {
    const message = arg.message;

    // 收到 Dialog 說 "DATA_RECEIVED"，可以停止廣播
    if (message === "DATA_RECEIVED") {
        if (broadcastTimer) clearInterval(broadcastTimer);
    }
    else if (message === "VERIFIED_PASS") {
        if (broadcastTimer) clearInterval(broadcastTimer);
        
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            props.saveAsync(() => {
                dialog.close();
                if (currentEvent) { currentEvent.completed(); currentEvent = null; }
            });
        });
    } 
    else if (message === "CANCEL") {
        if (broadcastTimer) clearInterval(broadcastTimer);
        dialog.close();
        if (currentEvent) { currentEvent.completed(); currentEvent = null; }
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;