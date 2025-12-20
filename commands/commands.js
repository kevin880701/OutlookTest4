/* global Office, console */

let dialog;
let currentEvent;
let pushInterval; // 用來持續廣播資料的定時器
let fetchedData = null; // 暫存資料

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

// 2. 開啟視窗 & 啟動廣播
function openDialog(event) {
    currentEvent = event;
    fetchedData = null; 

    // A. 8秒強制止血 (防止轉圈圈卡死)
    setTimeout(() => {
        stopBroadcasting(); // 時間到，停止廣播
        if (currentEvent) {
            currentEvent.completed();
            currentEvent = null;
        }
    }, 8000);

    // B. 開啟視窗
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
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                
                // 視窗開好了，嘗試開始廣播 (如果資料也已經準備好的話)
                startBroadcasting();
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
        // 資料準備好了，嘗試開始廣播 (如果視窗也已經開好的話)
        startBroadcasting();
    });
}

// 核心：持續廣播資料給 Dialog
function startBroadcasting() {
    // 必須同時滿足：1. Dialog物件存在 2. 資料已抓到
    if (!dialog || !fetchedData) return;

    // 清除舊的廣播 (避免重複)
    if (pushInterval) clearInterval(pushInterval);

    console.log("開始廣播資料給 Dialog...");
    
    // 每 500ms 發送一次，確保 Dialog 一打開就能收到
    pushInterval = setInterval(() => {
        try {
            // 使用 messageChild 主動推播
            dialog.messageChild(JSON.stringify(fetchedData));
        } catch (e) {
            console.log("廣播失敗 (視窗可能已關閉):", e);
            stopBroadcasting();
        }
    }, 500);
}

function stopBroadcasting() {
    if (pushInterval) {
        clearInterval(pushInterval);
        pushInterval = null;
    }
}

// 3. 處理訊息 (只處理使用者操作)
function processMessage(arg) {
    const message = arg.message;

    // A. 收到 Dialog 回傳 "DATA_RECEIVED" (代表它收到了，我們可以停止廣播)
    if (message === "DATA_RECEIVED") {
        stopBroadcasting(); // 任務達成，停止吵鬧
    }
    // B. 驗證通過
    else if (message === "VERIFIED_PASS") {
        stopBroadcasting();
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            props.saveAsync(() => {
                dialog.close();
                // 記得這裡也要結束轉圈圈
                if (currentEvent) { currentEvent.completed(); currentEvent = null; }
            });
        });
    } 
    // C. 取消
    else if (message === "CANCEL") {
        stopBroadcasting();
        dialog.close();
        if (currentEvent) { currentEvent.completed(); currentEvent = null; }
    }
}

// 綁定
if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;