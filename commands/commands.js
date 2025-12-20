/* global Office, console */

let dialog;
let currentEvent;
let broadcastTimer; // 廣播計時器

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

// 2. 開啟視窗 (廣播模式)
function openDialog(event) {
    currentEvent = event;
    
    // 【關鍵修正 1】3秒強制止血機制
    // 無論發生什麼事，3秒後強制告訴 Outlook "我做完了"，讓轉圈圈消失
    setTimeout(() => {
        if (currentEvent) {
            console.log("強制結束轉圈圈");
            currentEvent.completed(); 
            currentEvent = null;
        }
    }, 3000);

    // 【關鍵修正 2】使用乾淨的網址 (不帶 ?data=，避免崩潰)
    // 請確認路徑是否正確
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html'; 
    
    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog Open Failed:", asyncResult.error.message);
            } else {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                
                // 視窗成功開啟後，開始準備資料並廣播
                startBroadcasting();
            }
        }
    );
}

// 核心：抓資料並持續廣播
function startBroadcasting() {
    const item = Office.context.mailbox.item;
    
    // 平行抓取資料
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

        // 【關鍵修正 3】啟動廣播：每 500ms 發送一次資料
        // 不管 Dialog 有沒有回應，一直送就對了
        if (broadcastTimer) clearInterval(broadcastTimer);
        
        broadcastTimer = setInterval(() => {
            if (dialog) {
                try {
                    // 這行指令在 Mac 上最穩定，由母視窗主動推給子視窗
                    dialog.messageChild(jsonString);
                    console.log("Broadcasting data...");
                } catch (e) {
                    console.log("Broadcast waiting...");
                }
            }
        }, 500);
    });
}

// 3. 處理訊息
function processMessage(arg) {
    const message = arg.message;

    // A. 視窗說 "DATA_RECEIVED" -> 停止廣播
    if (message === "DATA_RECEIVED") {
        if (broadcastTimer) clearInterval(broadcastTimer);
    }
    // B. 驗證通過
    else if (message === "VERIFIED_PASS") {
        if (broadcastTimer) clearInterval(broadcastTimer);
        
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            props.saveAsync(() => dialog.close());
        });
    } 
    // C. 取消
    else if (message === "CANCEL") {
        if (broadcastTimer) clearInterval(broadcastTimer);
        dialog.close();
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;