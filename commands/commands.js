/* global Office, console */

let dialog;
let currentEvent;
let fetchedData = null;
// let isDataSent = false;  <-- 移除這個變數，我們不再限制發送次數

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
    fetchedData = null; // 重置資料

    // A. 8秒強制止血 (防止轉圈圈卡死)
    setTimeout(() => {
        if (currentEvent) {
            currentEvent.completed();
            currentEvent = null;
        }
    }, 8000);

    // B. 開啟視窗
    // 請確認這個路徑是您目前部署成功的路徑
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

    // A. 視窗說它打開了 (DIALOG_READY)
    if (message === "DIALOG_READY") {
        // 【關鍵修正】
        // 只要 Dialog 還在喊 READY，就代表它還沒收到資料 (或者漏接了)。
        // 所以我們不檢查 isDataSent，只要資料準備好了就發送！
        
        const waitForData = setInterval(() => {
            if (fetchedData) {
                clearInterval(waitForData);
                
                // 發送資料 (Dialog 收到後會自動停止喊 READY)
                // 這裡可能會發送多次，但沒關係，Dialog 的 renderData 會重繪，確保資料一定會顯示
                dialog.messageChild(JSON.stringify(fetchedData));
                
                // 資料送出後，嘗試停止 Ribbon 轉圈圈
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