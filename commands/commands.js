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
            // 驗證過了，清除標記並放行！
            props.remove("isVerified");
            props.remove("bridge_data"); 
            props.saveAsync(() => event.completed({ allowEvent: true }));
        } else {
            // 還沒驗證，去開視窗
            openDialog(event);
        }
    });
}

// 2. 開啟視窗
function openDialog(event) {
    currentEvent = event;

    // A. 60秒安全機制 (延長時間，讓使用者有足夠時間檢查)
    // 如果 60秒內沒反應，強制取消發送，避免 Outlook 卡死
    setTimeout(() => {
        if (currentEvent) {
            console.log("Timeout: 強制結束");
            // 逾時就阻擋，視窗會自動關閉
            currentEvent.completed({ 
                allowEvent: false, 
                errorMessage: "檢查逾時，請重新點擊傳送。" 
            });
            currentEvent = null;
        }
    }, 60000); // 改成 60 秒

    // B. 先開視窗
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html';
    
    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                // 如果連視窗都開不起來，就直接報錯並結束
                console.error("Dialog Failed:", asyncResult.error.message);
                if (currentEvent) {
                    currentEvent.completed({ allowEvent: false, errorMessage: "無法開啟檢查視窗: " + asyncResult.error.message });
                    currentEvent = null;
                }
            } else {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                
                // 【關鍵修改】
                // 視窗開好後，什麼都不要做！不要呼叫 completed！
                // 讓上方的轉圈圈繼續轉，這樣視窗才會留著。
                console.log("視窗已開啟，等待使用者操作...");
            }
        }
    );

    // C. 背景存資料 (維持不變)
    // 視窗開著的同時，我們在背後默默把資料準備好，dialog.js 會自己來抓
    const item = Office.context.mailbox.item;
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.attachments.getAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
        const payload = {
            subject: item.subject || "(無主旨)",
            recipients: [
                ...to.map(r => ({...r, type: 'To'})),
                ...cc.map(r => ({...r, type: 'Cc'}))
            ],
            attachments: attachments
        };
        const jsonString = JSON.stringify(payload);

        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const props = result.value;
                props.set("bridge_data", jsonString);
                props.saveAsync(() => console.log("資料已儲存至 Bridge"));
            }
        });
    });
}

// 3. 處理回傳 (當使用者在彈窗按了按鈕)
function processMessage(arg) {
    const message = arg.message;

    if (message === "VERIFIED_PASS") {
        // 使用者按了「確認發送」
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
             const props = result.value;
             props.set("isVerified", true); // 寫入驗證通過標記
             props.saveAsync(() => {
                 dialog.close();
                 // 這裡告訴 Outlook：驗證通過，請放行！(信件會直接寄出)
                 if (currentEvent) {
                     currentEvent.completed({ allowEvent: true });
                     currentEvent = null;
                 }
             });
        });
    } else if (message === "CANCEL") {
        // 使用者按了「取消」
        dialog.close();
        // 這裡告訴 Outlook：使用者後悔了，阻擋發送！
        if (currentEvent) {
            currentEvent.completed({ allowEvent: false });
            currentEvent = null;
        }
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;