/* global Office, console */

let dialog;
let currentEvent;
let isDialogOpened = false;

Office.onReady(() => {
  // Init
});

// 1. 攔截器
function validateSend(event) {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        const isVerified = props.get("isVerified");

        if (isVerified === true) {
            // 第二次進來：驗證過了，放行！
            props.remove("isVerified");
            props.remove("bridge_data"); 
            props.saveAsync(() => event.completed({ allowEvent: true }));
        } else {
            // 第一次進來：還沒驗證，去開視窗
            openDialog(event);
        }
    });
}

// 2. 開啟視窗
function openDialog(event) {
    currentEvent = event;
    isDialogOpened = false;

    // A. 8秒安全機制
    setTimeout(() => {
        if (currentEvent) {
            console.log("Timeout: 強制結束");
            // 【修正點 1】超時也要阻擋發送，不然沒檢查就寄出了
            currentEvent.completed({ allowEvent: false, errorMessage: "系統回應逾時，請再試一次。" });
            currentEvent = null;
        }
    }, 8000);

    // B. 先開視窗
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html';
    
    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog Failed:", asyncResult.error.message);
                if (currentEvent) {
                    currentEvent.completed({ allowEvent: false, errorMessage: "無法開啟檢查視窗: " + asyncResult.error.message });
                    currentEvent = null;
                }
            } else {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                
                isDialogOpened = true; 
                checkDone(); // 視窗開好了，去通知 Outlook
            }
        }
    );

    // C. 背景存資料 (維持不變)
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
                props.saveAsync(() => console.log("資料已儲存"));
            }
        });
    });
}

// 檢查是否完成
function checkDone() {
    if (isDialogOpened && currentEvent) {
        // 【修正點 2 - 最關鍵的一行】
        // 必須設定 allowEvent: false，告訴 Outlook「先別寄！先停下來！」
        // 這樣視窗才會留著，不會被關掉。
        currentEvent.completed({ 
            allowEvent: false, 
            errorMessage: "請在彈出的視窗中確認收件人與附件，確認無誤後請再次點擊「傳送」。" 
        });
        currentEvent = null;
    }
}

// 3. 處理回傳
function processMessage(arg) {
    const message = arg.message;

    if (message === "VERIFIED_PASS") {
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
             const props = result.value;
             props.set("isVerified", true); // 寫入驗證通過標記
             props.saveAsync(() => {
                 dialog.close();
                 // 使用者現在可以再按一次 Outlook 的「傳送」按鈕了
             });
        });
    } else if (message === "CANCEL") {
        dialog.close();
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;