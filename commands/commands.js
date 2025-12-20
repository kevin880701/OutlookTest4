/* global Office, console, localStorage */

let dialog;

Office.onReady(() => {
  // Init
});

// 1. 攔截器 (這部分沒變)
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

// 2. 開啟視窗 (修改重點：改用 LocalStorage)
function openDialog(event) {
    const item = Office.context.mailbox.item;

    // 讀取資料
    item.to.getAsync((resTo) => {
        item.cc.getAsync((resCc) => {
            item.attachments.getAsync((resAtt) => {
                
                // 整理資料
                const payload = {
                    subject: item.subject, 
                    recipients: [
                        ...(resTo.value ? resTo.value.map(r => ({...r, type: 'To'})) : []),
                        ...(resCc.value ? resCc.value.map(r => ({...r, type: 'Cc'})) : [])
                    ],
                    attachments: resAtt.value || []
                };

                // 【修正 1】將資料存入 LocalStorage，而不是塞在網址裡
                // 這樣可以避免網址過長導致的錯誤
                try {
                    localStorage.setItem("outlook_verify_data", JSON.stringify(payload));
                } catch (e) {
                    console.error("Storage Error:", e);
                }
                
                // 【修正 2】網址變乾淨了，不帶參數
                const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog.html';

                Office.context.ui.displayDialogAsync(
                    url, 
                    { height: 60, width: 50, displayInIframe: true },
                    (asyncResult) => {
                        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                            console.error(asyncResult.error.message);
                        } else {
                            dialog = asyncResult.value;
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                        }

                        // 【修正 3】無論成功失敗，都要告訴 Outlook 結束轉圈圈
                        if (event) event.completed();
                    }
                );
            });
        });
    });
}

// 3. 處理回傳 (這部分沒變)
function processMessage(arg) {
    const message = arg.message;
    if (message === "VERIFIED_PASS") {
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            props.saveAsync(() => dialog.close());
        });
    } else if (message === "CANCEL") {
        dialog.close();
    }
}

// 綁定
if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;