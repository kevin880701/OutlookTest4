/* global Office, console */

let dialog; // 用來存放彈窗物件

Office.onReady(() => {
  // 初始化完成
});

// --------------------------------------------------------
// 1. LaunchEvent: 守門員 (OnMessageSend)
// --------------------------------------------------------
function validateSend(event) {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        const isVerified = props.get("isVerified");

        if (isVerified === true) {
            props.remove("isVerified");
            props.saveAsync(() => {
                event.completed({ allowEvent: true });
            });
        } else {
            event.completed({ 
                allowEvent: false, 
                errorMessage: "⚠️ 發送中止：請按「不傳送」，然後點擊上方工具列的【LaunchEvent Test】按鈕進行檢查。" 
            });
        }
    });
}

// --------------------------------------------------------
// 2. Ribbon Action: 開啟檢查視窗 (openDialog)
// --------------------------------------------------------
function openDialog(event) {
    const item = Office.context.mailbox.item;

    // 使用巢狀 Callback 確保資料都讀取完畢
    item.to.getAsync((resTo) => {
        item.cc.getAsync((resCc) => {
            item.attachments.getAsync((resAtt) => {
                
                // 1. 整理資料
                const payload = {
                    subject: item.subject, 
                    recipients: [
                        ...(resTo.value ? resTo.value.map(r => ({...r, type: 'To'})) : []),
                        ...(resCc.value ? resCc.value.map(r => ({...r, type: 'Cc'})) : [])
                    ],
                    attachments: resAtt.value || []
                };

                // 2. 轉成字串並編碼
                const jsonString = encodeURIComponent(JSON.stringify(payload));
                
                // 3. 開啟視窗
                // 注意：URL 後面帶上參數
                const url = `https://icy-moss-034796200.2.azurestaticapps.net/dialog.html?data=${jsonString}`;

                Office.context.ui.displayDialogAsync(
                    url, 
                    { height: 60, width: 50, displayInIframe: true },
                    (asyncResult) => {
                        // 【關鍵修正】
                        // 必須等到 Dialog 嘗試開啟的 callback 回來後，才告訴 Outlook 結束事件
                        
                        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                            console.error(asyncResult.error.message);
                        } else {
                            dialog = asyncResult.value;
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                        }

                        // ---> 這行搬到這裡來了！ <---
                        // 只有當上面的 displayDialogAsync 執行之後，才通知 Outlook 按鈕動作結束
                        if (event) event.completed();
                    }
                );
            });
        });
    });
    
    // 原本這裡的 event.completed() 刪除，因為它跑太快了
}

// --------------------------------------------------------
// 3. 處理 Dialog 回傳的訊息
// --------------------------------------------------------
function processMessage(arg) {
    const message = arg.message;

    if (message === "VERIFIED_PASS") {
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            
            props.saveAsync((saveResult) => {
                if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                    dialog.close();
                } else {
                    console.error("存檔失敗");
                }
            });
        });
    } else if (message === "CANCEL") {
        dialog.close();
    }
}

// 綁定全域函式
if (typeof g === 'undefined') {
    var g = window;
}
g.validateSend = validateSend;
g.openDialog = openDialog;