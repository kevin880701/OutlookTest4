/* global Office, console */

let dialog; // 用來存放彈窗物件

Office.onReady(() => {
  // 初始化完成
});

// --------------------------------------------------------
// 1. LaunchEvent: 守門員 (OnMessageSend)
// --------------------------------------------------------
function validateSend(event) {
    // 讀取這封信的 Custom Properties (隱藏標籤)
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        const isVerified = props.get("isVerified"); // 檢查有沒有通行證

        if (isVerified === true) {
            // A. 通行證存在：允許寄出
            // (選擇性) 寄出後清除標記，確保下次編輯草稿時需要重新檢查
            props.remove("isVerified");
            props.saveAsync(() => {
                event.completed({ allowEvent: true });
            });
        } else {
            // B. 沒有通行證：阻擋寄出
            event.completed({ 
                allowEvent: false, 
                errorMessage: "⚠️ 發送中止：請先點擊上方工具列的【Hello彈窗】按鈕，確認收件人與附件後才可發送。" 
            });
        }
    });
}

// --------------------------------------------------------
// 2. Ribbon Action: 開啟檢查視窗 (openDialog)
// --------------------------------------------------------
function openDialog(event) {
    // 開啟 dialog.html
    Office.context.ui.displayDialogAsync(
        'https://icy-moss-034796200.2.azurestaticapps.net/dialog.html', // 請確認路徑正確
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
            } else {
                dialog = asyncResult.value;
                // 監聽來自 dialog 的訊息
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            }
        }
    );
    
    // 如果是從 Ribbon 按鈕觸發，需要告訴 Outlook 執行結束
    if (event) event.completed();
}

// --------------------------------------------------------
// 3. 處理 Dialog 回傳的訊息
// --------------------------------------------------------
function processMessage(arg) {
    const message = arg.message;

    if (message === "VERIFIED_PASS") {
        // A. 使用者在視窗按了「確認發送」
        // 寫入「通行證」到 Custom Properties
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); // 發放通行證
            
            // 重要：一定要 saveAsync 才會寫入伺服器/快取
            props.saveAsync((saveResult) => {
                if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                    dialog.close();
                    // 提示使用者現在可以按傳送了 (或者使用 Office.context.mailbox.item.close() 雖然不能直接觸發傳送)
                    // 在 Mac 上通常使用者需要手動再按一次傳送
                } else {
                    console.error("存檔失敗");
                }
            });
        });
    } else if (message === "CANCEL") {
        // B. 使用者取消
        dialog.close();
    }
}

// 綁定全域函式
if (typeof g === 'undefined') {
    var g = window;
}
g.validateSend = validateSend;
g.openDialog = openDialog;