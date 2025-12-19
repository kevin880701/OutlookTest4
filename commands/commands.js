/* global Office */

let sendEvent; // 用來儲存發送事件的變數

Office.onReady(() => {
    // 註冊攔截發送的函數
    Office.actions.associate("validateSend", validateSend);
});

function validateSend(event) {
    sendEvent = event; // 把事件存起來，稍後要決定是放行還是擋下

    // 定義彈窗網址
    const fullUrl = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog.html';

    // 打開視窗
    Office.context.ui.displayDialogAsync(
        fullUrl,
        { 
            height: 40, 
            width: 30, 
            displayInIframe: false // Mac 必須維持 false
        },
        dialogCallback // 視窗打開後，執行這個回呼函數
    );
}

function dialogCallback(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // 如果視窗根本打不開 (例如被瀏覽器擋住)，為了安全起見，通常選擇「擋下」或「放行」
        console.error("無法開啟彈窗: " + asyncResult.error.message);
        // 這裡我們選擇放行，避免使用者永遠寄不出去 (或者你可以改 false 擋下)
        sendEvent.completed({ allowEvent: true });
    } else {
        // 視窗成功打開，我們開始監聽來自視窗的訊息
        const dialog = asyncResult.value;
        
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            // 收到訊息了！
            const message = arg.message;

            // 關閉視窗
            dialog.close();

            // 判斷按了什麼按鈕
            if (message === "SEND") {
                // 使用者按了發送 -> 告訴 Outlook 放行
                sendEvent.completed({ allowEvent: true });
            } else {
                // 使用者按了取消 -> 告訴 Outlook 擋下
                sendEvent.completed({ allowEvent: false });
            }
        });
    }
}