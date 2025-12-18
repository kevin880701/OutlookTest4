/* global Office */

let sendEvent; // 用來儲存發送事件

Office.onReady();

// 這是 manifest.xml 裡指定的 FunctionName
function validateSend(event) {
    sendEvent = event; // 把事件存起來，稍後決定要不要放行

    // 打開彈跳視窗
    const dialogUrl = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html';

    Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 50, width: 30, displayInIframe: true },
        dialogCallback
    );
}

function dialogCallback(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // 如果視窗開失敗，直接允許發送 (避免卡死)
        sendEvent.completed({ allowEvent: true });
    } else {
        // 取得對話框物件
        const dialog = asyncResult.value;
        // 監聽對話框傳回來的訊息
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            processMessage(arg, dialog);
        });
    }
}

function processMessage(arg, dialog) {
    dialog.close(); // 收到訊息後先關閉視窗

    if (arg.message === "SEND_MAIL") {
        // 使用者按下確認，允許發送
        sendEvent.completed({ allowEvent: true });
    } else {
        // 使用者按下取消，阻止發送
        sendEvent.completed({ allowEvent: false });
    }
}