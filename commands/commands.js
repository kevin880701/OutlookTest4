/* global Office */

Office.onReady(() => {
    // ★★★ 關鍵修正：告訴 Outlook，XML 裡的 "openDialog" 對應到這裡的 openDialog 函數
    Office.actions.associate("openDialog", openDialog);
});

function openDialog(event) {
    // 定義要打開的彈窗網址 (請確認這個檔案存在)
    // 這裡我改回 helloworld.html 作為測試，確認單純彈窗沒問題
    const fullUrl = 'https://icy-moss-034796200.2.azurestaticapps.net/helloworld.html';

    // 打開視窗
    Office.context.ui.displayDialogAsync(
        fullUrl,
        { 
            height: 50, 
            width: 30, 
            displayInIframe: true // Mac 版建議設為 true，顯示效果較好
        },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("彈窗失敗: " + asyncResult.error.message);
            } else {
                // 視窗開啟成功
                const dialog = asyncResult.value;
                
                // 這裡可以加入接收訊息的監聽器 (稍後再加)
                // dialog.addEventHandler(...) 
            }
        }
    );

    // ★★★ 關鍵修正：按鈕點擊事件不需要回傳 { allowEvent: true }
    // 那是給「傳送檢查」用的。普通按鈕只要告訴 Outlook "我做完了" 即可。
    event.completed();
}