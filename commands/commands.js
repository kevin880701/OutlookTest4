/* global Office */

Office.onReady(() => {
    Office.actions.associate("openDialog", openDialog);
});

function openDialog(event) {
    const fullUrl = 'https://icy-moss-034796200.2.azurestaticapps.net/helloworld.html';

    Office.context.ui.displayDialogAsync(
        fullUrl,
        { 
            height: 50, 
            width: 30, 
            displayInIframe: true 
        },
        function (asyncResult) {
            // ★ 重點修正：這行原本在外面，現在搬進來這裡了 ★
            // 只有當 Outlook 回報「視窗狀態」後，我們才告訴它「動作結束」
            
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("彈窗失敗: " + asyncResult.error.message);
            }
            
            // 告訴 Outlook 動作結束，可以停止轉圈圈了
            // 因為現在視窗已經建立了，這時候結束背景程式通常是安全的
            event.completed();
        }
    );
    
    // 原本這裡的 event.completed() 刪掉
}