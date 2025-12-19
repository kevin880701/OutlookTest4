/* global Office */

// ★ 重點 1：把變數宣告在最外面，變成「全域變數」
// 這樣函數執行結束後，它才不會被當作垃圾回收
let dialog; 

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
            // ★ 建議：若還是閃退，將這裡改為 false 試試 (改成獨立視窗通常比較穩定)
            displayInIframe: false 
        },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("彈窗失敗: " + asyncResult.error.message);
                event.completed();
            } else {
                // ★ 重點 2：把開啟的視窗物件存入全域變數
                dialog = asyncResult.value;
                
                // 為了除錯，我們可以加一個事件監聽，確保視窗活著
                dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
                    console.log("視窗事件:", arg);
                });

                // ★ 重點 3：成功開啟後，再告訴 Outlook 結束轉圈圈
                event.completed();
            }
        }
    );
}