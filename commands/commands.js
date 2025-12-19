/* global Office */

Office.onReady(() => {
    // 註冊函數
    Office.actions.associate("openDialog", openDialog);
});

function openDialog(event) {
    // 確保這裡的 URL 是正確的
    const fullUrl = 'https://icy-moss-034796200.2.azurestaticapps.net/helloworld.html';

    Office.context.ui.displayDialogAsync(
        fullUrl,
        { 
            height: 50, 
            width: 30, 
            // ★ 修正 1：Mac 上務必設為 false，讓它變成獨立視窗
            displayInIframe: false 
        },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("彈窗失敗: " + asyncResult.error.message);
                event.completed(); // 失敗的話就直接結束
            } else {
                // 視窗成功開啟指令已送出...
                const dialog = asyncResult.value;
                
                // 可以在這裡加監聽器 (可選)
                 dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                     dialog.close();
                 });

                // ★ 修正 2：不要立刻自殺！使用 setTimeout 拖延 5 秒
                // 這給了彈窗足夠的時間完成載入並脫離父程序
                setTimeout(() => {
                    console.log("背景程式任務完成，正在結束...");
                    event.completed();
                }, 5000); // 延遲 5000 毫秒 (5秒)
            }
        }
    );
}