/* * 這是您的 commands.js 
 * 必須確保 Office.initialize 已宣告
 */

Office.onReady(() => {
  // 初始化
});

function validateSend(event) {
    // 這裡可以寫邏輯，例如檢查主旨有沒有寫
    // 為了測試，我們直接無條件攔截

    console.log("LaunchEvent 觸發了！");

    // 關鍵指令：告訴 Outlook 處理完畢
    // allowEvent: false 代表「禁止寄出」
    // errorMessage: 會顯示在那條 Banner 上提示使用者
    event.completed({ 
        allowEvent: false, 
        errorMessage: "測試成功！這封信被 LaunchEvent 強制攔截了，您無法寄出。" 
    });
}

// 重要：必須將函式綁定到全域變數，Manifest 才呼叫得到
// 如果您用 Webpack，寫法可能不同，若是原生 JS 這樣寫即可：
if (typeof g === 'undefined') {
    var g = window; // 或 global
}
g.validateSend = validateSend;