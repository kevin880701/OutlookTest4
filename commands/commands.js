/* global Office, console */

let dialog;
let fetchedData = null; // 用來暫存抓到的資料

Office.onReady(() => {
  // Init
});

// 1. 攔截器 (LaunchEvent) - 這部分維持不變
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

// 2. 開啟視窗 (大幅修改：先開視窗，再抓資料)
function openDialog(event) {
    // A. 立刻開啟視窗，避免被 Mac 攔截
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog.html';

    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog Failed:", asyncResult.error.message);
            } else {
                dialog = asyncResult.value;
                // 監聽 Dialog 傳來的 "DIALOG_READY" 或 "VERIFIED_PASS"
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            }

            // B. 視窗指令已送出，立刻停止 Ribbon 轉圈圈
            // 視窗會繼續開啟，背景會繼續抓資料
            if (event) event.completed();
        }
    );

    // C. 視窗開啟的同時，背景開始抓資料
    const item = Office.context.mailbox.item;
    
    // 使用 Promise.all 平行讀取，加速資料準備
    const pTo = new Promise(resolve => item.to.getAsync(r => resolve(r.value || [])));
    const pCc = new Promise(resolve => item.cc.getAsync(r => resolve(r.value || [])));
    const pAtt = new Promise(resolve => item.attachments.getAsync(r => resolve(r.value || [])));

    Promise.all([pTo, pCc, pAtt]).then(([to, cc, attachments]) => {
        // 資料準備好了，存起來等待 Dialog 來要
        fetchedData = {
            subject: item.subject, // Subject 通常是同步的，若讀不到可改用 getAsync
            recipients: [
                ...to.map(r => ({...r, type: 'To'})),
                ...cc.map(r => ({...r, type: 'Cc'}))
            ],
            attachments: attachments
        };
        
        // 如果 Dialog 已經開好並發送過 READY 訊號 (極少見，通常是 Dialog 比較慢)，可以這裡補發
        // 但主要邏輯還是交給 processMessage 處理
        console.log("Data fetched ready.");
    });
}

// 3. 處理訊息 (核心溝通邏輯)
function processMessage(arg) {
    const message = arg.message;

    // A. 視窗載入完畢，請求資料
    if (message === "DIALOG_READY") {
        if (dialog && fetchedData) {
            // 把資料「推」給視窗 (這是 Mac 上唯一可靠的傳輸方式)
            const jsonString = JSON.stringify(fetchedData);
            dialog.messageChild(jsonString);
        } else {
            // 萬一資料還沒抓完，設個簡單的輪詢等待 (或是直接再推一次)
            // 簡單解法：每 500ms 檢查一次 fetchedData
            const checkData = setInterval(() => {
                if (fetchedData) {
                    dialog.messageChild(JSON.stringify(fetchedData));
                    clearInterval(checkData);
                }
            }, 500);
        }
        return;
    }

    // B. 驗證通過
    if (message === "VERIFIED_PASS") {
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            props.saveAsync(() => dialog.close());
        });
    } 
    // C. 取消
    else if (message === "CANCEL") {
        dialog.close();
    }
}

// 綁定
if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;