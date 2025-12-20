/* global Office, console */

let currentEvent;

Office.onReady(() => {
  // Init
});

// 1. 攔截器 (維持不變)
function validateSend(event) {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        const isVerified = props.get("isVerified");

        if (isVerified === true) {
            props.remove("isVerified");
            // 順便清除暫存資料，保持乾淨
            props.remove("temp_data"); 
            props.saveAsync(() => event.completed({ allowEvent: true }));
        } else {
            event.completed({ 
                allowEvent: false, 
                errorMessage: "⚠️ 發送中止：請按「不傳送」，然後點擊上方工具列的【LaunchEvent Test】按鈕進行檢查。" 
            });
        }
    });
}

// 2. 開啟視窗 (資料橋接版)
function openDialog(event) {
    currentEvent = event;

    // A. 5秒強制止血 (防止轉圈圈卡死)
    setTimeout(() => {
        if (currentEvent) {
            console.log("Timeout: 強制結束轉圈");
            currentEvent.completed();
            currentEvent = null;
        }
    }, 5000);

    const item = Office.context.mailbox.item;

    // B. 抓取資料
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.attachments.getAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
        
        const payload = {
            subject: item.subject,
            recipients: [
                ...to.map(r => ({...r, type: 'To'})),
                ...cc.map(r => ({...r, type: 'Cc'}))
            ],
            attachments: attachments
        };
        
        const jsonString = JSON.stringify(payload);

        // C. 【關鍵】把資料寫入 CustomProperties (埋時空膠囊)
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error("無法讀取屬性");
                return;
            }

            const props = result.value;
            // 設定暫存資料 key: "temp_data"
            props.set("temp_data", jsonString);

            // D. 存檔成功後，才打開視窗 (確保視窗讀得到)
            props.saveAsync((saveResult) => {
                if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("資料已寫入橋接器，準備開視窗...");
                    launchDialog();
                } else {
                    console.error("存檔失敗");
                }
            });
        });
    });
}

function launchDialog() {
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html';
    
    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog Failed:", asyncResult.error.message);
            } else {
                const dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            }
            
            // 視窗只要一開，我們就可以結束轉圈圈了
            // 因為資料已經存在信件裡，視窗自己會去讀
            if (currentEvent) {
                currentEvent.completed();
                currentEvent = null;
            }
        }
    );
}

// 3. 處理回傳
function processMessage(arg) {
    const message = arg.message;
    // 這裡只需要處理關閉視窗的邏輯
    // 資料讀取都在 dialog.js 內部完成
    if (message === "CLOSE_DIALOG") {
        // 驗證通過的邏輯在 dialog.js 寫入 isVerified 後觸發
        // 這裡只要簡單關閉即可
        // 但為了保險，我們還是重整一下 isVerified
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
             const props = result.value;
             props.set("isVerified", true);
             props.saveAsync(() => {
                 if (currentEvent) currentEvent.completed(); // 雙重保險
             });
        });
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;