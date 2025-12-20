/* global Office, console */

let dialog;
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
            props.remove("bridge_data"); 
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

    // 【保險機制】5秒後強制停止轉圈圈 (防止程式崩潰導致 Outlook 卡死)
    setTimeout(() => {
        if (currentEvent) {
            console.log("Timeout: 強制結束轉圈");
            currentEvent.completed();
            currentEvent = null;
        }
    }, 5000);

    const item = Office.context.mailbox.item;

    // A. 抓取資料
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.attachments.getAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
        
        const payload = {
            subject: item.subject || "(無主旨)",
            recipients: [
                ...to.map(r => ({...r, type: 'To'})),
                ...cc.map(r => ({...r, type: 'Cc'}))
            ],
            attachments: attachments
        };
        
        const jsonString = JSON.stringify(payload);

        // B. 【關鍵】把資料寫入 CustomProperties (埋入橋接資料)
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error("無法讀取屬性");
                // 就算失敗也要開視窗，不然會沒反應
                launchDialog(); 
                return;
            }

            const props = result.value;
            // 設定暫存資料 key: "bridge_data"
            props.set("bridge_data", jsonString);

            // C. 存檔成功後，才打開視窗 (確保視窗讀得到)
            props.saveAsync((saveResult) => {
                if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("資料已寫入橋接器，準備開視窗...");
                    launchDialog();
                } else {
                    console.error("存檔失敗");
                    launchDialog(); // 失敗還是要開視窗
                }
            });
        });
    }).catch(e => {
        console.error("資料抓取失敗:", e);
        if (currentEvent) currentEvent.completed();
    });
}

// 輔助：開啟視窗
function launchDialog() {
    // 使用乾淨的網址，不帶參數，絕對不會崩潰
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html';
    
    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog Failed:", asyncResult.error.message);
            } else {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            }
            
            // 【關鍵】視窗指令送出後，立刻結束轉圈圈
            // 因為資料已經在信件裡了，視窗自己會去讀，不用我們管
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

    if (message === "VERIFIED_PASS") {
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
             const props = result.value;
             props.set("isVerified", true);
             props.saveAsync(() => {
                 dialog.close();
             });
        });
    } else if (message === "CANCEL") {
        dialog.close();
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;