/* global Office, console */

let dialog;
let currentEvent;
// 狀態標記
let isDialogOpened = false;
let isDataSaved = false;

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
            props.remove("bridge_data"); // 清除暫存
            props.saveAsync(() => event.completed({ allowEvent: true }));
        } else {
            event.completed({ 
                allowEvent: false, 
                errorMessage: "⚠️ 發送中止：請按「不傳送」，然後點擊上方工具列的【LaunchEvent Test】按鈕進行檢查。" 
            });
        }
    });
}

// 2. 開啟視窗 (先開視窗版)
function openDialog(event) {
    currentEvent = event;
    isDialogOpened = false;
    isDataSaved = false;

    // A. 8秒強制止血 (防止轉圈圈卡死)
    setTimeout(() => {
        if (currentEvent) {
            console.log("Timeout: 強制結束轉圈");
            currentEvent.completed();
            currentEvent = null;
        }
    }, 8000);

    // B. 【第一步】立刻開啟視窗 (確保使用者看得到反應)
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html';
    
    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog Failed:", asyncResult.error.message);
                // 開視窗失敗也要試著結束事件
                checkDone();
            } else {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                
                isDialogOpened = true; // 標記：視窗已開啟
                checkDone(); // 檢查是否可以結束轉圈
            }
        }
    );

    // C. 【第二步】在背景抓取並儲存資料
    const item = Office.context.mailbox.item;

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

        // 寫入 CustomProperties
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Failed) return;

            const props = result.value;
            props.set("bridge_data", jsonString);

            props.saveAsync((saveResult) => {
                if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("資料已儲存");
                    isDataSaved = true; // 標記：資料已儲存
                    checkDone(); // 檢查是否可以結束轉圈
                }
            });
        });
    }).catch(e => {
        console.error("資料處理失敗:", e);
        // 就算失敗也要讓轉圈消失
        if (currentEvent) currentEvent.completed();
    });
}

// 檢查是否所有動作都完成了，如果是，就結束轉圈
function checkDone() {
    // 邏輯：視窗開了，或者資料存完了，或者兩者都好了
    // 為了使用者體驗，只要「視窗開了」其實就可以結束轉圈，
    // 因為接下來是視窗自己去讀資料的事
    if (isDialogOpened && currentEvent) {
        currentEvent.completed();
        currentEvent = null;
    }
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