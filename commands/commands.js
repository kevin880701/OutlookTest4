/* global Office, console */

let dialog;
let currentEvent;
let cachedPayload = null; 

Office.onReady(() => {
  // Init
});

function validateSend(event) {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        const isVerified = props.get("isVerified");

        if (isVerified === true) {
            props.remove("isVerified");
            props.saveAsync(() => event.completed({ allowEvent: true }));
        } else {
            openDialog(event);
        }
    });
}

function openDialog(event) {
    currentEvent = event;
    cachedPayload = null; 

    // A. 60秒安全機制
    setTimeout(() => {
        if (currentEvent) {
            currentEvent.completed({ 
                allowEvent: false, 
                errorMessage: "檢查逾時，請重新點擊傳送。" 
            });
            currentEvent = null;
        }
    }, 60000);

    // B. 馬上開始抓資料
    fetchData();

    // C. 開啟視窗
    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html';
    
    Office.context.ui.displayDialogAsync(
        url, 
        { height: 60, width: 50, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog Failed:", asyncResult.error.message);
                if (currentEvent) {
                    currentEvent.completed({ allowEvent: false, errorMessage: "視窗開啟失敗" });
                    currentEvent = null;
                }
            } else {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                console.log("視窗已開，等待 PULL_DATA 請求...");
            }
        }
    );
}

// 抓取資料函式
function fetchData() {
    const item = Office.context.mailbox.item;
    
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        // 【修正點】Compose 模式下，讀取附件要用 getAttachmentsAsync
        new Promise(r => item.getAttachmentsAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
        // 資料抓好了，存起來
        cachedPayload = {
            subject: item.subject || "(無主旨)",
            recipients: [
                ...to.map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress, type: 'To' })),
                ...cc.map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress, type: 'Cc' }))
            ],
            attachments: attachments.map(a => ({ name: a.name }))
        };
        console.log("資料準備完畢 (Cached)");
    }).catch(e => {
        console.error("Fetch error:", e);
        cachedPayload = { error: "資料讀取失敗: " + e.message };
    });
}

// 處理 Dialog 傳來的訊息
function processMessage(arg) {
    const message = arg.message;

    // A. Dialog 主動來要資料 (Pull)
    if (message === "PULL_DATA") {
        if (cachedPayload) {
            // 資料已經好了，發送！
            dialog.messageChild(JSON.stringify(cachedPayload));
        } else {
            // 資料還沒好，跟他說還在 Loading
            dialog.messageChild(JSON.stringify({ status: "LOADING" }));
        }
    }
    // B. 驗證通過
    else if (message === "VERIFIED_PASS") {
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
             const props = result.value;
             props.set("isVerified", true);
             props.saveAsync(() => {
                 dialog.close();
                 if (currentEvent) {
                     currentEvent.completed({ allowEvent: true });
                     currentEvent = null;
                 }
             });
        });
    } 
    // C. 取消
    else if (message === "CANCEL") {
        dialog.close();
        if (currentEvent) {
            currentEvent.completed({ allowEvent: false });
            currentEvent = null;
        }
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;