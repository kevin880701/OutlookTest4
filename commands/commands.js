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
            // 已經驗證過了，直接放行 (秒傳)
            props.remove("isVerified");
            props.saveAsync(() => event.completed({ allowEvent: true }));
        } else {
            // 還沒驗證，去開視窗
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

function fetchData() {
    const item = Office.context.mailbox.item;
    
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.getAttachmentsAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
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

// 3. 處理回傳
function processMessage(arg) {
    const message = arg.message;

    if (message === "PULL_DATA") {
        if (cachedPayload) {
            dialog.messageChild(JSON.stringify(cachedPayload));
        } else {
            dialog.messageChild(JSON.stringify({ status: "LOADING" }));
        }
    }
    else if (message === "VERIFIED_PASS") {
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
             const props = result.value;
             props.set("isVerified", true);
             props.saveAsync(() => {
                 // 【關鍵修正】先告訴 Outlook 放行，再關閉視窗
                 // 這樣可以避免視窗關閉後 Context 丟失
                 if (currentEvent) {
                     currentEvent.completed({ allowEvent: true });
                     currentEvent = null;
                 }
                 dialog.close();
             });
        });
    } 
    else if (message === "CANCEL") {
        // 先告訴 Outlook 阻擋
        if (currentEvent) {
            currentEvent.completed({ allowEvent: false });
            currentEvent = null;
        }
        dialog.close();
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;