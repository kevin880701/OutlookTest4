/* global Office, console */

let dialog;
let currentEvent;
let broadcastTimer;

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

    // B. 開啟視窗
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
                
                // 視窗開啟成功，開始準備資料並廣播
                // 注意：這裡不呼叫 completed，讓轉圈圈繼續轉，視窗才會留著
                startBroadcasting();
            }
        }
    );
}

function startBroadcasting() {
    const item = Office.context.mailbox.item;
    
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.attachments.getAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
        
        // 整理資料 (簡化版，避免太複雜物件導致傳輸失敗)
        const payload = {
            subject: item.subject || "(無主旨)",
            recipients: [
                ...to.map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress, type: 'To' })),
                ...cc.map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress, type: 'Cc' }))
            ],
            attachments: attachments.map(a => ({ name: a.name }))
        };
        
        const jsonString = JSON.stringify(payload);

        // 【關鍵】持續廣播：每 500ms 發送一次
        if (broadcastTimer) clearInterval(broadcastTimer);
        
        broadcastTimer = setInterval(() => {
            if (dialog) {
                try {
                    // 這是 Mac 上唯一能跟 Dialog 溝通的方式
                    dialog.messageChild(jsonString);
                    console.log("Broadcasting...");
                } catch (e) {
                    console.log("Waiting for dialog...");
                }
            }
        }, 500);
    });
}

function processMessage(arg) {
    const message = arg.message;

    if (message === "DATA_RECEIVED") {
        // Dialog 說收到了，停止廣播
        if (broadcastTimer) clearInterval(broadcastTimer);
    }
    else if (message === "VERIFIED_PASS") {
        if (broadcastTimer) clearInterval(broadcastTimer);
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
             const props = result.value;
             props.set("isVerified", true);
             props.saveAsync(() => {
                 dialog.close();
                 // 告訴 Outlook 放行
                 if (currentEvent) {
                     currentEvent.completed({ allowEvent: true });
                     currentEvent = null;
                 }
             });
        });
    } else if (message === "CANCEL") {
        if (broadcastTimer) clearInterval(broadcastTimer);
        dialog.close();
        // 告訴 Outlook 阻擋
        if (currentEvent) {
            currentEvent.completed({ allowEvent: false });
            currentEvent = null;
        }
    }
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;