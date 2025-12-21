/* global Office, console */

let dialog;
// 移除 currentEvent 全域變數，因為我們不再需要在視窗關閉時"恢復"原本的發送事件
let cachedPayload = null; 

Office.onReady(() => {
  // Init
});

// 1. 發送攔截 (OnMessageSend)
function validateSend(event) {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        const isVerified = props.get("isVerified");

        if (isVerified === true) {
            // 情況 A: 已經驗證過了，放行
            // (選擇性：如果希望每次都要檢查，可以在這裡 remove("isVerified")，但在你的情境保留比較好)
            props.remove("isVerified"); // 發送後清除，以免下次轉寄時誤判
            props.saveAsync(() => event.completed({ allowEvent: true }));
        } else {
            // 情況 B: 還沒驗證，直接阻擋，並提示使用者
            // 注意：Mac 上不能在這裡呼叫 openDialog，會被系統擋掉
            event.completed({ 
                allowEvent: false, 
                errorMessage: "⚠️ 請點擊工具列上的 [Hello彈窗] 按鈕來確認收件人，確認後再按傳送。" 
            });
        }
    });
}

// 2. 使用者手動點擊 Ribbon 按鈕觸發此函式
function openDialog(event) {
    // 這裡的 event 是按鈕點擊事件，不是發送事件，所以不需要 completed({allowEvent...})
    
    // 準備資料
    fetchData();

    const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html'; // 請確認你的 URL
    
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
            // 處理 Ribbon 按鈕的 event 結束
            if (event) event.completed();
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
    }).catch(e => {
        cachedPayload = { error: "資料讀取失敗: " + e.message };
    });
}

// 3. 處理 Dialog 回傳的訊息
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
        // 使用者在視窗按了「確認發送」
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
             const props = result.value;
             // 設定驗證通過標記
             props.set("isVerified", true);
             props.saveAsync(() => {
                 dialog.close();
                 // 這裡不需要呼叫 event.completed，因為原本的發送請求已經在第一步被我們擋掉了
                 // 使用者現在只需要再次點擊 Outlook 的「傳送」按鈕即可
             });
        });
    } 
    else if (message === "CANCEL") {
        dialog.close();
    }
}

// 註冊全域函式
if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog;