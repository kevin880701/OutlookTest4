/* global Office, console */

let dialog;
let currentEvent;

// 務必註冊函數，不然 Outlook 找不到
Office.actions.associate("validateSend", validateSend);
Office.actions.associate("openDialog", openDialog);

Office.onReady();

// 1. 攔截器 (OnSend)
function validateSend(event) {
    currentEvent = event;
    // 直接呼叫 openDialog 邏輯，重用程式碼
    openDialogLogic(event, true);
}

// 2. 按鈕點擊 (Button)
function openDialog(event) {
    currentEvent = event;
    openDialogLogic(event, false);
}

// 核心邏輯
function openDialogLogic(event, isOnSend) {
    // 1. 先抓取資料
    const item = Office.context.mailbox.item;
    
    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.attachments.getAsync(x => r(x.value || []))),
        new Promise(r => item.subject.getAsync(x => r(x.value || "")))
    ]).then(([to, cc, attachments, subject]) => {
        
        // 2. 打包資料
        const dataPackage = {
            subject: subject,
            recipients: [
                ...to.map(r => ({...r, type: 'To'})),
                ...cc.map(r => ({...r, type: 'Cc'}))
            ],
            attachments: attachments
        };

        // 3. 將資料轉為字串並編碼 (放在 URL 裡)
        // 注意：URL 長度有限制，如果附件太多可能會爆，但一般使用足夠
        const jsonString = encodeURIComponent(JSON.stringify(dataPackage));
        const url = `https://icy-moss-034796200.2.azurestaticapps.net/dialog/dialog.html?data=${jsonString}`;

        // 4. 打開視窗 (帶著資料)
        Office.context.ui.displayDialogAsync(
            url, 
            { height: 60, width: 50, displayInIframe: true },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Dialog Failed:", asyncResult.error.message);
                    // 如果開窗失敗且是攔截發送模式，為了避免卡信，通常選擇放行或報錯
                    if (isOnSend) event.completed({ allowEvent: false, errorMessage: "無法開啟檢查視窗" });
                    else event.completed();
                } else {
                    dialog = asyncResult.value;
                    // 監聽來自彈窗的回報 (只聽結果，不傳資料了)
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                }
            }
        );
    }).catch(err => {
        console.error("Fetch Error:", err);
        if (isOnSend) event.completed({ allowEvent: true }); // 出錯就放行，避免卡死
        else event.completed();
    });
}

// 3. 處理訊息 (只負責接收結果)
function processMessage(arg) {
    const message = arg.message;

    if (message === "VERIFIED_PASS") {
        // 使用者按了「確認發送」
        // 寫入標記 (如果是 Smart Alerts 流程需要)
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            props.set("isVerified", true); 
            props.saveAsync(() => {
                dialog.close();
                // 告訴 Outlook 任務完成，允許發送
                if (currentEvent) currentEvent.completed({ allowEvent: true });
            });
        });
    } 
    else if (message === "CANCEL") {
        dialog.close();
        // 告訴 Outlook 任務完成，但禁止發送 (如果是 OnSend)
        if (currentEvent) currentEvent.completed({ allowEvent: false });
    }
}