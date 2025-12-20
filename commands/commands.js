/* global Office, console */

let dialog;
let fetchedData = null; // 用來暫存抓到的資料

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
            props.saveAsync(() => event.completed({ allowEvent: true }));
        } else {
            event.completed({ 
                allowEvent: false, 
                errorMessage: "⚠️ 發送中止：請按「不傳送」，然後點擊上方工具列的【LaunchEvent Test】按鈕進行檢查。" 
            });
        }
    });
}

// 2. 開啟視窗 (改用 Promise 與 訊息傳遞)
async function openDialog(event) {
    try {
        // A. 先開啟視窗 (讓使用者立刻看到反應)
        const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog.html';
        
        // 使用 Promise 包裝 displayDialogAsync
        await new Promise((resolve, reject) => {
            Office.context.ui.displayDialogAsync(
                url, 
                { height: 60, width: 50, displayInIframe: true },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(new Error(asyncResult.error.message));
                    } else {
                        dialog = asyncResult.value;
                        // 監聽 Dialog 傳來的訊息 ("READY", "VERIFIED_PASS"...)
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                        resolve();
                    }
                }
            );
        });

        // B. 視窗開好後，開始抓資料 (平行處理加速)
        const item = Office.context.mailbox.item;
        const [resTo, resCc, resAtt] = await Promise.all([
            getAsyncPromise(item.to),
            getAsyncPromise(item.cc),
            getAsyncPromise(item.attachments)
        ]);

        // C. 整理資料包
        fetchedData = {
            subject: item.subject, 
            recipients: [
                ...(resTo.value ? resTo.value.map(r => ({...r, type: 'To'})) : []),
                ...(resCc.value ? resCc.value.map(r => ({...r, type: 'Cc'})) : [])
            ],
            attachments: resAtt.value || []
        };

        // 注意：這裡我們不 call event.completed()，我們要等 Dialog 說它 Ready 之後傳資料給它
        // 但為了避免 Dialog 載入失敗導致卡死，可以設個簡單的 Timeout 保險 (非必要)

    } catch (error) {
        console.error("OpenDialog Error:", error);
        // 如果出錯，至少要讓轉圈圈停下來
        if (event) event.completed();
    }
}

// 輔助：把 getAsync 轉成 Promise
function getAsyncPromise(itemProperty) {
    return new Promise((resolve) => {
        itemProperty.getAsync((result) => {
            resolve(result); // 即使失敗也 resolve，由邏輯判斷 status
        });
    });
}

// 3. 處理訊息 (核心邏輯修改)
function processMessage(arg) {
    const message = arg.message; // 可能是字串或 JSON 字串

    // A. Dialog 載入完成，請求資料
    if (message === "DIALOG_READY") {
        if (dialog && fetchedData) {
            // 把資料推給 Dialog
            const jsonString = JSON.stringify(fetchedData);
            dialog.messageChild(jsonString);
            
            // 資料送出去了，任務完成！現在可以停止 Ribbon 的轉圈圈了
            // 注意：這裡假設 openDialog 的 event 全域變數可能存取不到，
            // 但在 FunctionFile 模式下，我們通常無法在此處 access 'event'。
            // 修正策略：在 openDialog 結束時就 call completed，讓 dialog 獨立運作。
            // 但如果 openDialog 結束太快，Dialog 可能還沒由收到 messageChild。
            // 最保險的做法：資料抓完就可以 completed 了。
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

// 修正後的 openDialog 結尾邏輯：
// 為了避免複雜的 event 傳遞，我們改採「抓完資料就結束轉圈，然後等 Dialog Ready 再送資料」
// 覆寫上面的 openDialog 後半段：

async function openDialog_Optimized(event) {
    try {
        const item = Office.context.mailbox.item;
        
        // 1. 平行抓資料
        const [resTo, resCc, resAtt] = await Promise.all([
            getAsyncPromise(item.to),
            getAsyncPromise(item.cc),
            getAsyncPromise(item.attachments)
        ]);

        fetchedData = {
            subject: item.subject, 
            recipients: [
                ...(resTo.value ? resTo.value.map(r => ({...r, type: 'To'})) : []),
                ...(resCc.value ? resCc.value.map(r => ({...r, type: 'Cc'})) : [])
            ],
            attachments: resAtt.value || []
        };

        // 2. 開視窗
        const url = 'https://icy-moss-034796200.2.azurestaticapps.net/dialog.html';
        Office.context.ui.displayDialogAsync(
            url, 
            { height: 60, width: 50, displayInIframe: true },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                }
                
                // 3. 關鍵：無論成功與否，立刻停止轉圈圈！
                // 視窗已經開了，資料也存好了，剩下的讓 Dialog 和 processMessage 去溝通
                if (event) event.completed();
            }
        );

    } catch (e) {
        console.error(e);
        if (event) event.completed();
    }
}

// 綁定
if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;
g.openDialog = openDialog_Optimized; // 使用優化版