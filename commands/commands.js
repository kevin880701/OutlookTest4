/* global Office */

Office.onReady();

function validateSend(event) {
    // 1. 讀取郵件資訊
    const item = Office.context.mailbox.item;

    // 定義我們要檢查的項目 (例如：檢查主旨)
    const pSubject = new Promise((resolve) => {
        item.subject.getAsync((r) => resolve(r.status === 'succeeded' ? r.value : ''));
    });

    const pTo = new Promise((resolve) => {
        item.to.getAsync((r) => resolve(r.status === 'succeeded' ? r.value : []));
    });

    Promise.all([pSubject, pTo])
        .then((values) => {
            const [subject, to] = values;

            // --- 檢查邏輯區 (這是您要自訂的地方) ---

            // 範例 1: 如果主旨是空的，就擋下來
            if (!subject || subject.trim() === "") {
                // 【關鍵修改】Mac 不會彈窗，而是顯示這行 errorMessage
                event.completed({ 
                    allowEvent: false, 
                    errorMessage: "⚠️ 傳送失敗：您忘記填寫主旨了！請填寫後再試。" 
                });
                return;
            }

            // 範例 2: 檢查收件人是否有外部信箱 (簡單示範)
            // const hasExternal = to.some(t => !t.emailAddress.includes("acer-ast.com"));
            // if (hasExternal) {
            //      注意：Smart Alerts 目前主要是用來「擋下錯誤」。
            //      如果要像以前一樣「詢問後放行」，在新版 Mac 上比較難做，
            //      通常只能選擇「直接擋下」請使用者確認。
            //      event.completed({ allowEvent: false, errorMessage: "⚠️ 偵測到外部收件人，請確認後再次傳送。" });
            //      return;
            // }

            // --- 檢查通過 ---
            // 如果沒問題，就允許傳送
            event.completed({ allowEvent: true });
        })
        .catch((error) => {
            console.error("檢查發生錯誤:", error);
            // 發生未知錯誤時，通常選擇放行，以免使用者無法寄信
            event.completed({ allowEvent: true });
        });
}