/* global Office, document, window */

Office.onReady(() => {
    // 按鈕綁定
    document.getElementById("btnSend").onclick = () => {
        if (!document.getElementById("btnSend").disabled) {
            Office.context.ui.messageParent("VERIFIED_PASS");
        }
    };
    document.getElementById("btnCancel").onclick = () => {
        Office.context.ui.messageParent("CANCEL");
    };

    // ★★★ 關鍵修改：直接從網址列讀取資料 ★★★
    try {
        const urlParams = new URLSearchParams(window.location.search);
        const dataStr = urlParams.get('data');

        if (dataStr) {
            const data = JSON.parse(decodeURIComponent(dataStr));
            renderData(data); // 直接渲染，不用等 Parent
        } else {
            document.getElementById("recipients-list").innerText = "未接收到資料 (URL Parameter Missing)";
        }
    } catch (e) {
        document.getElementById("recipients-list").innerText = "資料解析失敗: " + e.message;
    }
});

// 渲染函式 (維持您原本的，這裡只列出開頭，內容不用動)
function renderData(data) {
    // ... 您原本的 renderData 程式碼 ...
    // ... 包含 attachments 和 recipients 的渲染 ...
    const container = document.getElementById("recipients-list");
    // (請保留您原本完整的 renderData 邏輯)
    // ...
    // 最後記得呼叫 checkAllChecked()
    
    // 這裡我簡寫範例，您只需保留原本 dialog.js 下半部的 renderData 即可
    const rList = document.getElementById("recipients-list");
    rList.innerHTML = ""; 
    // ... (您的渲染邏輯)
    
    // 如果您原本的 dialog.js 已經寫好 renderData，就不用動下面，只要改上面的 Office.onReady 即可
    // 為了保險，建議您直接把原本的 renderData 函數貼在上面那段代碼下面
}

// 補上 checkAllChecked (維持不變)
function checkAllChecked() {
    const all = document.querySelectorAll(".verify-check");
    let pass = true;
    all.forEach(c => { if(!c.checked) pass = false; });
    const btn = document.getElementById("btnSend");
    if (all.length === 0) pass = true; // 如果沒東西要檢查，直接過

    btn.disabled = !pass;
    if (pass) {
        btn.style.opacity = "1";
        btn.style.cursor = "pointer";
        btn.innerText = "確認完畢，允許發送";
    } else {
        btn.style.opacity = "0.5";
        btn.style.cursor = "not-allowed";
        btn.innerText = "請勾選所有項目";
    }
}