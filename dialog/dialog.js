/* global Office, document */

Office.onReady(() => {
    // 1. 註冊接收器：準備被資料「砸中」
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onParentMessageReceived
    );

    // 2. 不需要 setInterval 去喊話了，被動等待即可

    // 按鈕綁定
    document.getElementById("btnSend").onclick = () => {
        if (!document.getElementById("btnSend").disabled) {
            Office.context.ui.messageParent("VERIFIED_PASS");
        }
    };
    document.getElementById("btnCancel").onclick = () => {
        Office.context.ui.messageParent("CANCEL");
    };
});

// 當收到 Parent 廣播來的資料時
function onParentMessageReceived(arg) {
    try {
        const message = arg.message;
        const data = JSON.parse(message); // 解析資料
        
        // 確保資料有效
        if (data && data.recipients) {
             renderData(data); // 渲染畫面 (Loading 會消失)
             
             // 禮貌性回覆：我收到了，別再吵了
             Office.context.ui.messageParent("DATA_RECEIVED");
        }
    } catch (e) {
        console.error("解析錯誤", e);
    }
}

// 渲染函式 (維持不變，請保留您原本的 renderData)
function renderData(data) {
    // 請直接使用您原本寫好的渲染邏輯
    // 這裡只寫開頭示意
    document.getElementById("subject").innerText = data.subject || "(無主旨)";
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    
    // ... (請保留您原本的收件人與附件渲染代碼) ...
    // ... 記得呼叫 checkAllChecked() ...
    
    // 為了完整性，這裡簡單補上您之前的邏輯框架
    if (data.recipients) {
        data.recipients.forEach((person, index) => {
            const row = document.createElement("div");
            row.className = "item-row";
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.className = "verify-check";
            checkbox.id = `recip_${index}`;
            // 簡單處理...
            const label = document.createElement("label");
            label.innerText = person.displayName; 
            row.appendChild(checkbox);
            row.appendChild(label);
            container.appendChild(row);
        });
        // 請用您原本詳細的 renderData 取代這裡
    }
    checkAllChecked();
}

function checkAllChecked() {
    const all = document.querySelectorAll(".verify-check");
    let pass = true;
    all.forEach(c => { if(!c.checked) pass = false; });
    const btn = document.getElementById("btnSend");
    if (all.length === 0) pass = true;
    btn.disabled = !pass;
    if (pass) {
        btn.style.opacity = "1";
        btn.style.cursor = "pointer";
    } else {
        btn.style.opacity = "0.5";
        btn.style.cursor = "not-allowed";
    }
}