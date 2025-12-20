/* global Office, document */

Office.onReady(() => {
    // 1. 註冊接收器：準備接收來自 Parent 的廣播
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onParentMessageReceived
    );

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
             renderData(data); // 渲染畫面
             
             // 禮貌性地回覆一聲：我收到了，別再廣播了
             // (如果 messageParent 失敗也沒關係，Parent 8秒後會自動停)
             Office.context.ui.messageParent("DATA_RECEIVED");
        }
    } catch (e) {
        console.error("解析錯誤", e);
    }
}

// 渲染函式 (維持不變)
function renderData(data) {
    document.getElementById("subject").innerText = data.subject || "(無主旨)";
    
    // ... 收件人渲染 ...
    const recipientContainer = document.getElementById("recipients-list");
    recipientContainer.innerHTML = "";
    const userDomain = "outlook.com"; 

    if (data.recipients && data.recipients.length > 0) {
        data.recipients.forEach((person, index) => {
            const row = document.createElement("div");
            row.className = "item-row";
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.className = "verify-check";
            checkbox.id = `recip_${index}`;
            checkbox.onchange = checkAllChecked;

            const email = person.emailAddress || "";
            let personDomain = "";
            if (email.includes("@")) personDomain = email.split('@')[1];
            const isExternal = personDomain && personDomain !== userDomain;

            let html = `<b>[${person.type}]</b> ${person.displayName || "Unknown"} <br><small>${email}</small>`;
            if (isExternal) {
                html += ` <span class="external-tag" style="color:red; border:1px solid red; font-size:10px; margin-left:5px;">External</span>`;
                checkbox.checked = false; 
            } else {
                checkbox.checked = true; 
            }
            const label = document.createElement("label");
            label.htmlFor = `recip_${index}`;
            label.innerHTML = html;
            row.appendChild(checkbox);
            row.appendChild(label);
            recipientContainer.appendChild(row);
        });
    } else {
        recipientContainer.innerHTML = "無收件人";
    }

    // ... 附件渲染 ...
    const attContainer = document.getElementById("attachments-list");
    attContainer.innerHTML = "";
    if (data.attachments && data.attachments.length > 0) {
        data.attachments.forEach((att, index) => {
             const row = document.createElement("div");
             row.className = "item-row";
             const checkbox = document.createElement("input");
             checkbox.type = "checkbox";
             checkbox.className = "verify-check";
             checkbox.id = `att_${index}`;
             checkbox.onchange = checkAllChecked;
             const label = document.createElement("label");
             label.htmlFor = `att_${index}`;
             label.innerText = `📎 ${att.name}`;
             row.appendChild(checkbox);
             row.appendChild(label);
             attContainer.appendChild(row);
        });
    } else {
        attContainer.innerText = "無附件";
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
        btn.innerText = "確認完畢，允許發送";
    } else {
        btn.style.opacity = "0.5";
        btn.style.cursor = "not-allowed";
        btn.innerText = "請勾選所有項目";
    }
}