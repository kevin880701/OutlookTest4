/* global Office, document */

Office.onReady(() => {
    // 1. 初始化按鈕
    document.getElementById("btnSend").onclick = () => {
        if (!document.getElementById("btnSend").disabled) {
            Office.context.ui.messageParent("VERIFIED_PASS");
        }
    };
    document.getElementById("btnCancel").onclick = () => {
        Office.context.ui.messageParent("CANCEL");
    };

    // 2. 【關鍵】直接從 URL 讀取資料
    try {
        const urlParams = new URLSearchParams(window.location.search);
        const dataString = urlParams.get('data');

        if (dataString) {
            const data = JSON.parse(decodeURIComponent(dataString));
            renderData(data); // 有資料就一定會畫出來
        } else {
            document.getElementById("recipients-list").innerText = "錯誤：網址中沒有資料";
        }
    } catch (e) {
        console.error(e);
        document.getElementById("recipients-list").innerText = "資料解析失敗 (可能資料量過大)";
    }
});

// 渲染函式 (維持不變，請直接使用)
function renderData(data) {
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    const userDomain = "outlook.com"; 

    // 收件人
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
            container.appendChild(row);
        });
    } else {
        container.innerHTML = "無收件人";
    }
    
    // 附件
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