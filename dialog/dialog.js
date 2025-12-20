/* global Office, document, localStorage */

Office.onReady(() => {
    // 按鈕事件綁定
    document.getElementById("btnSend").onclick = () => {
        if (!document.getElementById("btnSend").disabled) {
            Office.context.ui.messageParent("VERIFIED_PASS");
        }
    };
    document.getElementById("btnCancel").onclick = () => {
        Office.context.ui.messageParent("CANCEL");
    };

    // --- 修改重點 ---
    try {
        // 【修正 1】從 LocalStorage 讀取資料
        const dataString = localStorage.getItem("outlook_verify_data");

        if (dataString) {
            const data = JSON.parse(dataString);
            renderData(data); // 渲染畫面
            
            // (選擇性) 讀完後可以清除，保持乾淨
            // localStorage.removeItem("outlook_verify_data");
        } else {
            document.getElementById("recipients-list").innerText = "無法讀取信件資料 (Storage Empty)";
        }
    } catch (e) {
        // 如果出錯，直接把錯誤顯示在畫面上，方便除錯
        document.getElementById("recipients-list").innerHTML = `<span style="color:red">Error: ${e.message}</span>`;
    }
});

// 以下渲染函式不用動，維持原樣即可
function renderData(data) {
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    
    // 簡單模擬使用者 Domain (實務上可從 commands.js 傳入)
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
            // 簡單的 Domain 比對邏輯
            let personDomain = "";
            if (email.includes("@")) {
                personDomain = email.split('@')[1];
            }
            
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
    
    // 渲染附件
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
    // 如果沒有任何項目要檢查，預設也可以過
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