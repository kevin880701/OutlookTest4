/* global Office, document */

Office.onReady(() => {
    loadItemData();

    document.getElementById("btnCancel").onclick = () => {
        Office.context.ui.messageParent("CANCEL");
    };

    document.getElementById("btnSend").onclick = () => {
        // 只有在按鈕啟動時才發送訊號
        if (!document.getElementById("btnSend").disabled) {
            Office.context.ui.messageParent("VERIFIED_PASS");
        }
    };
});

// 讀取收件人與附件
function loadItemData() {
    const item = Office.context.mailbox.item;
    
    // 平行讀取 To, CC, Attachments
    // 這裡使用巢狀 callback 簡單示範，實務上可用 Promise 封裝
    item.to.getAsync((resultTo) => {
        item.cc.getAsync((resultCc) => {
            item.attachments.getAsync((resultAtt) => {
                
                const recipients = [
                    ...resultTo.value.map(r => ({...r, type: 'To'})),
                    ...resultCc.value.map(r => ({...r, type: 'Cc'}))
                ];
                
                const attachments = resultAtt.value;

                renderRecipients(recipients);
                renderAttachments(attachments);
                
                // 執行一次檢查，看是否需要啟用按鈕 (例如清單為空時)
                checkAllChecked();
            });
        });
    });
}

function renderRecipients(list) {
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";

    if (list.length === 0) {
        container.innerHTML = "<div>(無收件人)</div>";
        return;
    }

    // 取得當前使用者的 Domain 用來比對 (這裡簡單抓 user profile)
    const userEmail = Office.context.mailbox.userProfile.emailAddress;
    const userDomain = userEmail.split('@')[1];

    list.forEach((person, index) => {
        const row = document.createElement("div");
        row.className = "item-row";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.className = "verify-check";
        checkbox.id = `recip_${index}`;
        checkbox.onchange = checkAllChecked; // 綁定變更事件

        const label = document.createElement("label");
        label.htmlFor = `recip_${index}`;
        
        // 判斷是否為外部信箱
        const personDomain = person.emailAddress.split('@')[1];
        const isExternal = personDomain !== userDomain;
        
        let htmlText = `<b>[${person.type}]</b> ${person.displayName} &lt;${person.emailAddress}&gt;`;
        if (isExternal) {
            htmlText += ` <span class="external-tag">External</span>`;
            // 外部信箱預設不勾選，內部可考慮預設勾選
            checkbox.checked = false; 
        } else {
            // 內部信箱預設勾選 (模擬您的截圖需求)
            checkbox.checked = true;
        }

        label.innerHTML = htmlText;

        row.appendChild(checkbox);
        row.appendChild(label);
        container.appendChild(row);
    });
}

function renderAttachments(list) {
    const container = document.getElementById("attachments-list");
    container.innerHTML = "";

    if (list.length === 0) {
        container.innerHTML = "<div style='color:#888'>無附件</div>";
        return;
    }

    list.forEach((att, index) => {
        const row = document.createElement("div");
        row.className = "item-row";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.className = "verify-check";
        checkbox.id = `att_${index}`;
        checkbox.onchange = checkAllChecked;

        const label = document.createElement("label");
        label.htmlFor = `att_${index}`;
        label.innerText = `📎 ${att.name} (${Math.round(att.size / 1024)} KB)`;

        row.appendChild(checkbox);
        row.appendChild(label);
        container.appendChild(row);
    });
}

// 核心邏輯：檢查所有 checkbox 是否都勾選了
function checkAllChecked() {
    const allChecks = document.querySelectorAll(".verify-check");
    let allPassed = true;

    allChecks.forEach(ck => {
        if (!ck.checked) allPassed = false;
    });

    const btn = document.getElementById("btnSend");
    if (allPassed && allChecks.length > 0) {
        btn.disabled = false;
        btn.classList.add("active");
        btn.innerText = "確認完畢，允許發送";
    } else {
        btn.disabled = true;
        btn.classList.remove("active");
        if (allChecks.length === 0) {
             // 如果完全沒收件人沒附件，或許直接允許？
             btn.disabled = false;
             btn.classList.add("active");
        } else {
             btn.innerText = "請勾選所有項目";
        }
    }
}