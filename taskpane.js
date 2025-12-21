/* global Office, document */

Office.onReady(() => {
    loadItemData();
    document.getElementById("btnVerify").onclick = markAsVerified;
});

function loadItemData() {
    const item = Office.context.mailbox.item;

    Promise.all([
        new Promise(r => item.to.getAsync(x => r(x.value || []))),
        new Promise(r => item.cc.getAsync(x => r(x.value || []))),
        new Promise(r => item.getAttachmentsAsync(x => r(x.value || [])))
    ]).then(([to, cc, attachments]) => {
        renderList(to, cc, attachments);
    }).catch(err => {
        console.error(err);
        document.getElementById("recipients-list").innerText = "讀取錯誤: " + err.message;
    });
}

function renderList(to, cc, attachments) {
    const rList = document.getElementById("recipients-list");
    const aList = document.getElementById("attachments-list");
    
    rList.innerHTML = "";
    aList.innerHTML = "";

    let hasItems = false;

    [...to, ...cc].forEach((p, i) => {
        hasItems = true;
        const type = to.includes(p) ? "To" : "Cc";
        const div = document.createElement("div");
        div.className = "item-row";
        div.innerHTML = `
            <input type='checkbox' class='verify-check' id='r_${i}' onchange='checkAllChecked()'>
            <label for='r_${i}'>
                <span style="font-weight:bold; color:#666">[${type}]</span> 
                ${p.displayName || p.emailAddress}
            </label>
        `;
        rList.appendChild(div);
    });

    if ([...to, ...cc].length === 0) {
        rList.innerHTML = "<div style='padding:5px; font-size:12px'>(無收件人)</div>";
    }

    attachments.forEach((a, i) => {
        hasItems = true;
        const div = document.createElement("div");
        div.className = "item-row";
        div.innerHTML = `
            <input type='checkbox' class='verify-check' id='a_${i}' onchange='checkAllChecked()'>
            <label for='a_${i}'>📎 ${a.name}</label>
        `;
        aList.appendChild(div);
    });

    if (attachments.length === 0) {
        aList.innerHTML = "<div style='padding:5px; font-size:12px'>(無附件)</div>";
    }

    if (!hasItems) {
        enableButton();
    }
}

window.checkAllChecked = function() {
    const all = document.querySelectorAll(".verify-check");
    let pass = true;
    all.forEach(c => { if(!c.checked) pass = false; });
    
    if (pass) enableButton();
    else disableButton();
};

function enableButton() {
    const btn = document.getElementById("btnVerify");
    btn.disabled = false;
    btn.classList.add("active");
    btn.innerText = "✅ 確認無誤 (解除鎖定)";
}

function disableButton() {
    const btn = document.getElementById("btnVerify");
    btn.disabled = true;
    btn.classList.remove("active");
    btn.innerText = "請勾選所有項目...";
}

function markAsVerified() {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        props.set("isVerified", true);
        props.saveAsync((saveResult) => {
            if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                document.getElementById("btnVerify").style.display = "none";
                document.getElementById("status-msg").style.display = "block";
            } else {
                document.getElementById("btnVerify").innerText = "❌ 儲存失敗，請重試";
            }
        });
    });
}