/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("saveBtn").onclick = handleSaveAndApply;
        document.getElementById("applyNowBtn").onclick = handleApplyOnly;
        loadSettings();
    }
});

function parseEmails(str) {
    if (!str) return [];
    return str
        .split(/[;,]/)
        .map((addr) => addr.trim())
        .filter((addr) => addr.length > 0);
}

function loadSettings() {
    try {
        const roamingSettings = Office.context.roamingSettings;
        let cc = roamingSettings.get("auto_cc_emails");
        let bcc = roamingSettings.get("auto_bcc_emails");
        let enabled = roamingSettings.get("auto_cc_bcc_enabled");

        // フォールバック (localStorage)
        if (cc === undefined || cc === null) {
            cc = localStorage.getItem("auto_cc_emails") || "";
        }
        if (bcc === undefined || bcc === null) {
            bcc = localStorage.getItem("auto_bcc_emails") || "";
        }
        if (enabled === undefined || enabled === null) {
            const localEnabled = localStorage.getItem("auto_cc_bcc_enabled");
            enabled = localEnabled !== null ? localEnabled === "true" : true;
        }

        document.getElementById("ccInput").value = cc || "";
        document.getElementById("bccInput").value = bcc || "";
        document.getElementById("enabledCheck").checked = enabled !== false;
    } catch (e) {
        console.error("設定の読み込みに失敗しました:", e);
    }
}

function applyToCurrentEmail(ccList, bccList, callback) {
    const item = Office.context.mailbox.item;
    if (!item) {
        if (callback) callback();
        return;
    }

    let pending = 0;
    const done = () => {
        pending--;
        if (pending <= 0 && callback) {
            callback();
        }
    };

    if (ccList.length > 0 && item.cc && item.cc.addAsync) {
        pending++;
        const ccRecipients = ccList.map(email => ({ emailAddress: email, displayName: email }));
        item.cc.addAsync(ccRecipients, (res) => {
            if (res.status === Office.AsyncResultStatus.Failed) {
                // 文字列配列でのフォールバック
                item.cc.addAsync(ccList, () => done());
            } else {
                done();
            }
        });
    }

    if (bccList.length > 0 && item.bcc && item.bcc.addAsync) {
        pending++;
        const bccRecipients = bccList.map(email => ({ emailAddress: email, displayName: email }));
        item.bcc.addAsync(bccRecipients, (res) => {
            if (res.status === Office.AsyncResultStatus.Failed) {
                // 文字列配列でのフォールバック
                item.bcc.addAsync(bccList, () => done());
            } else {
                done();
            }
        });
    }

    if (pending === 0 && callback) {
        callback();
    }
}

function handleSaveAndApply() {
    const ccStr = document.getElementById("ccInput").value.trim();
    const bccStr = document.getElementById("bccInput").value.trim();
    const enabled = document.getElementById("enabledCheck").checked;

    try {
        const roamingSettings = Office.context.roamingSettings;
        roamingSettings.set("auto_cc_emails", ccStr);
        roamingSettings.set("auto_bcc_emails", bccStr);
        roamingSettings.set("auto_cc_bcc_enabled", enabled);

        localStorage.setItem("auto_cc_emails", ccStr);
        localStorage.setItem("auto_bcc_emails", bccStr);
        localStorage.setItem("auto_cc_bcc_enabled", enabled.toString());

        roamingSettings.saveAsync((result) => {
            const ccList = parseEmails(ccStr);
            const bccList = parseEmails(bccStr);

            // 現在開いているメールにも即座に反映
            applyToCurrentEmail(ccList, bccList, () => {
                const statusDiv = document.getElementById("status-message");
                statusDiv.innerHTML = "✅ <b>設定を保存し、現在のメールに反映しました！</b><br>次回以降は「新しいメール」作成時に自動入力されます。";
                statusDiv.className = "success";
                statusDiv.style.display = "block";

                setTimeout(() => {
                    statusDiv.style.display = "none";
                }, 5000);
            });
        });
    } catch (e) {
        alert("保存中にエラーが発生しました: " + e.message);
    }
}

function handleApplyOnly() {
    const ccStr = document.getElementById("ccInput").value.trim();
    const bccStr = document.getElementById("bccInput").value.trim();
    const ccList = parseEmails(ccStr);
    const bccList = parseEmails(bccStr);

    applyToCurrentEmail(ccList, bccList, () => {
        const statusDiv = document.getElementById("status-message");
        statusDiv.innerHTML = "✅ <b>現在のメールにCC/BCCを挿入しました！</b>";
        statusDiv.className = "success";
        statusDiv.style.display = "block";
        setTimeout(() => {
            statusDiv.style.display = "none";
        }, 3000);
    });
}
