/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("saveBtn").onclick = saveSettings;
        loadSettings();
    }
});

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

function saveSettings() {
    const cc = document.getElementById("ccInput").value.trim();
    const bcc = document.getElementById("bccInput").value.trim();
    const enabled = document.getElementById("enabledCheck").checked;

    try {
        // RoamingSettings (Outlookアカウントに紐づいて同期)
        const roamingSettings = Office.context.roamingSettings;
        roamingSettings.set("auto_cc_emails", cc);
        roamingSettings.set("auto_bcc_emails", bcc);
        roamingSettings.set("auto_cc_bcc_enabled", enabled);

        // localStorage (ローカルキャッシュ)
        localStorage.setItem("auto_cc_emails", cc);
        localStorage.setItem("auto_bcc_emails", bcc);
        localStorage.setItem("auto_cc_bcc_enabled", enabled.toString());

        roamingSettings.saveAsync((result) => {
            const statusDiv = document.getElementById("status-message");
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                statusDiv.innerText = "設定を正常に保存しました！";
                statusDiv.className = "success";
                statusDiv.style.display = "block";
            } else {
                statusDiv.innerText = "設定の保存に一部失敗した可能性があります。";
                statusDiv.style.display = "block";
            }

            setTimeout(() => {
                statusDiv.style.display = "none";
            }, 3000);
        });
    } catch (e) {
        alert("保存中にエラーが発生しました: " + e.message);
    }
}
