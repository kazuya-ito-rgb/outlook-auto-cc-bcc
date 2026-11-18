/* global Office */

/**
 * 新規メール作成（または返信・転送）時に自動的にトリガーされるイベントハンドラー
 * @param {Office.AddinCommands.Event} event
 */
function onNewMessageComposeHandler(event) {
    try {
        const roamingSettings = Office.context.roamingSettings;
        let enabled = roamingSettings.get("auto_cc_bcc_enabled");
        let ccStr = roamingSettings.get("auto_cc_emails");
        let bccStr = roamingSettings.get("auto_bcc_emails");

        // フォールバック (localStorage)
        if (enabled === undefined || enabled === null) {
            const localEnabled = (typeof localStorage !== "undefined") ? localStorage.getItem("auto_cc_bcc_enabled") : null;
            enabled = localEnabled !== null ? localEnabled === "true" : true;
        }
        if (ccStr === undefined || ccStr === null) {
            ccStr = (typeof localStorage !== "undefined") ? (localStorage.getItem("auto_cc_emails") || "") : "";
        }
        if (bccStr === undefined || bccStr === null) {
            bccStr = (typeof localStorage !== "undefined") ? (localStorage.getItem("auto_bcc_emails") || "") : "";
        }

        // 自動入力が無効の場合はスキップ
        if (enabled === false) {
            event.completed();
            return;
        }

        const item = Office.context.mailbox.item;

        // メールアドレス文字列をパースして配列に変換
        const parseRecipients = (str) => {
            if (!str) return [];
            return str
                .split(/[;,]/)
                .map((addr) => addr.trim())
                .filter((addr) => addr.length > 0)
                .map((addr) => ({
                    emailAddress: addr,
                    displayName: addr
                }));
        };

        const ccRecipients = parseRecipients(ccStr);
        const bccRecipients = parseRecipients(bccStr);

        let pendingTasks = 0;
        const checkDone = () => {
            pendingTasks--;
            if (pendingTasks <= 0) {
                event.completed();
            }
        };

        if (ccRecipients.length > 0) {
            pendingTasks++;
            item.cc.addAsync(ccRecipients, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("CC追加失敗:", asyncResult.error.message);
                }
                checkDone();
            });
        }

        if (bccRecipients.length > 0) {
            pendingTasks++;
            item.bcc.addAsync(bccRecipients, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("BCC追加失敗:", asyncResult.error.message);
                }
                checkDone();
            });
        }

        // CCもBCCも追加対象がない場合
        if (pendingTasks === 0) {
            event.completed();
        }
    } catch (err) {
        console.error("onNewMessageComposeHandler エラー:", err);
        event.completed();
    }
}

// Office Add-ins にイベントハンドラー関数を登録
if (typeof Office !== "undefined" && Office.actions) {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
