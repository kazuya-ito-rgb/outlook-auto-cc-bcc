/* global Office */

/**
 * 新規メール作成（または返信・転送）時に自動的にトリガーされるイベントハンドラー
 * @param {Office.AddinCommands.Event} event
 */
function onNewMessageComposeHandler(event) {
    let completed = false;
    const finish = () => {
        if (completed) return;
        completed = true;
        try {
            event.completed();
        } catch (e) {
            console.error("event.completed error:", e);
        }
    };

    // タイムアウト保護（最大4秒で必ず完了させてフリーズを防止）
    setTimeout(() => {
        finish();
    }, 4000);

    try {
        const roamingSettings = Office.context.roamingSettings;
        if (!roamingSettings) {
            finish();
            return;
        }

        const enabled = roamingSettings.get("auto_cc_bcc_enabled");
        const ccStr = roamingSettings.get("auto_cc_emails");
        const bccStr = roamingSettings.get("auto_bcc_emails");

        // 自動入力が無効（明示的にfalse）の場合はスキップ
        if (enabled === false || enabled === "false") {
            finish();
            return;
        }

        const parseEmails = (str) => {
            if (!str) return [];
            return str
                .split(/[;,]/)
                .map((addr) => addr.trim())
                .filter((addr) => addr.length > 0);
        };

        const ccList = parseEmails(ccStr);
        const bccList = parseEmails(bccStr);

        if (ccList.length === 0 && bccList.length === 0) {
            finish();
            return;
        }

        const item = Office.context.mailbox.item;
        if (!item) {
            finish();
            return;
        }

        let pending = 0;
        const checkDone = () => {
            pending--;
            if (pending <= 0) {
                finish();
            }
        };

        // CC の追加
        if (ccList.length > 0 && item.cc && item.cc.addAsync) {
            pending++;
            const ccRecipients = ccList.map(email => ({ emailAddress: email, displayName: email }));
            item.cc.addAsync(ccRecipients, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    // 文字列配列でのフォールバック
                    item.cc.addAsync(ccList, () => checkDone());
                } else {
                    checkDone();
                }
            });
        }

        // BCC の追加
        if (bccList.length > 0 && item.bcc && item.bcc.addAsync) {
            pending++;
            const bccRecipients = bccList.map(email => ({ emailAddress: email, displayName: email }));
            item.bcc.addAsync(bccRecipients, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    // 文字列配列でのフォールバック
                    item.bcc.addAsync(bccList, () => checkDone());
                } else {
                    checkDone();
                }
            });
        }

        if (pending === 0) {
            finish();
        }
    } catch (err) {
        console.error("onNewMessageComposeHandler エラー:", err);
        finish();
    }
}

// Office Add-ins にイベントハンドラーを登録（トップレベルおよび Office.onReady で二重登録）
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
if (Office.onReady) {
    Office.onReady(() => {
        try {
            Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
        } catch (e) {}
    });
}
