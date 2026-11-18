# Outlook New 自動 CC / BCC 挿入アドイン

新しい Outlook (New Outlook for Windows / Outlook on the web / Mac) に対応した、**メール新規作成時に指定した宛先を自動的に CC / BCC に追加する Office Web アドイン**です。

---

## 🌟 主な機能
- **自動挿入**: 新規メール作成（または返信・転送）を開いた瞬間に、事前設定したアドレスを自動で CC / BCC 欄に追加
- **簡単な設定変更**: メール作成画面のリボン/メニューにある「CC/BCC設定」ボタンから、いつでも対象アドレスを変更可能
- **複数アドレス対応**: セミコロン `;` で区切って複数の CC/BCC アドレスを登録可能
- **同期保存**: 設定は Microsoft 365 アカウント（RoamingSettings）に保存されるため、同一アカウントであれば別端末でも設定が維持されます

---

## 📁 フォルダ構成
```text
outlook-auto-cc-bcc-addin/
├── manifest.xml           # Outlookアドイン定義（LaunchEvent: OnNewMessageCompose）
├── assets/
│   ├── icon-16.png        # アイコン（16x16）
│   ├── icon-32.png        # アイコン（32x32）
│   └── icon-80.png        # アイコン（80x80）
├── src/
│   ├── launchevent.html   # バックグラウンド起動用HTML
│   ├── launchevent.js     # メール作成時に自動入力する処理
│   ├── taskpane.html      # CC/BCC設定画面UI
│   └── taskpane.js        # 設定の保存・読込処理
└── README.md
```

---

## 🚀 導入手順（GitHub Pages を利用）

### 1. GitHub リポジトリの作成とアップロード
1. GitHub で新規リポジトリを作成します（例: `outlook-auto-cc-bcc`）。
   - **Public**（公開）リポジトリとして作成することを推奨します（GitHub Pages が無料プランで使えるため）。
2. 本フォルダ内のファイル・フォルダ（`manifest.xml`, `assets/`, `src/`）をリポジトリのルートにプッシュまたはアップロードします。

### 2. GitHub Pages の有効化
1. リポジトリの **Settings** > **Pages** を開きます。
2. **Build and deployment** の Source で **Deploy from a branch** を選択し、Branch を `main`（または `master`）の `/ (root)` に設定して **Save** をクリックします。
3. 数分後、公開URLが発行されます：
   `https://<YOUR_GITHUB_USERNAME>.github.io/<YOUR_REPO_NAME>/`

### 3. `manifest.xml` の URL を書き換える
`manifest.xml` を開き、以下のプレースホルダーをご自身の GitHub Pages の URL に置換して保存・コミットします。

- 置換前:
  `https://YOUR_GITHUB_USERNAME.github.io/YOUR_REPO_NAME`
- 置換後（例: ユーザー名が `taro`、リポジトリ名が `outlook-auto-cc-bcc` の場合）:
  `https://taro.github.io/outlook-auto-cc-bcc`

### 4. Outlook New への登録（サイドロード）
1. **新しい Outlook** またはブラウザで [Outlook on the web](https://outlook.office.com/) を開きます。
2. 画面左上の「**新しいメール**」をクリックしてメール作成画面を開きます。
3. リボンの「**アプリ**（または「…」その他の操作）」>「**アプリの追加**」をクリックします。
4. ポップアップ画面の左メニューまたはタブにある「**アドインの管理**」を選択します。
5. 「**カスタム アドイン**」タブを開き、「**カスタム アドインの追加**」をクリックします。
6. 「**ファイルから追加...**」を選択し、先ほど書き換えた `manifest.xml` を選択してアップロードします。
7. 警告画面（開発者向けアドインの確認）が表示されたら「**インストール**」をクリックします。

---

## 💡 使い方

1. **初期設定**:
   - メール作成画面のリボンメニューに「**宛先自動挿入**」グループと「**CC/BCC設定**」ボタンが表示されます。
   - クリックすると右側に設定パネルが開きます。
   - 自動追加したいメールアドレス（セミコロン区切り）を入力し、「**設定を保存**」をクリックします。
2. **動作確認**:
   - 設定後、一度メール作成画面を閉じて、再度「**新しいメール**」をクリックします。
   - 設定した CC / BCC アドレスが自動的に入力されていることを確認してください。
