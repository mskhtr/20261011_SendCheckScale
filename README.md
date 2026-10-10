# Outlook への入れ方と外し方

Outlook は、外のスクリプトから VBA を書き換えられません。このフォルダのマクロは、Outlook の VBA エディタで手作業で入れます。

`src` の `modAttachmentCheck.bas` と `modAttachmentCheckTest.bas` は入れません。

## インストール

1. Outlook を開きます。
2. Alt+F11 で VBA エディタを開きます。
3. ファイル → ファイルのインポート で、次を入れます。
   - `src\modSendCheck.bas`
   - `src\modWheelScroll.bas`
   - `src\ClsButtonHandler.cls`
   - `src\ClsCardClick.cls`
   - `src\ClsScrollDrag.cls`
4. 挿入 → ユーザーフォーム で、空のフォームを作ります。フォームの上には部品を置きません。
5. プロパティの (オブジェクト名) を `frmSendCheck` にします。
6. そのフォームのコードを開き、`src\frmSendCheck.code.bas` の `Option Explicit` から末尾までを貼ります。それより前のコメントは貼りません。
7. Alt+F8 でマクロ「送信前チェックを開く」を実行します。アドイン タブに同じ名前のボタンが付きます。次からは、そのボタンか Alt+F8 から開きます。
8. Outlook を普通に終了すると、入れたマクロが保存されます。

同じ名前のモジュールが既にあると、インポートは失敗します。先に「アンインストール」の手順で外してから、入れ直してください。

## アンインストール

1. Outlook を開き、Alt+F11 で VBA エディタを開きます。
2. プロジェクトの中の次を右クリックし、解放します。ファイルに書き出すか聞かれたら、書き出さなくてよいです。
   - `frmSendCheck`
   - `modSendCheck`
   - `modWheelScroll`
   - `ClsButtonHandler`
   - `ClsCardClick`
   - `ClsScrollDrag`
3. 以前の名前が残っていれば、同じように外します。
   - `CButtonHandler`
   - `CCardClick`
   - `CScrollDrag`
4. `ThisOutlookSession` は消しません。`VbaProject.OTM` 自体も消しません。消すと、このマクロ以外のものも一緒に消えます。
5. アドイン タブに「送信前チェックを開く」が残っていれば、ボタンを右クリックして削除します。マクロを外したあとに押すと、呼び先がなくて失敗します。
6. Outlook を普通に終了すると、外した状態が保存されます。

## 動かないときに確認すること

### マクロが無効

ファイル → オプション → トラスト センター → トラスト センターの設定 → マクロの設定 で、マクロを実行できる状態にします。無効のままでは、Alt+F8 に出ても動きません。

新しい Outlook では VBA が動きません。従来の Outlook で開いてください。

### コンパイルで止まる

Alt+F11 で デバッグ → VBAProject のコンパイル を実行します。止まった行が、足りない部品です。次がプロジェクトにあり、名前が一致していることを確認します。

- `frmSendCheck`
- `modSendCheck`
- `modWheelScroll`
- `ClsButtonHandler`
- `ClsCardClick`
- `ClsScrollDrag`

フォームの (オブジェクト名) が `frmSendCheck` でないと、マクロから開けません。コードは `Option Explicit` から入っている必要があります。`CButtonHandler` のように、古い `C` だけの名前が残っていると、新しい `Cls` の名前と混ざってコンパイルに失敗します。古いほうは外します。

### メッセージが出て開かない

「この画面の解像度では、送信前チェックを表示できません」と出たときは、入れ方の失敗ではありません。画面の作業領域に、800 x 460 ポイントが必要です。

「この画面では 1.5 倍では収まりません」と出たときは、表示は 1.0 倍のままです。

### ボタンだけ失敗する

アドイン タブの「送信前チェックを開く」が、マクロを外したあとも残っていると、押したときに呼び先がなくて失敗します。ボタンを右クリックして削除し、Alt+F8 から「送信前チェックを開く」を実行すると、ボタンは付き直します。

### 開き直すと消えている

Outlook を普通に終了していないと、`VbaProject.OTM` に保存されません。タスクマネージャーで終了した場合は、入れた内容が残らないことがあります。
