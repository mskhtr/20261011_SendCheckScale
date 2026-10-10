# modAttachmentCheck

添付ファイル 1 件の中身を調べる VBA モジュールです。送信前チェックの画面には、まだ組み込んでいません。

拡張子は見ません。ファイルが ZIP かどうかは中身の署名で判定し、アーカイブ内のファイル名から拡張子を集めます。`.zi_` や `._IP` のように拡張子が変えられていても、中身が ZIP なら読めます。

## 読み込み

Excel の VBA エディタで **ファイル → ファイルのインポート** から `modAttachmentCheck.bas` を読み込みます。

この `.bas` は **Shift_JIS** です。UTF-8 のままインポートすると、日本語が文字化けします。すでに化けた状態で入っているときは、モジュールを削除してから読み込み直してください。

Cursor でこの `.bas` が化けて見えるときは、右下の文字コードから **Japanese (Shift JIS)** で開き直してください。

## 使い方

イミディエイト ウィンドウ（`Ctrl + G`）で、先頭に `?` を付けて実行します。

```vb
?InspectAttachmentFile("C:\temp\見積._IP")
```

拡張子の一覧だけ欲しいときは、こちらです。

```vb
?AttachmentExtensions("C:\temp\見積._IP")
```

`InspectAttachmentFile` は、次の形の複数行テキストを返します。

```text
file: C:\temp\資料.zip
isZip: True
outerExtension: .zip
entryCount: 3
extensions: .pdf
entries:
cover.pdf	.pdf
uchiwake.zip	.zip	container
uchiwake.zip/見積.pdf	.pdf
```

| 行 | 意味 |
| --- | --- |
| `isZip` | 中身が ZIP なら `True`。拡張子が `.zip` でも、中身が ZIP でなければ `False` |
| `outerExtension` | 渡したファイル自身の拡張子。小文字。無ければ `(none)` |
| `entryCount` | `entries` に並ぶファイル数。フォルダ行は含めない |
| `extensions` | 中のファイルの拡張子。中まで読めた ZIP は入れない |
| `entries` | 中のパス、タブ、拡張子。器の ZIP は末尾が `container`。暗号化フラグがあるときは `encrypted` |
| `errors` | 読めなかったときだけ付く |

ZIP でないファイルを渡したときは、そのファイル自身を 1 件として返します。`extensions` はそのファイルの拡張子です。

`AttachmentExtensions` は、上の `extensions` 行の値だけを返します。ZIP でなければ、そのファイル自身の拡張子です。

## 中をどこまで見るか

ZIP の一覧は、展開せずにセントラルディレクトリのファイル名から作ります。

入れ子の ZIP は、次のときに中も見ます。

- 無圧縮（stored）で、中身の先頭が ZIP 署名である
- 圧縮（deflate）で、ファイル名が ZIP らしい。対象は `zip` `zipx` `zip_` `zi_` `z_p` `_ip`

`hidden.bin` のように名前が ZIP でなくても、無圧縮で中身が ZIP なら開きます。圧縮されていて、名前が上のどれでもないものは、そのファイルの拡張子だけを返して中は開きません。

中まで読めた ZIP は器です。`entries` には `container` 付きで残しますが、`extensions` からは外します。`資料.zip` の中の `内訳.zip` の中身が PDF だけなら、`extensions` は `.pdf` だけです。中を読めなかった ZIP は器にしないので、その拡張子は `extensions` に残ります。

パスワードで中身が暗号化されていても、ファイル名まで暗号化されていなければ拡張子は取れます。その行には `encrypted` が付き、中の展開はしません。

ファイル名の文字コードは、ZIP の UTF-8 フラグがあれば UTF-8、無ければ Windows の ANSI（日本語環境では Shift_JIS）です。

`.xlsx` `.docx` `.pptx` も実体は ZIP です。渡すと、中の `.xml` などが拡張子として出ます。

## 制限

| 項目 | 上限 |
| --- | --- |
| 調べるファイル | 32MB |
| エントリ数 | 2000 件 |
| 入れ子 | 4 段 |
| 展開が必要な入れ子 1 件 | 8MB |
| 展開回数 | 30 回 |

分割 ZIP と ZIP64 には対応していません。フォルダを渡すとエラーになります。

## 動作確認

同じフォルダの `modAttachmentCheckTest.bas` を一緒にインポートし、イミディエイトで実行します。

```vb
?RunAttachmentCheckTests("c:\Users\masaki\Documents\python-projects\_tool\20261008_SendCheck\test")
```

結果は `test\result.txt` に出ます。テストが作る ZIP は `test\fixtures` に残ります。拡張子を変えたサンプルは `test\fixtures\stored._IP` です。
