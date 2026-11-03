# 個人用家計簿

LIFF画面とGoogle Apps Script（GAS）を同じリポジトリで管理します。

## 構成

| パス | 用途 |
| --- | --- |
| `index.html` | LIFF画面。既存の公開パスを維持 |
| `src/コード.js` | LINE登録、シート整理、投信管理のGAS |
| `src/appsscript.json` | GASマニフェスト |
| `.clasp.example.json` | GAS接続設定のひな形 |

## ローカル設定

1. `npm ci` で開発用ツールをインストールします。
2. `.clasp.example.json` を `.clasp.json` にコピーし、対象GASのスクリプトIDを設定します。
3. `npm run login` でGoogleにログインします。
4. GASのスクリプトプロパティに次の値を設定します。

| キー | 内容 |
| --- | --- |
| `ACCESS_TOKEN` | LINEチャネルアクセストークン |
| `SHEET_ID` | 家計簿のスプレッドシートID |
| `FUND_LINE_TO` | LINE通知先。通知しない場合は空文字 |
| `FUND_SPREADSHEET_ID` | 投信用シートID。家計簿と共通なら空文字 |

認証ファイル `.clasprc.json` とローカル接続設定 `.clasp.json` はGit管理の対象外です。

## 更新方法

- LIFF：`index.html` を更新し、既存のWebホスティングの方法で公開します。
- GAS：`npm run status` で対象を確認し、`npm run push` で `src/` のソースを反映します。
- GASの公開Webアプリは、ソースのpushとは別に、使用中の既存デプロイを更新します。
  新規デプロイを作らず既存のデプロイIDを維持すると、URLを維持できます。

GitHubへのpushとGASへのpushは別操作です。`.claspignore` により、LIFFのHTMLはGASへ送信しません。

## LIFFとGASの接続

LIFFの費目取得先とLINEのWebhookは、使用中のGAS Webアプリを参照します。
このブランチでは、削除された旧デプロイ51の費目取得URLを、現在使用中の52のURLに修正しています。
LIFF側の変更は、Webホスティングへの公開後に有効になります。

GASの `FUNDS`、`FUND_ACCOUNT_POSITIONS`、`FUND_MONTHLY_ACCOUNT` には個別の投信設定が含まれます。
公開リポジトリへ反映する際は、これらの情報も公開対象になります。
