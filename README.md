# LRN NW Collector（ライフリズムナビ NW情報収集ツール）

介護施設のネットワーク情報を**ダブルクリックだけ**で収集するツールです。
主にライフリズムナビ（LRN）の導入前のデモ機のネットワーク調査で使用します。

## 使い方

1. `NW_Collect.bat` と `NW_Collect.ps1` を**同じフォルダ**に置く
2. `NW_Collect.bat` を**ダブルクリック**
3. ブラウザにレポートが自動で開きます

## 自動取得される情報

| 項目 | 取得元 |
|------|--------|
| SSID（ネットワーク名） | `netsh wlan show interfaces` |
| 認証方式・暗号化方式 | `netsh wlan show interfaces` |
| 電波強度・無線規格・チャネル | `netsh wlan show interfaces` |
| BSSID（APのMACアドレス） | `netsh wlan show interfaces` |
| PCのMACアドレス | `Get-NetAdapter` |
| IPアドレス・サブネットマスク | `Get-NetIPConfiguration` |
| デフォルトゲートウェイ | `Get-NetIPConfiguration` |
| DNSサーバー（プライマリ/セカンダリ） | `Get-NetIPConfiguration` |
| DHCP状態 | `Get-NetIPInterface` |
| インターネット接続状態 | `Test-Connection` |

## 手動入力項目（ブラウザ上のフォームで入力）

- Wi-Fiパスワード
- 周波数帯（5GHz推奨）
- プライバシーセパレータ
- MACアドレス認証
- 固定IPの払い出し可否・利用可能IP
- UTM / ファイアウォールの有無
- プロキシ環境
- キャプティブポータルの有無
- VLAN分割の有無
- 施設名・担当者名・調査場所

## レポートの機能

- **テキストコピー** — メールやチャットに貼り付け可能
- **印刷 / PDF保存** — 紙やPDFで保管（自動取得/手動入力/施設情報がページ別に分かれます）
- **保存（HTML）** — 手動入力内容を含めてHTMLファイルとして保存

## 動作要件

- Windows 10 以降
- PowerShell 5.1 以降（Windows標準搭載）
- Wi-Fi接続中のPC

## ファイル構成

```
NW_Collect.bat   … ダブルクリック用エントリポイント（英語のみ・文字化け防止）
NW_Collect.ps1   … 実際の処理スクリプト（UTF-8 BOM付き）
README.md        … このファイル
```

## ライセンス

MIT License
