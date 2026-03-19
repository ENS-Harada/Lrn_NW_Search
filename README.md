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
| **IPアドレススキャン（使用中/空き）** | **並列Ping + ARPテーブル** |

## IPアドレススキャン機能（v1.1追加）

ネットワーク内の全IPアドレスをスキャンし、**使用中（赤）** と **空き（緑）** を一覧表示します。
カメラの固定IPアドレス設定時に「空いているはずのIPが実は使われていた」問題を事前に防ぎます。

- PCのIPアドレスとサブネットマスクからネットワーク範囲を自動計算
- Runspace Poolによる並列Ping（最大100スレッド同時実行）
- ARPテーブルも参照し、Ping非応答デバイスも検出
- /24（約254ホスト）：約1分 ／ /16（約65,534ホスト）：30分以上
- **注意:** Ping（ICMP）がブロックされている端末は「空き」と表示される場合があります

## 手動入力項目（ブラウザ上のフォームで入力 ※ブランクでも問題ない）

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
NW_Collect.bat           … ダブルクリック用エントリポイント（英語のみ・文字化け防止）
NW_Collect.ps1           … 実際の処理スクリプト（UTF-8 BOM付き）
NW_Collect_マニュアル.docx … 営業向け操作マニュアル
README.md                … このファイル
```

## 更新履歴

- **v1.1** — IPアドレススキャン機能を追加（並列Ping + ARP補完）
- **v1.0** — 初版リリース（Wi-Fi/IP/インターネット自動収集）

## ライセンス

MIT License
