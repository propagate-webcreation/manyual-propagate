# スプレッドシートとGASの読み方ガイド

Claude Code からスプレッドシートやGASにアクセスして読み書きする方法をまとめたガイドです。

---

# 1. スプレッドシートの読み書き（Google Sheets API）

## 認証

まずアクセストークンを取得します。

```bash
TOKEN=$(gcloud auth print-access-token)
```

このトークンは約1時間で期限切れになります。切れたら再度実行してください。

## 読み取り（セルの値を取得）

```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://sheets.googleapis.com/v4/spreadsheets/{シートID}/values/{範囲}?valueRenderOption=FORMATTED_VALUE"
```

### パラメータ説明

| パラメータ | 説明 | 例 |
|-----------|------|---|
| `{シートID}` | スプレッドシートのURLに含まれる44文字のID | `1OfNcHbWvRIYyMEy9bSU4-aCduL0rmdw_fVQla9Fa8Dc` |
| `{範囲}` | 「タブ名!セル範囲」の形式 | `メイン!B3:C` |
| `valueRenderOption` | 値の返し方 | 下記参照 |

### valueRenderOption の種類

| 値 | 意味 | 使い所 |
|---|------|--------|
| `FORMATTED_VALUE` | 画面に表示されている値 | 通常はこれを使う |
| `FORMULA` | セルに入っている数式 | 数式の確認・デバッグ |
| `UNFORMATTED_VALUE` | 生の値（日付はシリアル値） | 日付型の確認（時刻混入の検出等） |

### 具体例

**メインタブのB列（record_id）を全行取得:**
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://sheets.googleapis.com/v4/spreadsheets/1OfNcHbWvRIYyMEy9bSU4-aCduL0rmdw_fVQla9Fa8Dc/values/%E3%83%A1%E3%82%A4%E3%83%B3!B3:B"
```

> ⚠️ **日本語タブ名はURLエンコードが必要。** 「メイン」→ `%E3%83%A1%E3%82%A4%E3%83%B3`

**Pythonでエンコードする方法:**
```python
import urllib.parse
encoded = urllib.parse.quote('メイン!B3:C')
# → '%E3%83%A1%E3%82%A4%E3%83%B3%21B3%3AC'
```

### レスポンス例

```json
{
  "range": "'メイン'!B3:C656",
  "majorDimension": "ROWS",
  "values": [
    ["16786", "平和企画"],
    ["12104", "HTM合同会社"],
    ...
  ]
}
```

`values` が2次元配列で返ってきます。各行が1つの配列です。

## 書き込み

```bash
curl -s -X PUT -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
  "https://sheets.googleapis.com/v4/spreadsheets/{シートID}/values/{範囲}?valueInputOption=USER_ENTERED" \
  -d '{"values": [["値1", "値2"]]}'
```

### valueInputOption の種類

| 値 | 意味 |
|---|------|
| `USER_ENTERED` | ユーザーが手入力したのと同じ扱い（数式も解釈、日付も自動認識） |
| `RAW` | そのままテキストとして書き込み |

通常は `USER_ENTERED` を使います。

### 数式を書き込む例

```bash
curl -s -X PUT -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
  "https://sheets.googleapis.com/v4/spreadsheets/{シートID}/values/%E3%83%A1%E3%82%A4%E3%83%B3!Q3?valueInputOption=USER_ENTERED" \
  -d '{"values": [["=IF(R3=\"確認待ち\",\"ディレクター\",\"\")"]]}'
```

## Pythonでの使い方（まとめ）

```python
import json, urllib.request, urllib.parse, os

TOKEN = os.popen('gcloud auth print-access-token').read().strip()
SHEET_ID = '1OfNcHbWvRIYyMEy9bSU4-aCduL0rmdw_fVQla9Fa8Dc'

def fetch(range_str, render='FORMATTED_VALUE'):
    """スプレッドシートからデータを読み取る"""
    encoded = urllib.parse.quote(range_str)
    url = f'https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{encoded}?valueRenderOption={render}'
    req = urllib.request.Request(url, headers={'Authorization': f'Bearer {TOKEN}'})
    resp = urllib.request.urlopen(req)
    return json.loads(resp.read().decode('utf-8')).get('values', [])

def write(range_str, values):
    """スプレッドシートにデータを書き込む"""
    encoded = urllib.parse.quote(range_str)
    url = f'https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{encoded}?valueInputOption=USER_ENTERED'
    body = json.dumps({'values': values}).encode('utf-8')
    req = urllib.request.Request(url, data=body, method='PUT', headers={
        'Authorization': f'Bearer {TOKEN}',
        'Content-Type': 'application/json'
    })
    return json.loads(urllib.request.urlopen(req).read().decode('utf-8'))

# 使用例
rows = fetch('メイン!B3:C')       # メインタブのB-C列を全行取得
formulas = fetch('メイン!Q3', 'FORMULA')  # Q3の数式を取得
write('メイン!T100', [['2026/04/20']])    # T100に日付を書き込み
```

---

# 2. GASコードの読み書き（clasp）

## claspとは

Google Apps Script（GAS）をローカルで編集するためのCLIツール。ブラウザのGASエディタを使わず、ローカルのエディタで編集→pushする運用。

> ⚠️ **GASエディタのブラウザ操作は禁止**（Monaco API誤上書き事故あり）。必ずclasp経由。

## 基本コマンド

### インストール
```bash
npm install -g @google/clasp
clasp login  # ブラウザでGoogle認証
```

### プロジェクトをローカルに取得
```bash
mkdir gas && cd gas
clasp clone {GASプロジェクトID}
```

### リモートの最新をローカルに取得
```bash
clasp pull
```

### ローカルの変更をリモートに反映
```bash
# push前に必ずファイル一覧を確認（意図しないファイル削除を防ぐ）
ls *.js *.json
clasp push
```

> ⚠️ **`clasp push` はローカルにないファイルをリモートから削除する。** push前に `clasp pull` で最新化し、ファイル一覧を確認すること。

## GASコードの読み方

### ファイル構成

GASプロジェクトは複数の `.js`（または `.gs`）ファイルで構成。全ファイルが1つのグローバルスコープを共有しているので、あるファイルの関数を別ファイルから呼び出せる。

### トリガーの種類

| 種類 | 説明 | 例 |
|------|------|---|
| onEdit | スプレッドシートのセルが手動編集された時 | `onEditAll(e)` |
| 時間ベース | 定期実行（毎分、毎時、毎日等） | `fillMissingPolishTimestamps()` 毎1分 |
| 手動実行 | GASエディタから手動で実行 | `setupDailyAutoMoveTrigger()` |

### onEditイベントオブジェクト（e）

onEdit関数が受け取る `e` の中身：

```javascript
e.range          // 編集されたセル
e.range.getRow()     // 行番号
e.range.getColumn()  // 列番号
e.range.getValue()   // 新しい値
e.oldValue       // 変更前の値（単一セル編集時のみ）
e.source         // スプレッドシート全体
e.source.getActiveSheet()  // 編集されたシート
```

### よく使うパターン

**ヘッダーから列番号を取得（ハードコード禁止）:**
```javascript
var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
// headers = ['record_id', 'client_name', 'plan_type', ...]

// NG: var COL = 17;  ← 列番号ハードコード禁止
// OK:
function getColumnIndexByHeader(headers, name) {
  var clean = headers.map(function(h) {
    return h.toString().replace(/\s/g, '').replace(/\n/g, '');
  });
  return clean.indexOf(name) + 1;  // 1-based
}
var qCol = getColumnIndexByHeader(headers, 'ball_holder');  // → 17
```

**セルの値を取得:**
```javascript
var value = sheet.getRange(row, column).getValue();
```

**セルに値を書き込み:**
```javascript
sheet.getRange(row, column).setValue('新しい値');
```

**行を追加（末尾）:**
```javascript
sheet.appendRow(['値1', '値2', '値3', ...]);
```

**行を削除:**
```javascript
sheet.deleteRow(rowNumber);
```

> ⚠️ ループ内でdeleteRowすると行番号がズレる。下から順に削除すること。

### トリガーの確認方法

GASエディタ（ブラウザ）で確認する場合：
1. https://script.google.com を開く
2. 対象プロジェクトを選択
3. 左メニュー「トリガー」（時計アイコン）をクリック
4. 登録されているトリガー一覧が表示される

ローカルからの確認：
```bash
# GAS_TRIGGERS.md に全トリガー一覧を記載済み
cat GAS_TRIGGERS.md
```

---

# 3. よくあるトラブルと対処

## 認証エラー（401 Unauthorized）

```
HTTP Error 401: Unauthorized
```

トークンの期限切れ。再取得：
```bash
TOKEN=$(gcloud auth print-access-token)
```

## 日本語タブ名のエラー

```
Invalid range
```

タブ名をURLエンコードしていない。Pythonで：
```python
urllib.parse.quote('メイン!B3:C')
```

## clasp pushでファイルが消えた

`clasp push` はローカルにないファイルをリモートから削除する。

対処：
1. `clasp pull` で最新を取得
2. ローカルで編集
3. `ls` でファイル一覧確認
4. `clasp push`

## ARRAYFORMULA列に値を書き込んでしまった

AA列（client_email）、AB列（line_display_name）、Q/U/V/W列はARRAYFORMULA管理。直接値を書き込むと `#REF!` エラーになる。

対処：書き込んだセルの値をクリア（clearContent）すればARRAYFORMULAが再展開される。

---

# 4. 制作進行シート固有の情報

## シートID一覧

| 用途 | ID |
|------|-----|
| 本番シート | `1OfNcHbWvRIYyMEy9bSU4-aCduL0rmdw_fVQla9Fa8Dc` |
| 旧原本（参照のみ） | `{OLD_SHEET_ID}` |
| 出勤カレンダー | `{ATTENDANCE_SHEET_ID}` |
| 既存サイト修正依頼 | `{EXISTING_FIX_SHEET_ID}` |

## GASプロジェクトID

`1-DSOzpTqCMQ-eCi76ZQxjh4QuP-c3BsAgEhyog8NtCsuoFllmctpHqP8`

## ARRAYFORMULA管理列（直接値を書き込まない）

| 列 | 内容 | 数式の場所 |
|---|------|----------|
| Q | ボール保持者 | Row2 ARRAYFORMULA |
| U | 提出警告日 | Row2 BYROW |
| V | 遅延フラグ | Row2 BYROW |
| W | 遅延日数 | Row2 ARRAYFORMULA |
| AA | メールアドレス | Row2 ARRAYFORMULA |
| AB | LINE表示名 | Row2 ARRAYFORMULA |

## 保護列（編集禁止）

A-H列（Kintone連携）、L列（未払い）、S列（GAS自動）はシート保護されています。
