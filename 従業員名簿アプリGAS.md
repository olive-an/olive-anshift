# 従業員名簿アプリ（roster.html）の設定

## 全体の形

| 部品 | 場所 | 中身 |
| --- | --- | --- |
| アプリの画面 | このリポジトリの `roster.html` | 画面だけ。名簿のデータは入っていません |
| 名簿のデータ | Googleスプレッドシート（非公開） | 今のExcel名簿をそのままアップロードしたもの |
| 受け渡し役 | スプレッドシートに付けるGAS | 合言葉が合ったときだけ名簿を返す |

このリポジトリは公開されています。名簿をこのリポジトリに入れると、住所や電話番号が誰でも見られる状態になります。そのためデータは必ずスプレッドシート側に置きます。

## 手順

### 1. Excelをスプレッドシートにする

1. Googleドライブにアップロードしたいxlsxを入れます。
2. 右クリックして「アプリで開く」の「Googleスプレッドシート」を選びます。
3. 開いたあと「ファイル」の「Googleスプレッドシートとして保存」をします。
4. 共有設定は「制限付き」のままにします（リンクを知っている全員にはしない）。

シート名は次の2つを今のまま使います。名前を変えるとアプリが読めません。

- `従業員名簿（おりーぶ庵）`
- `おりーぶ庵.退職者`
- `従業員ぼのぼの`（ヘルパーステーションぼのぼの。氏名・性別・生年月日・住所のみ）

### 2. GASを付ける

1. そのスプレッドシートで「拡張機能」の「Apps Script」を開きます。
2. 下のコードを貼り付けて保存します。
3. 左の歯車（プロジェクトの設定）の「スクリプト プロパティ」に、次を追加します。

| プロパティ | 値 |
| --- | --- |
| `ROSTER_PASS` | 管理者で決めた合言葉（8文字以上を推奨） |

4. 右上の「デプロイ」の「新しいデプロイ」を選びます。種類は「ウェブアプリ」、実行ユーザーは「自分」、アクセスできるユーザーは「全員」にします。
5. 出てきたURL（`…/exec` で終わるもの）をコピーします。

「全員」にしても、合言葉が合わなければ名簿は返りません。

### 3. アプリにURLを入れる

`roster.html` の中ほどにある次の行を、コピーしたURLに書き換えます。

```javascript
const GAS_URL = "YOUR_ROSTER_GAS_URL";
```

### 4. スマホに追加する

`https://olive-an.github.io/olive-anshift/roster.html` を開き、ブラウザのメニューから「ホーム画面に追加」します。

## GASのコード

```javascript
// 従業員名簿アプリ用。スプレッドシートに付けて使う（閲覧専用）。
// 個人番号は、この一覧に載せていないので、アプリには絶対に届かない。

const ACTIVE_SHEET = "従業員名簿（おりーぶ庵）";
const LEFT_SHEET = "おりーぶ庵.退職者";
const BONO_SHEET = "従業員ぼのぼの";

// アプリに渡す項目。左がアプリ側の名前、右が見出しの候補（改行や空白は無視して照合）
const ACTIVE_COLS = {
  no: ["列1", "No."], name: ["社員氏名"], kana: ["社員氏名カナ"], sex: ["性別"],
  birth: ["生年月日"], role: ["役職"], group: ["類人猿分類"], type: ["常勤・非常勤"],
  site: ["拠点"], zip: ["郵便番号"], addr: ["住所0", "住所1", "住所"],
  tel: ["電話番号"], mobile: ["携帯電話"], mail: ["メールアドレス"],
  joined: ["入社年月日"], kenpo: ["けんぽ対象者"], resume: ["履歴書・資格証提出"],
  line: ["個人LINEで聞き取り"], dw: ["ダブルワーク有無"]
};
const LEFT_COLS = {
  no: ["No.", "列1"], name: ["社員氏名"], sex: ["性別"], birth: ["生年月日"],
  role: ["役職"], type: ["常勤・非常勤"], zip: ["郵便番号"], addr: ["住所1", "住所0", "住所"],
  tel: ["電話番号"], mobile: ["携帯電話"], joined: ["入社年月日"], left: ["退社年月日"]
};

// ぼのぼのシートは項目が少ない（見出し行は3行目。右側の別の表は見出しがないので読まれない）
const BONO_COLS = {
  no: ["No."], name: ["社員氏名"], sex: ["性別"], birth: ["生年月日"],
  zip: ["郵便番号"], addr: ["住所1", "住所0", "住所"]
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = /^[A-Za-z0-9_]{1,80}$/.test(p.callback || "") ? p.callback : "";
  const out = (obj) => {
    const json = JSON.stringify(obj);
    return ContentService
      .createTextOutput(cb ? cb + "(" + json + ");" : json)
      .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  };

  const pass = PropertiesService.getScriptProperties().getProperty("ROSTER_PASS");
  if (!pass) return out({ ok: false, error: "GASに合言葉（ROSTER_PASS）が設定されていません。" });
  if (p.action !== "roster" || p.pass !== pass) {
    Utilities.sleep(1500); // 当てずっぽうの連打を遅くする
    return out({ ok: false, error: "合言葉が違います。" });
  }

  const ss = SpreadsheetApp.getActive();
  return out({
    ok: true,
    active: readSheet_(ss, ACTIVE_SHEET, ACTIVE_COLS),
    left: readSheet_(ss, LEFT_SHEET, LEFT_COLS),
    bono: readSheet_(ss, BONO_SHEET, BONO_COLS)
  });
}

function norm_(s) { return String(s == null ? "" : s).replace(/\s+/g, ""); }

function readSheet_(ss, sheetName, cols) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const raw = sheet.getDataRange().getValues();

  // 見出し行は「社員氏名」が入っている最初の行（退職者シートは先頭に空行がある）
  let h = -1;
  for (let i = 0; i < Math.min(values.length, 15); i++) {
    if (values[i].some((v) => norm_(v) === "社員氏名")) { h = i; break; }
  }
  if (h < 0) return [];

  const index = {};
  const heads = values[h].map(norm_);
  Object.keys(cols).forEach((key) => {
    for (const cand of cols[key]) {
      const at = heads.indexOf(norm_(cand));
      if (at >= 0) { index[key] = at; break; }
    }
  });

  const tz = ss.getSpreadsheetTimeZone();
  const rows = [];
  for (let r = h + 1; r < raw.length; r++) {
    const nameAt = index.name;
    if (nameAt == null || !norm_(raw[r][nameAt])) continue;
    const rec = {};
    Object.keys(index).forEach((key) => {
      const v = raw[r][index[key]];
      if (v instanceof Date) rec[key] = Utilities.formatDate(v, tz, "yyyy-MM-dd");
      else rec[key] = v == null ? "" : String(v).trim();
    });
    rows.push(rec);
  }
  return rows;
}
```

## 動作の確認

1. GASの画面でURLを開く場合は、末尾に `?action=roster&pass=合言葉` を付けます。`{"ok":true,...}` が出れば成功です。
2. 返ってきた中に個人番号（12桁の番号）が出ていないことを確認してください（出ない作りです）。
3. `roster.html` を開き、合言葉を入れて名簿が出ることを確認します。

## 運用のメモ

- 名簿を直すのは、今までどおりスプレッドシートで行います。アプリは見るだけです。アプリの右上「更新」で最新になります。
- 合言葉を変えるときは、GASのスクリプト プロパティ `ROSTER_PASS` を書き換えるだけです。コードの再デプロイは要りません。
- 合言葉はURLの一部として送られます。GoogleのログにURLが残る可能性があるため、他のサービスと同じ合言葉にしないでください。
- 年齢と在籍期間は、アプリが生年月日・入社年月日から毎回計算します（シートの数式は使っていません）。
- 見出しと中身がずれている列があります。「履歴書・資格証提出」「個人LINEで聞き取り」には資格名が、「ダブルワーク有無」には〇が入っています。アプリは見出しの名前のまま表示します。見出しを直す場合は、シートとGASの `ACTIVE_COLS` の両方をそろえてください。
- 退職者シートの「年齢」「在籍期間」は数式で、アプリでは使いません。
- ぼのぼのシートの右側（J列から右）にある別の表は、退職者らしき一覧で、見出しがないため読み込みません。ぼのぼのシートの更新日は2021年1月25日のままです。最新に直してから使ってください。
