# 画面モード仕様書 - inputwork.asp（勤務入力画面）

## 1. 画面モード概要

inputwork.aspは、ユーザーの権限とアクセスパラメータに応じて3つのモードで動作します。

| モード | screen値 | 説明 | アクセス条件 |
|--------|---------|------|------------|
| 勤務表入力モード | 0 | 本人が自身の勤務を入力・編集 | デフォルト（本人アクセス） |
| 上長チェックモード | 2 | 上長が部下の勤務を確認・承認 | URLパラメータ`s`あり + 上長権限 |
| 参照モード | 1 | 閲覧のみ（編集不可） | 他者データ参照 or 人事担当者 |

## 2. モード判定ロジック

```vbscript
' inputCommon1.asp での判定処理
screen = 0  ' デフォルト：入力モード

' 他者のデータを参照している場合
If Request.QueryString("p") <> "" And Session("MM_Username") <> Request.QueryString("p") Then
    screen = 1  ' 参照モード
End If

' 人事担当者（charge）の場合
If Request.QueryString("c") <> "" And Session("MM_is_charge") = "1" Then
    screen = 1  ' 参照モード
End If

' 上長チェックの場合（最優先）
If Request.QueryString("s") <> "" And Session("MM_is_superior") = "1" Then
    screen = 2  ' 上長チェックモード
End If
```

## 3. 各モードの詳細仕様

### 3.1 勤務表入力モード（screen = 0）

#### アクセス要件
- `Session("MM_is_input") = "1"` （勤務入力権限あり）
- 自身のデータを表示（`target_personalcode = Session("MM_Username")`）

#### フィールド制御

| フィールドカテゴリ | 編集可否 | 制御変数 |
|------------------|---------|----------|
| 時刻入力欄 | ○（条件付き） | text_disabled |
| 勤務区分選択 | ○（条件付き） | text_disabled |
| 休暇区分選択 | ○（条件付き） | text_disabled |
| メモ欄 | ○（条件付き） | text_disabled |
| 承認チェックボックス | × | text_approval_disabled = "disabled" |
| エラーフラグ | × | text_approval_disabled = "disabled" |

#### 編集可能条件
```vbscript
' 編集不可となる条件
If target_ymb <= processed_ymb Then
    ' 給与処理済み月は編集不可
    text_disabled = "disabled"
ElseIf target_ymb > inputLimitYmb Then
    ' 1ヶ月以上先の月は編集不可
    text_disabled = "disabled"
Else
    text_disabled = ""  ' 編集可能
End If
```

#### ボタン制御
```html
<!-- 出社ボタン：本日のタイムカード打刻がない場合のみ有効 -->
<button id="comeButton" <%=button_disabled%> 
        onclick="timecardUpdate('<%=target_personalcode%>', 1)">
    出社
</button>

<!-- 退社ボタン：本日のタイムカード退社がない場合のみ有効 -->
<button id="leaveButton" <%=button_disabled%> 
        onclick="timecardUpdate('<%=target_personalcode%>', 2)">
    退社
</button>

<!-- 登録ボタン：編集可能期間のみ有効 -->
<button type="submit" <%=button_disabled%>>登録</button>
```

#### フレックスタイム勤務者の追加フィールド
```html
<!-- workshift = "9" の場合のみ表示 -->
<tr class="flex-only">
    <td>勤務開始</td>
    <td><input type="text" name="work_begin" <%=text_disabled%>></td>
</tr>
<tr class="flex-only">
    <td>勤務終了</td>
    <td><input type="text" name="work_end" <%=text_disabled%>></td>
</tr>
<tr class="flex-only">
    <td>休憩開始</td>
    <td><input type="text" name="break_begin1" <%=text_disabled%>></td>
</tr>
<!-- 以下、休憩終了、中抜け開始・終了も同様 -->
```

### 3.2 上長チェックモード（screen = 2）

#### アクセス要件
- URLパラメータ `s` が存在
- `Session("MM_is_superior") = "1"` （上長権限あり）
- URLパラメータ `p` で部下の個人コードを指定

#### フィールド制御

| フィールドカテゴリ | 編集可否 | 制御変数 | 説明 |
|------------------|---------|----------|------|
| 承認チェックボックス | ○ | text_approval_disabled | 上長のみ編集可 |
| エラーフラグ | ○ | text_approval_disabled | 労基法違反の確認 |
| タイムカード時刻 | ○ | text_timecard_disabled | 出社・退社時刻の修正 |
| メモ欄 | △ | - | screen=0の場合のみ編集可 |
| その他全フィールド | × | text_disabled = "disabled" | 閲覧のみ |

#### 特別な処理
```vbscript
' 上長チェックモードでの更新処理
If screen = 2 Then
    ' update_worktbl_is_approval.asp で処理
    ' 更新可能フィールド：
    ' - is_approval（承認フラグ）
    ' - is_error（エラーフラグ）
    ' - beginTime, endTime（タイムカード時刻）
    ' - memo（メモ）※条件付き
    
    ' 更新後はchecklist.aspへリダイレクト
    Response.Redirect("checklist.asp?y=" & target_ymb)
End If
```

#### 承認処理の詳細
```html
<!-- 承認チェックボックス -->
<td class="approval-cell">
    <input type="checkbox" 
           name="apr_<%=workingdate%>" 
           value="1" 
           <%=approval_checked%>
           <%=text_approval_disabled%>>
    承認
</td>

<!-- エラーフラグ表示 -->
<td class="error-cell">
    <% If rs_work("is_error") = "1" Then %>
        <span class="error-mark">⚠️</span>
        <input type="hidden" 
               name="err_<%=workingdate%>" 
               value="1">
    <% End If %>
</td>
```

### 3.3 参照モード（screen = 1）

#### アクセス要件（いずれか）
1. URLパラメータ `p` が存在 かつ `p ≠ Session("MM_Username")`
2. URLパラメータ `c` が存在 かつ `Session("MM_is_charge") = "1"`

#### フィールド制御
| フィールドカテゴリ | 編集可否 | 制御変数 |
|------------------|---------|----------|
| 全フィールド | × | text_disabled = "disabled" |
| 全ボタン | × | button_disabled = "disabled" |

```vbscript
' 参照モードでは全て無効化
text_disabled = "disabled"
text_approval_disabled = "disabled"
text_timecard_disabled = "disabled"
button_disabled = "disabled"
```

## 4. URLパラメータ仕様

| パラメータ | 説明 | 使用例 |
|-----------|------|--------|
| ymb | 表示年月（YYYYMM形式） | ymb=202401 |
| p | 表示対象の個人コード | p=10001 |
| s | 上長チェックモードフラグ | s=1 |
| c | 人事担当者モードフラグ | c=1 |

## 5. アクセスパターン例

### 5.1 本人による勤務入力
```
inputwork.asp?ymb=202401
→ screen = 0（入力モード）
```

### 5.2 上長による部下の勤務承認
```
inputwork.asp?ymb=202401&p=10001&s=1
→ screen = 2（上長チェックモード）
```

### 5.3 人事担当者による確認
```
inputwork.asp?ymb=202401&p=10001&c=1
→ screen = 1（参照モード）
```

### 5.4 一般職員が他者のデータを見ようとした場合
```
inputwork.asp?ymb=202401&p=10002
→ screen = 1（参照モード）※編集不可
```

## 6. セキュリティ考慮事項

1. **権限チェック**
   - 各モードへのアクセス時にセッション変数で権限確認
   - 不正なパラメータでのアクセスは参照モードに降格

2. **データ保護**
   - 給与処理済みデータは全モードで編集不可
   - 承認済みデータは上長のみ変更可能

3. **監査証跡**
   - 承認者と承認日時を記録
   - 変更履歴の保持

## 7. 画面表示制御の実装例

### 7.1 JavaScript側の制御
```javascript
// 画面モードによる表示制御
function initializeScreen(screenMode) {
    switch(screenMode) {
        case 0: // 入力モード
            enableInputMode();
            break;
        case 1: // 参照モード
            enableReferenceMode();
            break;
        case 2: // 上長チェックモード
            enableSuperiorMode();
            break;
    }
}

function enableInputMode() {
    // 入力欄を有効化
    document.querySelectorAll('.time-input').forEach(el => {
        el.removeAttribute('disabled');
    });
    // 承認欄は無効のまま
    document.querySelectorAll('.approval-checkbox').forEach(el => {
        el.setAttribute('disabled', 'disabled');
    });
}

function enableSuperiorMode() {
    // 承認欄のみ有効化
    document.querySelectorAll('.approval-checkbox').forEach(el => {
        el.removeAttribute('disabled');
    });
    // 他の入力欄は無効
    document.querySelectorAll('.time-input').forEach(el => {
        el.setAttribute('disabled', 'disabled');
    });
}

function enableReferenceMode() {
    // 全て無効化
    document.querySelectorAll('input, select, button').forEach(el => {
        el.setAttribute('disabled', 'disabled');
    });
}
```

### 7.2 サーバー側の制御
```vbscript
' フォーム送信時の処理分岐
Select Case screen
    Case 0  ' 入力モード
        ' upsert_worktbl.asp で全フィールド更新
        Server.Execute("inc/upsert_worktbl.asp")
        
    Case 2  ' 上長チェックモード
        ' update_worktbl_is_approval.asp で承認関連のみ更新
        Server.Execute("inc/update_worktbl_is_approval.asp")
        
    Case 1  ' 参照モード
        ' 更新処理なし
        Response.Write("参照モードでは更新できません")
End Select
```

## 8. エラーハンドリング

1. **権限エラー**
   ```vbscript
   If screen = 2 And Session("MM_is_superior") <> "1" Then
       Response.Redirect("sorry.html")
   End If
   ```

2. **期間エラー**
   ```vbscript
   If target_ymb <= processed_ymb Then
       error_message = "給与処理済みのため編集できません"
   End If
   ```

3. **データ不整合**
   ```vbscript
   If screen = 0 And target_personalcode <> Session("MM_Username") Then
       ' 他者のデータを入力モードで開こうとした場合
       screen = 1  ' 強制的に参照モードへ
   End If
   ```