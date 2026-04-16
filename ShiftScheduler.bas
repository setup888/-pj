Attribute VB_Name = "ShiftScheduler"
' =============================================================
' 執務表 一次指定者 自動生成マクロ
'   エントリポイント: GenerateShift
'   依存シート:       執務表 / 当日入力 / 名簿 / 前日実績 / 設定
' =============================================================
Option Explicit

Private Const SHEET_OUT As String = "執務表"
Private Const SHEET_INPUT As String = "当日入力"
Private Const SHEET_ROSTER As String = "名簿"
Private Const SHEET_HISTORY As String = "履歴"
Private Const SHEET_SET As String = "設定"

Private Const ROW_SLOT_START As Long = 5      ' 執務表/設定 の時間枠開始行
Private Const ROW_HIST_START As Long = 5      ' 履歴 のデータ開始行
Private Const N_SLOTS As Long = 25            ' 時間枠数

' 時間枠インデックス定数 (0 origin)
Private Const S_8_9 As Long = 0     ' 8:40〜9
Private Const S_9_10 As Long = 1    ' 9〜10
Private Const S_10_11 As Long = 2
Private Const S_12_13 As Long = 4
Private Const S_14_15 As Long = 6
Private Const S_17_18 As Long = 9
Private Const S_18_19 As Long = 10
Private Const S_20_21 As Long = 12
Private Const S_22_23 As Long = 14
Private Const S_0_1 As Long = 16
Private Const S_4_5 As Long = 20
Private Const S_6_7 As Long = 22
Private Const S_8_840 As Long = 24

Public Sub GenerateShift()
    On Error GoTo EH
    Application.ScreenUpdating = False

    Dim roster As Object:    Set roster = LoadRoster()            ' name -> role
    If roster.Count = 0 Then
        MsgBox "名簿シートが空です。氏名と役職カテゴリを入力してください。", vbExclamation
        GoTo DONE_EXIT
    End If

    Dim daily As Object:     Set daily = LoadDailyInput()
    Dim todayDate As Date
    If Not IsDate(daily("当番日")) Then
        MsgBox "当日入力シートの B3 に当番日 (yyyy/m/d) を入力してください。", vbExclamation
        GoTo DONE_EXIT
    End If
    todayDate = CDate(daily("当番日"))

    Dim prevDay As Object:   Set prevDay = LoadPrevDayFromHistory(todayDate)
    Dim settings As Object:  Set settings = LoadSettings()        ' (slot|role) -> mark

    Dim slots() As String:   slots = ReadTimeSlots()

    ' 当日勤務可能な隊員リスト (休暇・研修を除外)
    Dim members As Variant: members = AvailableMembers(roster, daily)
    If IsEmptyArray(members) Then
        MsgBox "割当可能な隊員が0名です。名簿と当日入力を確認してください。", vbExclamation
        GoTo DONE_EXIT
    End If

    ' 割当結果の格納: 2 枠 (通信=0 / 受付=1) × N_SLOTS
    Dim assign(0 To 1, 0 To N_SLOTS - 1) As String
    Dim workCount As Object: Set workCount = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(members) To UBound(members)
        workCount(CStr(members(i))) = 0
    Next i

    ' メインループ: 各時間枠で 通信 → 受付 の順に割当
    Dim slotIdx As Long, col As Long
    For slotIdx = 0 To N_SLOTS - 1
        For col = 0 To 1  ' 0=通信, 1=受付
            assign(col, slotIdx) = PickAssignee( _
                col, slotIdx, slots, members, roster, _
                daily, prevDay, settings, assign, workCount)
        Next col
    Next slotIdx

    ' 制約修正パス: 12勤 と 17勤 が同じ人なら入れ替え
    Call EnforceDistinct12_17(assign, members, roster, daily, settings, slots)

    ' 出力
    Call WriteOutput(assign, slots, todayDate)

    ' 履歴に追記 (既存同日分は上書き)
    Call AppendToHistory(assign, slots, todayDate)

    ' 検証 & 警告コメント
    Dim warnings As String
    warnings = Validate(assign, roster, daily, prevDay, slots)
    If Len(warnings) > 0 Then
        MsgBox "生成完了。履歴にも追記しました。" & vbCrLf & vbCrLf & _
               "以下の注意点を確認してください:" & vbCrLf & warnings, vbInformation
    Else
        MsgBox "生成完了。履歴にも追記しました。", vbInformation
    End If

DONE_EXIT:
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description & " (行 " & Erl & ")", vbCritical
End Sub

' =============================================================
' データ読み込み
' =============================================================
Private Function LoadRoster() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_ROSTER)
    Dim r As Long, name As String, role As String
    For r = 2 To 31
        name = Trim(CStr(ws.Cells(r, 2).Value))
        role = Trim(CStr(ws.Cells(r, 3).Value))
        If Len(name) > 0 And Left(name, 1) <> "例" Then
            d(name) = role
        End If
    Next r
    Set LoadRoster = d
End Function

Private Function LoadDailyInput() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)
    d("当番日") = ws.Range("B3").Value
    d("曜日") = ws.Range("B4").Value
    d("部") = ws.Range("B5").Value
    d("当直主任") = ws.Range("B6").Value
    d("当直副主任") = ws.Range("B7").Value
    d("休日") = CLng(Nz(ws.Range("B8").Value, 0))

    ' カテゴリリスト (列: 食当=D, 休暇=F, 研修=H, 救助訓練=J, 警防力=L)
    Set d("食当") = ReadNameList(ws, 4)
    Set d("休暇") = ReadNameList(ws, 6)
    Set d("研修") = ReadNameList(ws, 8)
    Set d("救助訓練") = ReadNameList(ws, 10)
    Set d("警防力") = ReadNameList(ws, 12)

    Set LoadDailyInput = d
End Function

Private Function ReadNameList(ws As Worksheet, colIdx As Long) As Object
    ' 当日入力シートのリスト入力欄 (行 12〜19 の固定レイアウト)
    Dim s As Object: Set s = CreateObject("Scripting.Dictionary")
    s.CompareMode = vbTextCompare
    Dim r As Long, name As String
    For r = 12 To 19
        name = Trim(CStr(Nz(ws.Cells(r, colIdx).Value, "")))
        If Len(name) > 0 Then s(name) = True
    Next r
    Set ReadNameList = s
End Function

Private Function LoadPrevDayFromHistory(beforeDate As Date) As Object
    ' 履歴シートから beforeDate より前の最新当番日のレコードを返す
    ' 戻り値: slot_index -> Variant array(0=通信, 1=受付)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    ' 空で初期化
    For i = 0 To N_SLOTS - 1
        d(i) = Array("", "")
    Next i

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_HISTORY)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < ROW_HIST_START Then Set LoadPrevDayFromHistory = d: Exit Function

    ' 前当番日 = beforeDate より小さい最大の日付
    Dim maxDate As Date: maxDate = 0
    Dim r As Long, v As Variant, cellDate As Date
    For r = ROW_HIST_START To lastRow
        v = ws.Cells(r, 1).Value
        If IsDate(v) Then
            cellDate = CDate(v)
            If cellDate < beforeDate And cellDate > maxDate Then
                maxDate = cellDate
            End If
        End If
    Next r
    If maxDate = 0 Then
        Set LoadPrevDayFromHistory = d
        Exit Function
    End If

    ' maxDate の 25 行を読み込む
    Dim slotMap As Object: Set slotMap = BuildSlotLabelMap()
    For r = ROW_HIST_START To lastRow
        v = ws.Cells(r, 1).Value
        If IsDate(v) Then
            If CDate(v) = maxDate Then
                Dim slotLabel As String
                slotLabel = Trim(CStr(Nz(ws.Cells(r, 2).Value, "")))
                If slotMap.Exists(slotLabel) Then
                    Dim idx As Long: idx = CLng(slotMap(slotLabel))
                    d(idx) = Array( _
                        Trim(CStr(Nz(ws.Cells(r, 3).Value, ""))), _
                        Trim(CStr(Nz(ws.Cells(r, 4).Value, ""))) _
                    )
                End If
            End If
        End If
    Next r
    Set LoadPrevDayFromHistory = d
End Function

Private Function BuildSlotLabelMap() As Object
    ' 時間帯ラベル -> index の辞書
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim slots() As String: slots = ReadTimeSlots()
    Dim i As Long
    For i = 0 To N_SLOTS - 1
        d(slots(i)) = i
    Next i
    Set BuildSlotLabelMap = d
End Function

Private Function LoadSettings() As Object
    ' key = slot_index & "|" & role -> mark text
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_SET)
    Dim roles() As String: roles = ReadRoleHeaders(ws)
    Dim i As Long, j As Long, v As String
    For i = 0 To N_SLOTS - 1
        For j = 0 To UBound(roles)
            v = Trim(CStr(ws.Cells(ROW_SLOT_START + i, 2 + j).Value))
            If Len(v) > 0 Then
                d(i & "|" & roles(j)) = v
            End If
        Next j
    Next i
    Set LoadSettings = d
End Function

Private Function ReadRoleHeaders(ws As Worksheet) As String()
    Dim arr(0 To 7) As String
    Dim j As Long
    For j = 0 To 7
        arr(j) = Trim(CStr(ws.Cells(4, 2 + j).Value))
    Next j
    ReadRoleHeaders = arr
End Function

Private Function ReadTimeSlots() As String()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_SET)
    Dim arr(0 To N_SLOTS - 1) As String
    Dim i As Long
    For i = 0 To N_SLOTS - 1
        arr(i) = CStr(ws.Cells(ROW_SLOT_START + i, 1).Value)
    Next i
    ReadTimeSlots = arr
End Function

' =============================================================
' 可用メンバー抽出
' =============================================================
Private Function AvailableMembers(roster As Object, daily As Object) As Variant
    ' 戻り値は Variant 配列。呼び出し側は len = CountAvailable で判定する。
    Dim col As Collection: Set col = New Collection
    Dim k As Variant
    For Each k In roster.Keys
        If Not daily("休暇").Exists(CStr(k)) _
           And Not daily("研修").Exists(CStr(k)) Then
            col.Add CStr(k)
        End If
    Next k
    If col.Count = 0 Then
        AvailableMembers = Array()
        Exit Function
    End If
    Dim arr() As String
    ReDim arr(0 To col.Count - 1)
    Dim i As Long
    For i = 1 To col.Count
        arr(i - 1) = col.Item(i)
    Next i
    AvailableMembers = arr
End Function

Private Function IsEmptyArray(v As Variant) As Boolean
    On Error Resume Next
    Dim u As Long: u = -1
    u = UBound(v)
    On Error GoTo 0
    IsEmptyArray = (u < 0)
End Function

' =============================================================
' 割当ロジック
' =============================================================
Private Function PickAssignee(col As Long, slotIdx As Long, _
    slots() As String, members As Variant, roster As Object, _
    daily As Object, prevDay As Object, settings As Object, _
    assign() As String, workCount As Object) As String

    Dim candidates As Collection: Set candidates = New Collection
    Dim i As Long, name As String, role As String
    For i = LBound(members) To UBound(members)
        name = CStr(members(i))
        If Not IsBlocked(name, slotIdx, roster, daily, settings, slots) Then
            ' 同時間帯で通信と受付は別人
            If col = 1 Then
                If assign(0, slotIdx) = name Then GoTo SKIP_ADD
            End If
            candidates.Add name
        End If
SKIP_ADD:
    Next i

    If candidates.Count = 0 Then
        PickAssignee = ""
        Exit Function
    End If

    ' スコア最小のものを選ぶ (低いほど優先)
    Dim bestName As String: bestName = ""
    Dim bestScore As Double: bestScore = 1E+18
    Dim c As Variant
    For Each c In candidates
        Dim s As Double
        s = ScoreCandidate(CStr(c), col, slotIdx, roster, daily, prevDay, _
                           assign, workCount, members, settings, slots)
        If s < bestScore Then
            bestScore = s
            bestName = CStr(c)
        End If
    Next c

    PickAssignee = bestName
    If Len(bestName) > 0 Then
        workCount(bestName) = workCount(bestName) + 1
    End If
End Function

Private Function IsBlocked(name As String, slotIdx As Long, _
    roster As Object, daily As Object, settings As Object, _
    slots() As String) As Boolean
    Dim role As String: role = CStr(roster(name))

    ' 1) 設定シートの × / △ ( × は絶対禁止 )
    Dim key As String: key = slotIdx & "|" & role
    If settings.Exists(key) Then
        If InStr(settings(key), "×") > 0 Then
            IsBlocked = True
            Exit Function
        End If
    End If

    ' 2) 動的ブロック: 食当 → 14-17時
    If daily("食当").Exists(name) Then
        If slotIdx >= S_14_15 And slotIdx <= S_14_15 + 2 Then  ' 14,15,16
            IsBlocked = True
            Exit Function
        End If
    End If

    ' 3) 動的ブロック: 救助訓練期間 → 10-17時×
    If daily("救助訓練").Exists(name) Then
        If slotIdx >= S_10_11 And slotIdx <= S_17_18 - 1 Then  ' 10..16
            IsBlocked = True
            Exit Function
        End If
    End If

    ' 4) 動的ブロック: 警防力 → 日中×（休日=1なら緩和）、深夜×
    If daily("警防力").Exists(name) Or role = "警防力" Then
        If daily("休日") <> 1 Then
            If slotIdx >= S_8_9 And slotIdx <= S_17_18 Then
                IsBlocked = True
                Exit Function
            End If
        End If
        ' 深夜(0-6時)は基本×
        If slotIdx >= S_0_1 And slotIdx <= S_6_7 - 1 Then
            IsBlocked = True
            Exit Function
        End If
    End If

    ' 5) 当直主任/副主任 は 18-19時, 6-8時 × (当直×相当)
    If name = CStr(Nz(daily("当直主任"), "")) Or _
       name = CStr(Nz(daily("当直副主任"), "")) Then
        If slotIdx = S_18_19 Or slotIdx = S_18_19 + 1 Or _
           slotIdx = S_6_7 Or slotIdx = S_6_7 + 1 Then
            IsBlocked = True
            Exit Function
        End If
    End If

    IsBlocked = False
End Function

Private Function ScoreCandidate(name As String, col As Long, slotIdx As Long, _
    roster As Object, daily As Object, prevDay As Object, _
    assign() As String, workCount As Object, _
    members As Variant, settings As Object, slots() As String) As Double

    Dim score As Double: score = 0

    ' (a) 総勤務回数が少ない人を優先
    score = score + CDbl(workCount(name)) * 10

    ' (b) 前日同時刻と同じ人なら 夜20時以降は強ペナルティ
    Dim prevName As String
    prevName = prevDay(slotIdx)(col)
    If Len(prevName) > 0 And prevName = name Then
        If slotIdx >= S_20_21 Or slotIdx <= S_4_5 Then
            score = score + 1000    ' ほぼ回避
        Else
            score = score + 50      ' 日中は多少ペナルティ
        End If
    End If

    ' (c) 前日実績の 1 個上 (=前日 slotIdx+1 の人) を軽く優先 → 「前日の一個上にずらす」
    If slotIdx + 1 <= N_SLOTS - 1 Then
        Dim prevNext As String: prevNext = prevDay(slotIdx + 1)(col)
        If prevNext = name Then score = score - 5
    End If

    ' (d) 警防力は 18-22時を優先 (それ以外で使われるとペナルティ少し軽減)
    If daily("警防力").Exists(name) Or CStr(roster(name)) = "警防力" Then
        If slotIdx >= S_18_19 And slotIdx <= S_20_21 + 1 Then
            score = score - 20
        End If
    End If

    ' (e) 救助訓練期間者は 17-19時を優先
    If daily("救助訓練").Exists(name) Then
        If slotIdx = S_17_18 Or slotIdx = S_18_19 Then
            score = score - 20
        End If
    End If

    ' (f) 10-17時未勤務者を優先: その人が 10-17時まだ1回も入ってないなら強優先
    If slotIdx >= S_10_11 And slotIdx <= S_17_18 - 1 Then
        If DaytimeCount(name, assign, slotIdx) = 0 Then
            score = score - 30
        End If
    End If

    ' (g) 深夜(22-4時)は一人1回まで: 既に深夜1回ならペナルティ
    If IsLateNight(slotIdx) Then
        If LateNightCount(name, assign, slotIdx) >= 1 Then
            score = score + 500
        End If
    End If

    ' (h) 12勤と17勤は別人: 既に 12勤にこの人入ってたら 17勤で +500
    If slotIdx = S_17_18 Then
        If assign(col, S_12_13) = name Then score = score + 500
    End If
    If slotIdx = S_12_13 Then
        If assign(col, S_17_18) = name Then score = score + 500
    End If

    ' (i) 直前の同じ枠と連続しないように軽いペナルティ
    If slotIdx > 0 Then
        If assign(col, slotIdx - 1) = name Then score = score + 3
    End If

    ScoreCandidate = score
End Function

Private Function IsLateNight(slotIdx As Long) As Boolean
    ' 22-4時 = S_22_23(14) .. S_4_5(20) の 22〜3時台
    IsLateNight = (slotIdx >= S_22_23 And slotIdx <= S_4_5 - 1)
End Function

Private Function DaytimeCount(name As String, assign() As String, _
    currentSlot As Long) As Long
    Dim i As Long, col As Long, cnt As Long
    For i = S_10_11 To S_17_18 - 1
        If i >= currentSlot Then Exit For
        For col = 0 To 1
            If assign(col, i) = name Then cnt = cnt + 1
        Next col
    Next i
    DaytimeCount = cnt
End Function

Private Function LateNightCount(name As String, assign() As String, _
    currentSlot As Long) As Long
    Dim i As Long, col As Long, cnt As Long
    For i = S_22_23 To S_4_5 - 1
        If i >= currentSlot Then Exit For
        For col = 0 To 1
            If assign(col, i) = name Then cnt = cnt + 1
        Next col
    Next i
    LateNightCount = cnt
End Function

' =============================================================
' 後処理: 12勤 と 17勤 の重複解消
' =============================================================
Private Sub EnforceDistinct12_17(assign() As String, _
    members As Variant, roster As Object, daily As Object, _
    settings As Object, slots() As String)
    Dim col As Long
    For col = 0 To 1
        If Len(assign(col, S_12_13)) > 0 And assign(col, S_12_13) = assign(col, S_17_18) Then
            ' 17勤を別の可能な人に差し替える
            Dim i As Long, alt As String
            For i = LBound(members) To UBound(members)
                alt = CStr(members(i))
                If alt <> assign(col, S_12_13) Then
                    If Not IsBlocked(alt, S_17_18, roster, daily, settings, slots) Then
                        If 1 - col = 0 Then
                            If assign(0, S_17_18) <> alt Then
                                assign(col, S_17_18) = alt
                                Exit For
                            End If
                        Else
                            If assign(1 - col, S_17_18) <> alt Then
                                assign(col, S_17_18) = alt
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    Next col
End Sub

' =============================================================
' 出力
' =============================================================
Private Sub WriteOutput(assign() As String, slots() As String, _
    displayDate As Date)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_OUT)
    Dim i As Long
    For i = 0 To N_SLOTS - 1
        ws.Cells(ROW_SLOT_START + i, 2).Value = assign(0, i)
        ws.Cells(ROW_SLOT_START + i, 5).Value = assign(1, i)
    Next i
    ws.Range("B2").Value = displayDate
    ws.Range("B2").NumberFormat = "yyyy/m/d"
    ws.Activate
    ws.Cells(ROW_SLOT_START, 2).Select
End Sub

' =============================================================
' 履歴への追記 (既存同日分は上書き)
' =============================================================
Private Sub AppendToHistory(assign() As String, slots() As String, _
    todayDate As Date)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_HISTORY)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < ROW_HIST_START - 1 Then lastRow = ROW_HIST_START - 1

    ' 既存の同日行を削除 (下から)
    Dim r As Long, v As Variant
    For r = lastRow To ROW_HIST_START Step -1
        v = ws.Cells(r, 1).Value
        If IsDate(v) Then
            If CDate(v) = todayDate Then ws.Rows(r).Delete
        End If
    Next r

    ' 再計算後の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < ROW_HIST_START - 1 Then lastRow = ROW_HIST_START - 1

    ' 追記
    Dim i As Long, writeRow As Long
    For i = 0 To N_SLOTS - 1
        writeRow = lastRow + 1 + i
        With ws.Cells(writeRow, 1)
            .Value = todayDate
            .NumberFormat = "yyyy/m/d"
        End With
        ws.Cells(writeRow, 2).Value = slots(i)
        ws.Cells(writeRow, 3).Value = assign(0, i)
        ws.Cells(writeRow, 4).Value = assign(1, i)
    Next i
End Sub

' =============================================================
' 履歴から過去日を執務表に表示
'   執務表!B2 にある日付を読み、その日のデータをレンダリング
' =============================================================
Public Sub ShowDateFromHistory()
    On Error GoTo EH
    Application.ScreenUpdating = False

    Dim wsOut As Worksheet: Set wsOut = ThisWorkbook.Worksheets(SHEET_OUT)
    Dim dv As Variant: dv = wsOut.Range("B2").Value
    If Not IsDate(dv) Then
        MsgBox "執務表シートの B2 に yyyy/m/d 形式の日付を入力してください。", vbExclamation
        GoTo DONE_EXIT
    End If
    Dim targetDate As Date: targetDate = CDate(dv)

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_HISTORY)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < ROW_HIST_START Then
        MsgBox "履歴が空です。", vbExclamation
        GoTo DONE_EXIT
    End If

    ' 出力欄をクリア
    Dim i As Long
    For i = 0 To N_SLOTS - 1
        wsOut.Cells(ROW_SLOT_START + i, 2).Value = ""
        wsOut.Cells(ROW_SLOT_START + i, 5).Value = ""
    Next i

    Dim slotMap As Object: Set slotMap = BuildSlotLabelMap()
    Dim found As Boolean: found = False
    Dim r As Long, v As Variant
    For r = ROW_HIST_START To lastRow
        v = ws.Cells(r, 1).Value
        If IsDate(v) Then
            If CDate(v) = targetDate Then
                found = True
                Dim slotLabel As String: slotLabel = Trim(CStr(Nz(ws.Cells(r, 2).Value, "")))
                If slotMap.Exists(slotLabel) Then
                    Dim idx As Long: idx = CLng(slotMap(slotLabel))
                    wsOut.Cells(ROW_SLOT_START + idx, 2).Value = _
                        Trim(CStr(Nz(ws.Cells(r, 3).Value, "")))
                    wsOut.Cells(ROW_SLOT_START + idx, 5).Value = _
                        Trim(CStr(Nz(ws.Cells(r, 4).Value, "")))
                End If
            End If
        End If
    Next r

    wsOut.Activate
    If Not found Then
        MsgBox "指定日付のデータが履歴に見つかりません: " & Format(targetDate, "yyyy/m/d"), vbExclamation
    End If

DONE_EXIT:
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub

' =============================================================
' 検証
' =============================================================
Private Function Validate(assign() As String, roster As Object, _
    daily As Object, prevDay As Object, slots() As String) As String
    Dim msgs As String, i As Long, col As Long

    ' 空欄チェック
    For i = 0 To N_SLOTS - 1
        For col = 0 To 1
            If Len(assign(col, i)) = 0 Then
                msgs = msgs & " - " & slots(i) & " / " & IIf(col = 0, "通信", "受付") & " が空欄 (割当候補なし)" & vbCrLf
            End If
        Next col
    Next i

    ' 10-17時 全員カバー
    Dim k As Variant, covered As Object
    Set covered = CreateObject("Scripting.Dictionary")
    covered.CompareMode = vbTextCompare
    For Each k In roster.Keys
        covered(CStr(k)) = 0
    Next k
    For i = S_10_11 To S_17_18 - 1
        For col = 0 To 1
            If Len(assign(col, i)) > 0 And covered.Exists(assign(col, i)) Then
                covered(assign(col, i)) = covered(assign(col, i)) + 1
            End If
        Next col
    Next i
    For Each k In roster.Keys
        If covered(CStr(k)) = 0 _
            And Not daily("休暇").Exists(CStr(k)) _
            And Not daily("研修").Exists(CStr(k)) _
            And CStr(roster(k)) <> "警防力" _
            And Not daily("警防力").Exists(CStr(k)) Then
            msgs = msgs & " - " & CStr(k) & " は 10-17時に未勤務" & vbCrLf
        End If
    Next k

    ' 夜20時以降 前日同一人物チェック
    For i = S_20_21 To N_SLOTS - 1
        For col = 0 To 1
            If Len(assign(col, i)) > 0 And assign(col, i) = prevDay(i)(col) Then
                msgs = msgs & " - " & slots(i) & " / " & IIf(col = 0, "通信", "受付") & _
                       " が前日と同じ (" & assign(col, i) & ")" & vbCrLf
            End If
        Next col
    Next i

    Validate = msgs
End Function

' =============================================================
' ユーティリティ
' =============================================================
Private Function Nz(v As Variant, alt As Variant) As Variant
    ' VBA は Or に短絡評価が無いので段階チェック
    If IsNull(v) Then
        Nz = alt
    ElseIf IsEmpty(v) Then
        Nz = alt
    ElseIf VarType(v) = vbString Then
        If Len(CStr(v)) = 0 Then Nz = alt Else Nz = v
    Else
        Nz = v
    End If
End Function
