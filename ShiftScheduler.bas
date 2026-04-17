Attribute VB_Name = "ShiftScheduler"
' =============================================================
' 執務表 一次指定者 自動生成マクロ (完全データ駆動版)
'   エントリポイント: GenerateShift / ShowDateFromHistory
'
' ポジション×時間帯のブロック設定はすべて「ポジション定義」シートで管理。
' VBAには業務ルールのハードコードなし。
' =============================================================
Option Explicit

Private Const SHEET_OUT As String = "執務表"
Private Const SHEET_INPUT As String = "当日チェック"
Private Const SHEET_POS As String = "ポジション定義"
Private Const SHEET_ROSTER As String = "名簿"
Private Const SHEET_HISTORY As String = "履歴"

Private Const ROW_SLOT_START As Long = 5
Private Const ROW_POS_START As Long = 5          ' ポジション定義 データ開始行
Private Const ROW_POS_END As Long = 54           ' 〃 (50行バッファ)
Private Const ROW_CHECK_START As Long = 7
Private Const ROW_CHECK_END As Long = 36
Private Const ROW_HIST_START As Long = 5
Private Const N_SLOTS As Long = 25

' スコアリング用の時間帯定数 (0 origin)
Private Const S_10_11 As Long = 2
Private Const S_12_13 As Long = 4
Private Const S_17_18 As Long = 9
Private Const S_18_19 As Long = 10
Private Const S_20_21 As Long = 12
Private Const S_22_23 As Long = 14
Private Const S_4_5 As Long = 20

' =============================================================
' メインエントリ: 執務表を生成して履歴に追記
' =============================================================
Public Sub GenerateShift()
    On Error GoTo EH
    Application.ScreenUpdating = False

    Dim slots() As String: slots = ReadTimeSlots()

    Dim roster As Object: Set roster = LoadRoster()
    If roster.Count = 0 Then
        MsgBox "名簿シートが空です。氏名を入力してください。", vbExclamation
        GoTo DONE_EXIT
    End If

    Dim positions As Object: Set positions = LoadPositions(slots)
    ' positions("ポジション名") -> Dictionary(slotIdx -> True) (ブロックセット)

    Dim daily As Object: Set daily = LoadDailyInput(slots)
    ' daily("当番日"), daily("休日"),
    ' daily("ポジション") -> Dict: name -> position_name
    ' daily("除外")       -> Dict: name -> Collection of Array(startIdx, endIdx)

    If Not IsDate(daily("当番日")) Then
        MsgBox "当日チェックシートの B3 に当番日 (yyyy/m/d) を入力してください。", vbExclamation
        GoTo DONE_EXIT
    End If
    Dim todayDate As Date: todayDate = CDate(daily("当番日"))

    Dim prevDay As Object: Set prevDay = LoadPrevDayFromHistory(todayDate, slots)

    ' 各人が「今日全スロット×」なら勤務可否リストから除外
    Dim members As Variant: members = AvailableMembers(roster, daily, positions, slots)
    If IsEmptyArray(members) Then
        MsgBox "割当可能な隊員が0名です。名簿・当日チェックを確認してください。", vbExclamation
        GoTo DONE_EXIT
    End If

    Dim assign(0 To 1, 0 To N_SLOTS - 1) As String
    Dim workCount As Object: Set workCount = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(members) To UBound(members)
        workCount(CStr(members(i))) = 0
    Next i

    ' 処理順: 10-17 → 17-24 → 0-8 → 8-8:40 → 8:40-10
    Dim slotOrder() As Long: slotOrder = BuildSlotProcessOrder()
    Dim k As Long, slotIdx As Long, col As Long
    For k = 0 To UBound(slotOrder)
        slotIdx = slotOrder(k)
        For col = 0 To 1
            assign(col, slotIdx) = PickAssignee( _
                col, slotIdx, members, roster, daily, positions, _
                prevDay, assign, workCount)
        Next col
    Next k

    Call EnforceDistinct12_17(assign, members, roster, daily, positions)

    Call WriteOutput(assign, todayDate)
    Call AppendToHistory(assign, slots, todayDate)

    Dim warnings As String
    warnings = Validate(assign, roster, daily, positions, prevDay, slots)
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
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub

' =============================================================
' 履歴から過去日を表示
' =============================================================
Public Sub ShowDateFromHistory()
    On Error GoTo EH
    Application.ScreenUpdating = False

    Dim wsOut As Worksheet: Set wsOut = ThisWorkbook.Worksheets(SHEET_OUT)
    Dim dv As Variant: dv = wsOut.Range("B2").Value
    If Not IsDate(dv) Then
        MsgBox "執務表シートの B2 に yyyy/m/d 形式で日付を入力してください。", vbExclamation
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

    Dim i As Long
    For i = 0 To N_SLOTS - 1
        wsOut.Cells(ROW_SLOT_START + i, 2).Value = ""
        wsOut.Cells(ROW_SLOT_START + i, 5).Value = ""
    Next i

    Dim slots() As String: slots = ReadTimeSlots()
    Dim slotMap As Object: Set slotMap = BuildSlotLabelMap(slots)
    Dim found As Boolean: found = False
    Dim r As Long, v As Variant
    For r = ROW_HIST_START To lastRow
        v = ws.Cells(r, 1).Value
        If IsDate(v) Then
            If CDate(v) = targetDate Then
                found = True
                Dim slotLabel As String
                slotLabel = Trim(CStr(Nz(ws.Cells(r, 2).Value, "")))
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
        MsgBox "指定日付のデータが履歴に見つかりません: " & _
               Format(targetDate, "yyyy/m/d"), vbExclamation
    End If

DONE_EXIT:
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub

' =============================================================
' データ読込
' =============================================================
Private Function ReadTimeSlots() As String()
    ' 執務表シートの A5..A29 から時間帯ラベルを読む
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_OUT)
    Dim arr(0 To N_SLOTS - 1) As String
    Dim i As Long
    For i = 0 To N_SLOTS - 1
        arr(i) = CStr(ws.Cells(ROW_SLOT_START + i, 1).Value)
    Next i
    ReadTimeSlots = arr
End Function

Private Function BuildSlotLabelMap(slots() As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim i As Long
    For i = 0 To UBound(slots)
        d(slots(i)) = i
    Next i
    Set BuildSlotLabelMap = d
End Function

Private Function LoadRoster() As Object
    ' 氏名の一覧 (名簿シート B2:B31)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_ROSTER)
    Dim r As Long, name As String
    For r = 2 To 31
        name = Trim(CStr(Nz(ws.Cells(r, 2).Value, "")))
        If Len(name) > 0 And Left(name, 1) <> "例" Then
            d(name) = True
        End If
    Next r
    Set LoadRoster = d
End Function

Private Function LoadPositions(slots() As String) As Object
    ' ポジション定義シートを読む
    ' 戻り値: position_name -> Dictionary(slotIdx -> True) のブロックセット
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_POS)

    Dim r As Long, posName As String
    For r = ROW_POS_START To ROW_POS_END
        posName = Trim(CStr(Nz(ws.Cells(r, 1).Value, "")))
        If Len(posName) > 0 Then
            Dim blocks As Object: Set blocks = CreateObject("Scripting.Dictionary")
            Dim j As Long
            For j = 0 To N_SLOTS - 1
                Dim cellVal As String
                cellVal = Trim(CStr(Nz(ws.Cells(r, 2 + j).Value, "")))
                If InStr(cellVal, "×") > 0 Or LCase(cellVal) = "x" Then
                    blocks(j) = True
                End If
            Next j
            Set d(posName) = blocks
        End If
    Next r
    Set LoadPositions = d
End Function

Private Function LoadDailyInput(slots() As String) As Object
    ' 当日チェックシートを読む
    ' レイアウト (1人 1行):
    '   A: No, B: 氏名,
    '   C: 休暇 (○), D: 当直 (○), E: 食当 (○),
    '   F: ポジション1, G: ポジション2,
    '   H: 除外1 から(時刻), I: 除外1 まで(時刻),
    '   J: 除外2 から, K: 除外2 まで,
    '   L: 備考
    ' 除外時間帯は時刻境界で指定 (例: 9時 から 17時 = slot 1〜8 ブロック)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)

    d("当番日") = ws.Range("B3").Value
    d("休日") = CLng(Nz(ws.Range("B4").Value, 0))

    Dim posByName As Object: Set posByName = CreateObject("Scripting.Dictionary")
    posByName.CompareMode = vbTextCompare
    Dim exclByName As Object: Set exclByName = CreateObject("Scripting.Dictionary")
    exclByName.CompareMode = vbTextCompare

    Dim markerMap As Object: Set markerMap = BuildTimeMarkerMap()

    Dim r As Long
    For r = ROW_CHECK_START To ROW_CHECK_END
        Dim name As String
        name = Trim(CStr(Nz(ws.Cells(r, 2).Value, "")))
        If Len(name) = 0 Then GoTo NEXT_R

        Dim posList As Collection: Set posList = New Collection

        ' チェックボックス3列 → 対応ポジションを付与
        If IsChecked(ws.Cells(r, 3).Value) Then posList.Add "休暇"
        If IsChecked(ws.Cells(r, 4).Value) Then posList.Add "当直"
        If IsChecked(ws.Cells(r, 5).Value) Then posList.Add "食当"

        ' ポジション1/2
        Dim p1 As String, p2 As String
        p1 = Trim(CStr(Nz(ws.Cells(r, 6).Value, "")))
        p2 = Trim(CStr(Nz(ws.Cells(r, 7).Value, "")))
        If Len(p1) > 0 Then posList.Add p1
        If Len(p2) > 0 Then posList.Add p2

        If posList.Count > 0 Then
            Set posByName(name) = posList
        End If

        ' 除外1/2: 時刻境界→スロット範囲変換
        Dim k As Long, coll As Collection
        Set coll = Nothing
        For k = 0 To 1
            Dim sCol As Long, eCol As Long
            sCol = 8 + k * 2
            eCol = sCol + 1
            Dim sLbl As String, eLbl As String
            sLbl = Trim(CStr(Nz(ws.Cells(r, sCol).Value, "")))
            eLbl = Trim(CStr(Nz(ws.Cells(r, eCol).Value, "")))
            If markerMap.Exists(sLbl) And markerMap.Exists(eLbl) Then
                Dim sBoundary As Long, eBoundary As Long
                sBoundary = CLng(markerMap(sLbl))
                eBoundary = CLng(markerMap(eLbl))
                ' "から"<"まで" に揃える
                If eBoundary < sBoundary Then
                    Dim tmp As Long: tmp = sBoundary: sBoundary = eBoundary: eBoundary = tmp
                End If
                ' ブロック範囲 = [sBoundary, eBoundary - 1] のスロット
                ' 例: 9時 から 17時 → boundary 1〜9 → block slot 1..8
                If eBoundary > sBoundary Then
                    If coll Is Nothing Then Set coll = New Collection
                    coll.Add Array(sBoundary, eBoundary - 1)
                End If
            End If
        Next k
        If Not coll Is Nothing Then
            Set exclByName(name) = coll
        End If
NEXT_R:
    Next r

    Set d("ポジション") = posByName
    Set d("除外") = exclByName
    Set LoadDailyInput = d
End Function

Private Function BuildTimeMarkerMap() As Object
    ' 時刻境界ラベル → boundary index (0..25) の辞書
    ' 26境界 = 25スロットの両端
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    d("8:40") = 0
    d("9時") = 1
    d("10時") = 2
    d("11時") = 3
    d("12時") = 4
    d("13時") = 5
    d("14時") = 6
    d("15時") = 7
    d("16時") = 8
    d("17時") = 9
    d("18時") = 10
    d("19時") = 11
    d("20時") = 12
    d("21時") = 13
    d("22時") = 14
    d("23時") = 15
    d("0時") = 16
    d("1時") = 17
    d("2時") = 18
    d("3時") = 19
    d("4時") = 20
    d("5時") = 21
    d("6時") = 22
    d("7時") = 23
    d("8時") = 24
    d("8:40(翌)") = 25
    Set BuildTimeMarkerMap = d
End Function

Private Function IsChecked(v As Variant) As Boolean
    If IsNull(v) Or IsEmpty(v) Then IsChecked = False: Exit Function
    IsChecked = (Len(Trim(CStr(v))) > 0)
End Function

Private Function LoadPrevDayFromHistory(beforeDate As Date, slots() As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To N_SLOTS - 1
        d(i) = Array("", "")
    Next i

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_HISTORY)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < ROW_HIST_START Then Set LoadPrevDayFromHistory = d: Exit Function

    Dim maxDate As Date: maxDate = 0
    Dim r As Long, v As Variant, cellDate As Date
    For r = ROW_HIST_START To lastRow
        v = ws.Cells(r, 1).Value
        If IsDate(v) Then
            cellDate = CDate(v)
            If cellDate < beforeDate And cellDate > maxDate Then maxDate = cellDate
        End If
    Next r
    If maxDate = 0 Then Set LoadPrevDayFromHistory = d: Exit Function

    Dim slotMap As Object: Set slotMap = BuildSlotLabelMap(slots)
    For r = ROW_HIST_START To lastRow
        v = ws.Cells(r, 1).Value
        If IsDate(v) Then
            If CDate(v) = maxDate Then
                Dim lbl As String
                lbl = Trim(CStr(Nz(ws.Cells(r, 2).Value, "")))
                If slotMap.Exists(lbl) Then
                    Dim idx As Long: idx = CLng(slotMap(lbl))
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

' =============================================================
' 可用メンバー = 1日全スロット × でない人
' =============================================================
Private Function AvailableMembers(roster As Object, daily As Object, _
    positions As Object, slots() As String) As Variant
    Dim col As Collection: Set col = New Collection
    Dim k As Variant
    For Each k In roster.Keys
        ' 全スロット × でなければ追加
        Dim hasAny As Boolean: hasAny = False
        Dim i As Long
        For i = 0 To N_SLOTS - 1
            If Not IsBlocked(CStr(k), i, daily, positions) Then
                hasAny = True
                Exit For
            End If
        Next i
        If hasAny Then col.Add CStr(k)
    Next k
    If col.Count = 0 Then
        AvailableMembers = Array()
        Exit Function
    End If
    Dim arr() As String
    ReDim arr(0 To col.Count - 1)
    Dim j As Long
    For j = 1 To col.Count
        arr(j - 1) = col.Item(j)
    Next j
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
' ブロック判定
' =============================================================
Private Function IsBlocked(name As String, slotIdx As Long, _
    daily As Object, positions As Object) As Boolean
    ' 1) 全ポジション (チェックボックス + ポジション1/2) の × を合算判定
    If daily("ポジション").Exists(name) Then
        Dim posList As Collection: Set posList = daily("ポジション")(name)
        Dim pos As Variant
        For Each pos In posList
            Dim posName As String: posName = CStr(pos)
            If positions.Exists(posName) Then
                If positions(posName).Exists(slotIdx) Then
                    IsBlocked = True
                    Exit Function
                End If
            End If
        Next pos
    End If

    ' 2) 追加除外時間帯
    If daily("除外").Exists(name) Then
        Dim coll As Collection: Set coll = daily("除外")(name)
        Dim it As Variant
        For Each it In coll
            If slotIdx >= CLng(it(0)) And slotIdx <= CLng(it(1)) Then
                IsBlocked = True
                Exit Function
            End If
        Next it
    End If

    IsBlocked = False
End Function

' =============================================================
' 割当ロジック
' =============================================================
Private Function PickAssignee(col As Long, slotIdx As Long, _
    members As Variant, roster As Object, daily As Object, _
    positions As Object, prevDay As Object, _
    assign() As String, workCount As Object) As String

    Dim candidates As Collection: Set candidates = New Collection
    Dim i As Long, name As String
    For i = LBound(members) To UBound(members)
        name = CStr(members(i))
        If Not IsBlocked(name, slotIdx, daily, positions) Then
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

    Dim bestName As String: bestName = ""
    Dim bestScore As Double: bestScore = 1E+18
    Dim c As Variant
    For Each c In candidates
        Dim s As Double
        s = ScoreCandidate(CStr(c), col, slotIdx, prevDay, assign, workCount)
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

Private Function ScoreCandidate(name As String, col As Long, slotIdx As Long, _
    prevDay As Object, assign() As String, workCount As Object) As Double
    Dim score As Double: score = 0

    ' (a) 総勤務回数 (バランス)
    score = score + CDbl(workCount(name)) * 10

    ' (b) 前日同時刻と同一人物
    Dim prevName As String: prevName = prevDay(slotIdx)(col)
    If Len(prevName) > 0 And prevName = name Then
        If slotIdx >= S_20_21 Then
            score = score + 1000   ' 夜20時以降は強回避
        Else
            score = score + 50     ' 日中は軽ペナルティ
        End If
    End If

    ' (c) 前日 slot+1 を軽優先 (前日の1つ上にずらす)
    If slotIdx + 1 <= N_SLOTS - 1 Then
        Dim prevNext As String: prevNext = prevDay(slotIdx + 1)(col)
        If prevNext = name Then score = score - 5
    End If

    ' (f) 10-17時未勤務者強優先
    If slotIdx >= S_10_11 And slotIdx <= S_17_18 - 1 Then
        If DaytimeCount(name, assign, slotIdx) = 0 Then
            score = score - 200
        End If
    End If

    ' (g) 深夜(22-4時)は一人1回まで
    If IsLateNight(slotIdx) Then
        If LateNightCount(name, assign, slotIdx) >= 1 Then
            score = score + 500
        End If
    End If

    ' (h) 12勤と17勤は別人
    If slotIdx = S_17_18 Then
        If assign(col, S_12_13) = name Then score = score + 500
    End If
    If slotIdx = S_12_13 Then
        If assign(col, S_17_18) = name Then score = score + 500
    End If

    ' (i) 連続割当ペナルティ (時間軸上の前後どちらかが同一人物なら避ける)
    If slotIdx > 0 Then
        If assign(col, slotIdx - 1) = name Then score = score + 50
    End If
    If slotIdx < N_SLOTS - 1 Then
        If assign(col, slotIdx + 1) = name Then score = score + 50
    End If

    ScoreCandidate = score
End Function

Private Function IsLateNight(slotIdx As Long) As Boolean
    IsLateNight = (slotIdx >= S_22_23 And slotIdx <= S_4_5 - 1)
End Function

Private Function DaytimeCount(name As String, assign() As String, _
    currentSlot As Long) As Long
    Dim i As Long, col As Long, cnt As Long
    For i = S_10_11 To S_17_18 - 1
        If i = currentSlot Then Exit For
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
        If i = currentSlot Then Exit For
        For col = 0 To 1
            If assign(col, i) = name Then cnt = cnt + 1
        Next col
    Next i
    LateNightCount = cnt
End Function

' =============================================================
' 処理順 (10-17 を優先)
' =============================================================
Private Function BuildSlotProcessOrder() As Long()
    Dim arr(0 To N_SLOTS - 1) As Long
    Dim n As Long: n = 0
    Dim i As Long
    For i = S_10_11 To S_17_18 - 1    ' 10-17 (slot 2..8)
        arr(n) = i: n = n + 1
    Next i
    For i = 9 To 15                    ' 17-24 (slot 9..15)
        arr(n) = i: n = n + 1
    Next i
    For i = 16 To 23                   ' 0-8 (slot 16..23)
        arr(n) = i: n = n + 1
    Next i
    arr(n) = 24: n = n + 1              ' 8-8:40
    arr(n) = 0: n = n + 1               ' 8:40-9
    arr(n) = 1: n = n + 1               ' 9-10
    BuildSlotProcessOrder = arr
End Function

' =============================================================
' 後処理: 12勤 と 17勤 の重複解消
' =============================================================
Private Sub EnforceDistinct12_17(assign() As String, _
    members As Variant, roster As Object, daily As Object, _
    positions As Object)
    Dim col As Long
    For col = 0 To 1
        If Len(assign(col, S_12_13)) > 0 And assign(col, S_12_13) = assign(col, S_17_18) Then
            Dim i As Long, alt As String
            For i = LBound(members) To UBound(members)
                alt = CStr(members(i))
                If alt = assign(col, S_12_13) Then GoTo NEXT_ALT
                If IsBlocked(alt, S_17_18, daily, positions) Then GoTo NEXT_ALT
                If assign(1 - col, S_17_18) = alt Then GoTo NEXT_ALT
                assign(col, S_17_18) = alt
                Exit For
NEXT_ALT:
            Next i
        End If
    Next col
End Sub

' =============================================================
' 出力 / 履歴追記
' =============================================================
Private Sub WriteOutput(assign() As String, displayDate As Date)
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

Private Sub AppendToHistory(assign() As String, slots() As String, _
    todayDate As Date)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_HISTORY)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < ROW_HIST_START - 1 Then lastRow = ROW_HIST_START - 1

    Dim r As Long, v As Variant
    For r = lastRow To ROW_HIST_START Step -1
        v = ws.Cells(r, 1).Value
        If IsDate(v) Then
            If CDate(v) = todayDate Then ws.Rows(r).Delete
        End If
    Next r

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < ROW_HIST_START - 1 Then lastRow = ROW_HIST_START - 1

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
' 検証
' =============================================================
Private Function Validate(assign() As String, roster As Object, _
    daily As Object, positions As Object, prevDay As Object, _
    slots() As String) As String
    Dim msgs As String, i As Long, col As Long

    For i = 0 To N_SLOTS - 1
        For col = 0 To 1
            If Len(assign(col, i)) = 0 Then
                msgs = msgs & " - " & slots(i) & " / " & _
                       IIf(col = 0, "通信", "受付") & " が空欄 (割当候補なし)" & vbCrLf
            End If
        Next col
    Next i

    Dim covered As Object: Set covered = CreateObject("Scripting.Dictionary")
    covered.CompareMode = vbTextCompare
    Dim k As Variant
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
        ' 10-17時 全スロット × の人は警告しない
        If covered(CStr(k)) = 0 And Not AllBlockedInDaytime(CStr(k), daily, positions) Then
            msgs = msgs & " - " & CStr(k) & " は 10-17時に未勤務" & vbCrLf
        End If
    Next k

    For i = S_20_21 To N_SLOTS - 1
        For col = 0 To 1
            If Len(assign(col, i)) > 0 And assign(col, i) = prevDay(i)(col) Then
                msgs = msgs & " - " & slots(i) & " / " & _
                       IIf(col = 0, "通信", "受付") & _
                       " が前日と同じ (" & assign(col, i) & ")" & vbCrLf
            End If
        Next col
    Next i

    Validate = msgs
End Function

Private Function AllBlockedInDaytime(name As String, daily As Object, _
    positions As Object) As Boolean
    Dim i As Long
    For i = S_10_11 To S_17_18 - 1
        If Not IsBlocked(name, i, daily, positions) Then
            AllBlockedInDaytime = False
            Exit Function
        End If
    Next i
    AllBlockedInDaytime = True
End Function

' =============================================================
' ユーティリティ
' =============================================================
Private Function Nz(v As Variant, alt As Variant) As Variant
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
