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

    ' Phase A: 固定枠配置
    '   - 通信 slot 0, 1 = 残留ポジションの人
    '   - 受付 slot 1, 24 = 署隊長伝令ポジションの人
    '   - 通信 slot 2 = 伝令 or 通信担当
    '   - 受付 slot 2 = 伝令 or 通信担当 (もう一人)
    Call ApplyFixedSlots(assign, members, roster, daily, positions, workCount)

    ' Phase B: 深夜スライド (slot 14..20 = 22時〜5時)
    '   前日 slot+1 の人を今日 slot に入れる
    Call ApplyNightSlide(assign, members, roster, daily, positions, _
                        prevDay, workCount)

    ' Phase C: 残りのスロットをスコアリングで埋める
    Dim slotOrder() As Long: slotOrder = BuildSlotProcessOrder()
    Dim k As Long, slotIdx As Long, col As Long
    For k = 0 To UBound(slotOrder)
        slotIdx = slotOrder(k)
        For col = 0 To 1
            ' 受付 8:40〜9 は斜線
            If col = 1 And slotIdx = 0 Then
                assign(1, 0) = ""
                GoTo NEXT_SLOT
            End If
            ' 既に固定枠 or 深夜スライドで埋まっていればスキップ
            If Len(assign(col, slotIdx)) > 0 Then GoTo NEXT_SLOT

            assign(col, slotIdx) = PickAssignee( _
                col, slotIdx, members, roster, daily, positions, _
                prevDay, assign, workCount)
NEXT_SLOT:
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
' 診断: 各隊員のポジション/除外/ブロック時間帯を一覧表示
'   除外が反映されない等のトラブル時に使用
' =============================================================
Public Sub DebugShowBlocks()
    On Error GoTo EH
    Application.ScreenUpdating = False

    Dim slots() As String: slots = ReadTimeSlots()
    Dim roster As Object: Set roster = LoadRoster()
    Dim positions As Object: Set positions = LoadPositions(slots)
    Dim daily As Object: Set daily = LoadDailyInput(slots)

    ' 診断シート作成 (既存あれば上書き)
    Dim wsName As String: wsName = "_診断"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo EH
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = wsName
    Else
        ws.Cells.Clear
    End If

    ws.Cells(1, 1).Value = "氏名"
    ws.Cells(1, 2).Value = "階級"
    ws.Cells(1, 3).Value = "ポジション (認識)"
    ws.Cells(1, 4).Value = "除外 (認識)"
    ws.Cells(1, 5).Value = "ブロック時間帯"
    Dim hdr As Range: Set hdr = ws.Range("A1:E1")
    hdr.Font.Bold = True
    hdr.Interior.Color = RGB(220, 230, 241)

    Dim r As Long: r = 2
    Dim k As Variant
    For Each k In roster.Keys
        Dim name As String: name = CStr(k)
        ws.Cells(r, 1).Value = name
        ws.Cells(r, 2).Value = CStr(roster(name))

        ' ポジション
        Dim posStr As String: posStr = ""
        If daily("ポジション").Exists(name) Then
            Dim pos As Variant
            For Each pos In daily("ポジション")(name)
                If Len(posStr) > 0 Then posStr = posStr & " + "
                posStr = posStr & CStr(pos)
            Next pos
        End If
        ws.Cells(r, 3).Value = posStr

        ' 除外
        Dim exclStr As String: exclStr = ""
        If daily("除外").Exists(name) Then
            Dim coll As Collection: Set coll = daily("除外")(name)
            Dim it As Variant
            For Each it In coll
                If Len(exclStr) > 0 Then exclStr = exclStr & " / "
                exclStr = exclStr & slots(CLng(it(0))) & " 〜 " & slots(CLng(it(1)))
            Next it
        End If
        ws.Cells(r, 4).Value = exclStr

        ' ブロック時間帯
        Dim blocked As String: blocked = ""
        Dim i As Long
        For i = 0 To N_SLOTS - 1
            If IsBlocked(name, i, daily, positions) Then
                If Len(blocked) > 0 Then blocked = blocked & ", "
                blocked = blocked & slots(i)
            End If
        Next i
        ws.Cells(r, 5).Value = blocked
        r = r + 1
    Next k

    ws.Columns("A:E").AutoFit
    ws.Activate

    Application.ScreenUpdating = True
    MsgBox "診断シート『_診断』を作成しました。" & vbCrLf & _
           "除外列やポジション列に意図した内容が反映されているか確認してください。", _
           vbInformation
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
    ' 氏名 + 階級 (名簿シート B2:C31)
    ' 戻り値: name -> rank_string
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_ROSTER)
    Dim r As Long, name As String, rank As String
    For r = 2 To 31
        name = Trim(CStr(Nz(ws.Cells(r, 2).Value, "")))
        rank = Trim(CStr(Nz(ws.Cells(r, 3).Value, "")))
        If Len(name) > 0 And Left(name, 1) <> "例" Then
            d(name) = rank
        End If
    Next r
    Set LoadRoster = d
End Function

Private Function RankValue(rank As String) As Long
    ' 高いほど上級。通信優先の数値。
    Select Case rank
        Case "司令補": RankValue = 4
        Case "士長": RankValue = 3
        Case "副士長": RankValue = 2
        Case "消防士": RankValue = 1
        Case Else: RankValue = 0   ' 未設定
    End Select
End Function

Private Function LoadPositions(slots() As String) As Object
    ' ポジション定義シートを読む (新レイアウト)
    ' A: ポジション名, B: 通信可(○), C: 受付可(○), D..AB: 時間帯25列 (×)
    ' 戻り値は Dictionary:
    '   "<name>" -> Scripting.Dictionary(slotIdx -> True) ブロックセット
    '   "<name>:comm" -> Boolean (通信に入れるか)
    '   "<name>:recv" -> Boolean (受付に入れるか)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_POS)

    Dim r As Long, posName As String
    For r = ROW_POS_START To ROW_POS_END
        posName = Trim(CStr(Nz(ws.Cells(r, 1).Value, "")))
        ' セパレータ (「↓ 以下は〜」等) はスキップ
        If Len(posName) > 0 And Left(posName, 1) <> "↓" And InStr(posName, "以下は") = 0 Then
            Dim blocks As Object: Set blocks = CreateObject("Scripting.Dictionary")
            Dim commFlag As String, recvFlag As String
            commFlag = Trim(CStr(Nz(ws.Cells(r, 2).Value, "")))
            recvFlag = Trim(CStr(Nz(ws.Cells(r, 3).Value, "")))

            Dim j As Long
            For j = 0 To N_SLOTS - 1
                Dim cellVal As String
                cellVal = Trim(CStr(Nz(ws.Cells(r, 4 + j).Value, "")))
                If InStr(cellVal, "×") > 0 Or LCase(cellVal) = "x" Then
                    blocks(j) = True
                End If
            Next j
            Set d(posName) = blocks

            ' カラム専属判定:
            '   両方空 or 両方○ → 両方OK
            '   通信可のみ○ → 通信専属
            '   受付可のみ○ → 受付専属
            Dim hasComm As Boolean, hasRecv As Boolean
            hasComm = (Len(commFlag) > 0)
            hasRecv = (Len(recvFlag) > 0)
            If hasComm And Not hasRecv Then
                d(posName & ":comm") = True
                d(posName & ":recv") = False
            ElseIf hasRecv And Not hasComm Then
                d(posName & ":comm") = False
                d(posName & ":recv") = True
            Else
                ' 両方指定 or 両方空 → 両方OK
                d(posName & ":comm") = True
                d(posName & ":recv") = True
            End If
        End If
    Next r
    Set LoadPositions = d
End Function

Private Function LoadDailyInput(slots() As String) As Object
    ' 当日チェックシートを読む (新レイアウト)
    '   A: No, B: 氏名
    '   C: 休暇, D: 当直, E: 食当, F: 研修, G: 警防力  (チェックボックス5個)
    '   H: ポジション1, I: ポジション2  (隊役割のみ)
    '   J: 除外1 から, K: 除外1 まで
    '   L: 除外2 から, M: 除外2 まで
    '   N: 備考
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

        ' チェックボックス5列 (C..G) → 対応ポジションを付与
        If IsChecked(ws.Cells(r, 3).Value) Then posList.Add "休暇"
        If IsChecked(ws.Cells(r, 4).Value) Then posList.Add "当直"
        If IsChecked(ws.Cells(r, 5).Value) Then posList.Add "食当"
        If IsChecked(ws.Cells(r, 6).Value) Then posList.Add "研修/出向"
        If IsChecked(ws.Cells(r, 7).Value) Then posList.Add "警防力"

        ' ポジション1/2 (H=8, I=9)
        Dim p1 As String, p2 As String
        p1 = Trim(CStr(Nz(ws.Cells(r, 8).Value, "")))
        p2 = Trim(CStr(Nz(ws.Cells(r, 9).Value, "")))
        If Len(p1) > 0 Then posList.Add p1
        If Len(p2) > 0 Then posList.Add p2

        If posList.Count > 0 Then
            Set posByName(name) = posList
        End If

        ' 除外1/2: J=10, K=11, L=12, M=13 (時刻境界→スロット範囲変換)
        Dim k As Long, coll As Collection
        Set coll = Nothing
        For k = 0 To 1
            Dim sCol As Long, eCol As Long
            sCol = 10 + k * 2
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
        ' 通信 or 受付 のどこか1スロットでも入れるなら対象
        Dim hasAny As Boolean: hasAny = False
        Dim i As Long, c As Long
        For i = 0 To N_SLOTS - 1
            For c = 0 To 1
                If Not IsBlocked(CStr(k), i, daily, positions, c) Then
                    hasAny = True
                    Exit For
                End If
            Next c
            If hasAny Then Exit For
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
'   col = -1: カラム非依存 (ポジションの時間ブロックと除外のみ)
'   col = 0 (通信) / 1 (受付): 専属カラム判定も行う
' =============================================================
Private Function IsBlocked(name As String, slotIdx As Long, _
    daily As Object, positions As Object, _
    Optional col As Long = -1) As Boolean

    ' 1) 全ポジションの × を合算判定
    Dim hasCommAllow As Boolean, hasRecvAllow As Boolean
    hasCommAllow = True
    hasRecvAllow = True

    If daily("ポジション").Exists(name) Then
        Dim posList As Collection: Set posList = daily("ポジション")(name)
        Dim pos As Variant
        ' まず時間ブロックチェック
        For Each pos In posList
            Dim posName As String: posName = CStr(pos)
            If positions.Exists(posName) Then
                If positions(posName).Exists(slotIdx) Then
                    IsBlocked = True
                    Exit Function
                End If
            End If
        Next pos

        ' カラム専属: ポジションの中に 通信専属 / 受付専属 があれば絞る
        ' 複数ポジション持ってる場合は AND (全部許可してる方のみ可)
        For Each pos In posList
            Dim pn As String: pn = CStr(pos)
            If positions.Exists(pn & ":comm") Then
                If Not CBool(positions(pn & ":comm")) Then hasCommAllow = False
            End If
            If positions.Exists(pn & ":recv") Then
                If Not CBool(positions(pn & ":recv")) Then hasRecvAllow = False
            End If
        Next pos
    End If

    ' 2) カラム専属違反チェック
    If col = 0 And Not hasCommAllow Then
        IsBlocked = True
        Exit Function
    End If
    If col = 1 And Not hasRecvAllow Then
        IsBlocked = True
        Exit Function
    End If

    ' 3) 追加除外時間帯
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

    ' 2段階候補選別:
    '   Strict: 4時間(±4スロット)以内に同一人物なし
    '   Relaxed: Strict で 0名なら ±2 まで緩和
    '   Final: それでも 0 なら制約外してスコアリング
    Dim strict As Collection: Set strict = New Collection
    Dim relaxed As Collection: Set relaxed = New Collection
    Dim anyOk As Collection: Set anyOk = New Collection
    Dim i As Long, name As String
    For i = LBound(members) To UBound(members)
        name = CStr(members(i))
        If IsBlocked(name, slotIdx, daily, positions, col) Then GoTo SKIP_ADD
        If col = 1 And assign(0, slotIdx) = name Then GoTo SKIP_ADD

        anyOk.Add name

        ' 近接チェック: 直前直後 同じ col に居ないか
        Dim nearDist As Long: nearDist = MinGapSameCol(name, col, slotIdx, assign)
        If nearDist >= 4 Then
            strict.Add name
        ElseIf nearDist >= 2 Then
            relaxed.Add name
        End If
SKIP_ADD:
    Next i

    Dim candidates As Collection
    If strict.Count > 0 Then
        Set candidates = strict
    ElseIf relaxed.Count > 0 Then
        Set candidates = relaxed
    Else
        Set candidates = anyOk
    End If

    If candidates.Count = 0 Then
        PickAssignee = ""
        Exit Function
    End If

    Dim bestName As String: bestName = ""
    Dim bestScore As Double: bestScore = 1E+18
    Dim c As Variant
    For Each c In candidates
        Dim s As Double
        s = ScoreCandidate(CStr(c), col, slotIdx, prevDay, assign, workCount, _
                           roster, daily, positions)
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
    prevDay As Object, assign() As String, workCount As Object, _
    roster As Object, _
    Optional daily As Object = Nothing, _
    Optional positions As Object = Nothing) As Double
    Dim score As Double: score = 0

    ' (a) 総勤務回数 (バランス)
    score = score + CDbl(workCount(name)) * 10

    ' (a2) 階級による 通信/受付 優先度
    '   司令補/士長 → 通信優先、副士長/消防士 → 受付優先
    '   消防士は電話対応が難しいので 通信 に大きなペナルティ
    Dim rank As String: rank = CStr(roster(name))
    Dim rv As Long: rv = RankValue(rank)
    If col = 0 Then  ' 通信
        Select Case rv
            Case 4: score = score - 30          ' 司令補: 通信優先
            Case 3: score = score - 30          ' 士長:   通信優先
            Case 2: score = score + 30          ' 副士長: やや受付寄り
            Case 1: score = score + 300         ' 消防士: 通信は原則×
        End Select
    Else             ' 受付
        Select Case rv
            Case 4: score = score + 20          ' 司令補: やや通信寄り
            Case 3: score = score + 20          ' 士長:   やや通信寄り
            Case 2: score = score - 30          ' 副士長: 受付優先
            Case 1: score = score - 30          ' 消防士: 受付優先
        End Select
    End If

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

    ' (f2) 専属ポジション保持者は自分のカラム内でローテ優先
    '   例: 署隊長伝令(受付専属) が受付ローテに外れないようにボーナス
    If Not daily Is Nothing And Not positions Is Nothing Then
        If IsExclusiveForCol(name, col, daily, positions) Then
            score = score - 40
        End If
    End If

    ' (f3) 食当者は 10-14時 (slot 2..5) を優先
    '   14時以降は食当×のため日中勤務の機会が 10-14 に限られる
    If Not daily Is Nothing Then
        If slotIdx >= S_10_11 And slotIdx <= S_14_15 - 1 Then
            If HasPosition(name, "食当", daily) Then
                score = score - 60
            End If
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

    ' (i) 近接ペナルティ (±4 スロット以内で同一人物は重く回避)
    '   ±1: +200, ±2: +100, ±3: +50, ±4: +30
    Dim dist As Long
    For dist = 1 To 4
        Dim penalty As Double
        Select Case dist
            Case 1: penalty = 200
            Case 2: penalty = 100
            Case 3: penalty = 50
            Case 4: penalty = 30
        End Select
        If slotIdx - dist >= 0 Then
            If assign(col, slotIdx - dist) = name Then score = score + penalty
        End If
        If slotIdx + dist < N_SLOTS Then
            If assign(col, slotIdx + dist) = name Then score = score + penalty
        End If
    Next dist

    ScoreCandidate = score
End Function

Private Function IsLateNight(slotIdx As Long) As Boolean
    IsLateNight = (slotIdx >= S_22_23 And slotIdx <= S_4_5 - 1)
End Function

' name が指定ポジションを今日持っているか
Private Function HasPosition(name As String, positionName As String, _
    daily As Object) As Boolean
    If Not daily("ポジション").Exists(name) Then
        HasPosition = False
        Exit Function
    End If
    Dim posList As Collection: Set posList = daily("ポジション")(name)
    Dim p As Variant
    For Each p In posList
        If CStr(p) = positionName Then
            HasPosition = True
            Exit Function
        End If
    Next p
    HasPosition = False
End Function

' name が持つポジションのいずれかが col 専属なら True
Private Function IsExclusiveForCol(name As String, col As Long, _
    daily As Object, positions As Object) As Boolean
    If Not daily("ポジション").Exists(name) Then
        IsExclusiveForCol = False
        Exit Function
    End If
    Dim posList As Collection: Set posList = daily("ポジション")(name)
    Dim p As Variant
    For Each p In posList
        Dim pn As String: pn = CStr(p)
        If col = 0 And positions.Exists(pn & ":recv") Then
            ' 受付不可 = 通信専属 (そのポジションが 通信可=○ かつ 受付可=空 のとき)
            If CBool(positions(pn & ":comm")) And Not CBool(positions(pn & ":recv")) Then
                IsExclusiveForCol = True
                Exit Function
            End If
        End If
        If col = 1 And positions.Exists(pn & ":comm") Then
            If CBool(positions(pn & ":recv")) And Not CBool(positions(pn & ":comm")) Then
                IsExclusiveForCol = True
                Exit Function
            End If
        End If
    Next p
    IsExclusiveForCol = False
End Function

' 指定人物 name が同じ col に割当されている最も近いスロットとの距離
' 1 以上の最小距離を返す。見つからなければ 999。
Private Function MinGapSameCol(name As String, col As Long, slotIdx As Long, _
    assign() As String) As Long
    Dim minGap As Long: minGap = 999
    Dim i As Long
    For i = 0 To N_SLOTS - 1
        If i = slotIdx Then GoTo NEXT_I
        If assign(col, i) = name Then
            Dim d As Long: d = Abs(i - slotIdx)
            If d < minGap Then minGap = d
        End If
NEXT_I:
    Next i
    MinGapSameCol = minGap
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
' Phase A: 固定枠配置
'   通信 slot 0, 1 = その日の「残留」ポジションの人
'   受付 slot 1, 24 = その日の「署隊長伝令」ポジションの人
'   通信 slot 2, 受付 slot 2 = 「伝令」か「通信担当」のペア
' =============================================================
Private Sub ApplyFixedSlots(assign() As String, members As Variant, _
    roster As Object, daily As Object, positions As Object, _
    workCount As Object)

    ' 残留の人 (通信 slot 0, 1)
    Dim zanName As String: zanName = FindMemberWithPosition(members, daily, "残留")
    If Len(zanName) > 0 Then
        If Not IsBlocked(zanName, 0, daily, positions, 0) Then
            assign(0, 0) = zanName
            workCount(zanName) = workCount(zanName) + 1
        End If
        If Not IsBlocked(zanName, 1, daily, positions, 0) Then
            assign(0, 1) = zanName
            workCount(zanName) = workCount(zanName) + 1
        End If
    End If

    ' 署隊長伝令の人 (受付 slot 1, 24)
    Dim deName As String: deName = FindMemberWithPosition(members, daily, "署隊長伝令")
    If Len(deName) > 0 Then
        If Not IsBlocked(deName, 1, daily, positions, 1) Then
            assign(1, 1) = deName
            workCount(deName) = workCount(deName) + 1
        End If
        If Not IsBlocked(deName, 24, daily, positions, 1) Then
            assign(1, 24) = deName
            workCount(deName) = workCount(deName) + 1
        End If
    End If

    ' 伝令 or 通信担当 (通信 slot 2, 受付 slot 2)
    Dim denName As String, tsuName As String
    denName = FindMemberWithPosition(members, daily, "伝令")
    tsuName = FindMemberWithPosition(members, daily, "通信担当")
    ' 通信 slot 2 = 伝令 or 通信担当 を優先
    If Len(denName) > 0 And Not IsBlocked(denName, S_10_11, daily, positions, 0) Then
        assign(0, S_10_11) = denName
        workCount(denName) = workCount(denName) + 1
    ElseIf Len(tsuName) > 0 And Not IsBlocked(tsuName, S_10_11, daily, positions, 0) Then
        assign(0, S_10_11) = tsuName
        workCount(tsuName) = workCount(tsuName) + 1
    End If
    ' 受付 slot 2 = もう一人
    Dim recvCandidate As String
    If Len(tsuName) > 0 And tsuName <> assign(0, S_10_11) Then
        recvCandidate = tsuName
    ElseIf Len(denName) > 0 And denName <> assign(0, S_10_11) Then
        recvCandidate = denName
    End If
    If Len(recvCandidate) > 0 And Not IsBlocked(recvCandidate, S_10_11, daily, positions, 1) Then
        assign(1, S_10_11) = recvCandidate
        workCount(recvCandidate) = workCount(recvCandidate) + 1
    End If
End Sub

' 指定したポジションを持つ最初のメンバー名を返す
Private Function FindMemberWithPosition(members As Variant, _
    daily As Object, positionName As String) As String
    Dim i As Long, name As String
    For i = LBound(members) To UBound(members)
        name = CStr(members(i))
        If daily("ポジション").Exists(name) Then
            Dim posList As Collection: Set posList = daily("ポジション")(name)
            Dim p As Variant
            For Each p In posList
                If CStr(p) = positionName Then
                    FindMemberWithPosition = name
                    Exit Function
                End If
            Next p
        End If
    Next i
    FindMemberWithPosition = ""
End Function

' =============================================================
' Phase B: 深夜スライド (slot 14..20 = 22時〜5時)
'   前日の slot+1 の人を今日の slot に入れる
'   = 「前日1時勤務の人は今回0時勤務」
' =============================================================
Private Sub ApplyNightSlide(assign() As String, members As Variant, _
    roster As Object, daily As Object, positions As Object, _
    prevDay As Object, workCount As Object)

    Dim slotIdx As Long, col As Long
    For slotIdx = S_22_23 To S_4_5   ' 14..20
        For col = 0 To 1
            ' 既に埋まってたらスキップ (通常このタイミングは未割当)
            If Len(assign(col, slotIdx)) > 0 Then GoTo NEXT_N

            ' 前日の slot+1 の人を今日の slot に
            If slotIdx + 1 > N_SLOTS - 1 Then GoTo NEXT_N
            Dim prevArr As Variant: prevArr = prevDay(slotIdx + 1)
            Dim prevName As String: prevName = CStr(prevArr(col))
            If Len(prevName) = 0 Then GoTo NEXT_N

            ' その人が今日の名簿に存在し、ブロックされておらず、
            ' 通信と受付で同時刻別人制約も満たすなら採用
            If Not roster.Exists(prevName) Then GoTo NEXT_N
            If IsBlocked(prevName, slotIdx, daily, positions, col) Then GoTo NEXT_N
            If col = 1 And assign(0, slotIdx) = prevName Then GoTo NEXT_N

            assign(col, slotIdx) = prevName
            workCount(prevName) = workCount(prevName) + 1
NEXT_N:
        Next col
    Next slotIdx
End Sub

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
                If IsBlocked(alt, S_17_18, daily, positions, col) Then GoTo NEXT_ALT
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
            ' 受付 8:40〜9 は斜線なので空欄チェックから除外
            If col = 1 And i = 0 Then GoTo NEXT_CHECK
            If Len(assign(col, i)) = 0 Then
                msgs = msgs & " - " & slots(i) & " / " & _
                       IIf(col = 0, "通信", "受付") & " が空欄 (割当候補なし)" & vbCrLf
            End If
NEXT_CHECK:
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
        ' 消防士階級は数学的に日中未勤務になりやすいので警告対象外
        If covered(CStr(k)) = 0 _
                And Not AllBlockedInDaytime(CStr(k), daily, positions) _
                And CStr(roster(k)) <> "消防士" Then
            msgs = msgs & " - " & CStr(k) & " は 10-17時に未勤務" & vbCrLf
        End If
    Next k

    ' 夜20時以降 前日同時刻と同一人物 (固定枠由来は除外)
    Dim fixedExempt As Object: Set fixedExempt = BuildFixedExemptSet(daily)
    For i = S_20_21 To N_SLOTS - 1
        For col = 0 To 1
            If Len(assign(col, i)) > 0 And assign(col, i) = prevDay(i)(col) Then
                ' 固定枠由来 (残留 8:40-10通信, 署隊長伝令 受付9-10/8-8:40) はスキップ
                Dim exemptKey As String: exemptKey = i & "|" & col & "|" & assign(col, i)
                If Not fixedExempt.Exists(exemptKey) Then
                    msgs = msgs & " - " & slots(i) & " / " & _
                           IIf(col = 0, "通信", "受付") & _
                           " が前日と同じ (" & assign(col, i) & ")" & vbCrLf
                End If
            End If
        Next col
    Next i

    ' 12勤と17勤が同じ
    For col = 0 To 1
        If Len(assign(col, S_12_13)) > 0 And assign(col, S_12_13) = assign(col, S_17_18) Then
            msgs = msgs & " - 12勤と17勤が同じ人物 (" & assign(col, S_12_13) & ")" & vbCrLf
        End If
    Next col

    Validate = msgs
End Function

' 固定枠 (残留・署隊長伝令) で配置された slot+col+name の組合せ
' これらは前日と同じ人物が割当られても自然なので警告除外
Private Function BuildFixedExemptSet(daily As Object) As Object
    Dim s As Object: Set s = CreateObject("Scripting.Dictionary")
    s.CompareMode = vbTextCompare
    ' 残留 → 通信 slot 0, 1
    Dim names As Variant
    names = FindAllMembersWithPosition(daily, "残留")
    Dim name As Variant
    For Each name In names
        s("0|0|" & CStr(name)) = True
        s("1|0|" & CStr(name)) = True
    Next name
    ' 署隊長伝令 → 受付 slot 1, 24
    names = FindAllMembersWithPosition(daily, "署隊長伝令")
    For Each name In names
        s("1|1|" & CStr(name)) = True
        s("24|1|" & CStr(name)) = True
    Next name
    Set BuildFixedExemptSet = s
End Function

Private Function FindAllMembersWithPosition(daily As Object, _
    positionName As String) As Variant
    Dim arr() As String
    ReDim arr(0 To 30)
    Dim n As Long: n = 0
    Dim k As Variant
    For Each k In daily("ポジション").Keys
        Dim posList As Collection: Set posList = daily("ポジション")(k)
        Dim p As Variant
        For Each p In posList
            If CStr(p) = positionName Then
                arr(n) = CStr(k)
                n = n + 1
                Exit For
            End If
        Next p
    Next k
    If n = 0 Then
        FindAllMembersWithPosition = Array()
        Exit Function
    End If
    ReDim Preserve arr(0 To n - 1)
    FindAllMembersWithPosition = arr
End Function

Private Function AllBlockedInDaytime(name As String, daily As Object, _
    positions As Object) As Boolean
    ' 通信でも受付でも入れないなら True
    Dim i As Long
    For i = S_10_11 To S_17_18 - 1
        If Not IsBlocked(name, i, daily, positions, 0) Then
            AllBlockedInDaytime = False
            Exit Function
        End If
        If Not IsBlocked(name, i, daily, positions, 1) Then
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
