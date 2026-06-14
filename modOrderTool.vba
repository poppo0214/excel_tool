Option Explicit

Public Const SHEET_SETTING As String = "基本設定"
Public Const SHEET_CALC As String = "発注数計算"

Public Const MAX_ITEMS As Long = 20
Public Const MONEY_FORMAT As String = "¥#,##0"

Public Const BUDGET_CELL As String = "B2"

Public Const SET_PRODUCT_START_ROW As Long = 4
Public Const SET_PRODUCT_COL As Long = 2
Public Const SET_DEPT_START_ROW As Long = 4
Public Const SET_DEPT_COL As Long = 5

Public Const HOPE_START_ROW As Long = 13
Public Const START_COL As Long = 1
Public Const PRICE_COL As Long = 2
Public Const DEPT_START_COL As Long = 3

Public Const STATUS_COL As Long = 12

Public Const PENDING_TYPE_CELL As String = "Z1"
Public Const PENDING_ROW_CELL As String = "Z2"
Public Const PENDING_COL_CELL As String = "Z3"

'========================
' 始める
'========================
Public Sub 始める()

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim wsS As Worksheet
    Dim wsC As Worksheet

    Set wsS = GetOrCreateSheet(SHEET_SETTING)
    Set wsC = GetOrCreateSheet(SHEET_CALC)

    wsS.Cells.Clear
    wsC.Cells.Clear

    CreateSettingSheetLayout wsS
    CreateCalcSheetInitialLayout wsC

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "初期設定が完了しました。基本設定シートに商品名・課名を入力してください。", vbInformation

End Sub

Private Function GetOrCreateSheet(sheetName As String) As Worksheet

    On Error Resume Next
    Set GetOrCreateSheet = Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If

End Function

Private Sub CreateSettingSheetLayout(ws As Worksheet)

    Dim i As Long

    With ws
        .Range("A1").Value = "基本設定"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16

        .Range("A3").Value = "No."
        .Range("B3").Value = "商品名"
        .Range("D3").Value = "No."
        .Range("E3").Value = "課名"

        For i = 1 To MAX_ITEMS
            .Cells(SET_PRODUCT_START_ROW + i - 1, 1).Value = i
            .Cells(SET_DEPT_START_ROW + i - 1, 4).Value = i
        Next i

        .Range("B4:B23,E4:E23").Interior.Color = RGB(221, 235, 247)
        .Range("A3:B23,D3:E23").Borders.LineStyle = xlContinuous

        .Range("A25").Value = "※商品名・課名は最大20個まで入力できます。"
        .Range("A26").Value = "※入力された商品名・課名の数から、商品数・課数を自動取得します。"

        AddCellMemo .Range("B4"), _
            "商品名を上から詰めて入力してください。" & vbCrLf & vbCrLf & _
            "入力ルール" & vbCrLf & _
            "・最大20個まで" & vbCrLf & _
            "・1つ以上必須" & vbCrLf & _
            "・途中に空欄を入れない" & vbCrLf & _
            "・同じ商品名は使えません"

        AddCellMemo .Range("E4"), _
            "課名を上から詰めて入力してください。" & vbCrLf & vbCrLf & _
            "入力ルール" & vbCrLf & _
            "・最大20個まで" & vbCrLf & _
            "・1つ以上必須" & vbCrLf & _
            "・途中に空欄を入れない" & vbCrLf & _
            "・同じ課名は使えません"

        .Columns("A:E").AutoFit
    End With

End Sub

Private Sub CreateCalcSheetInitialLayout(ws As Worksheet)

    With ws
        .Range("A1").Value = "発注数計算"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16

        .Range("A2").Value = "予算"
        .Range(BUDGET_CELL).Interior.Color = RGB(221, 235, 247)
        .Range(BUDGET_CELL).NumberFormat = MONEY_FORMAT

        .Range("A4").Value = "使い方"
        .Range("A5").Value = "① 基本設定シートに商品名・課名を入力"
        .Range("A6").Value = "②「発注数計算表作成」を押す"
        .Range("A7").Value = "③ 予算・単価・課別希望数を入力"
        .Range("A8").Value = "④「予算チェック・自動計算」を押す"
        .Range("A9").Value = "⑤ 必要に応じて下段の発注数を手動調整"
        .Range("A10").Value = "⑥「手動調整・再計算」を押す"

        .Range("D2").Value = "セル色の意味"
        .Range("D3").Value = "青色：入力セル"
        .Range("D4").Value = "黄色：手動調整セル"
        .Range("D5").Value = "赤字：直近の変更・反映箇所"

        AddCellMemo .Range(BUDGET_CELL), _
            "発注に使用できる予算を入力してください。" & vbCrLf & vbCrLf & _
            "入力ルール" & vbCrLf & _
            "・正の整数" & vbCrLf & _
            "・0、マイナス、小数、文字、空欄は不可" & vbCrLf & _
            "・単位は円" & vbCrLf & vbCrLf & _
            "入力例" & vbCrLf & _
            "100000"

        AddCellMemo .Range("A4"), _
            "使い方" & vbCrLf & vbCrLf & _
            "① 基本設定シートに商品名・課名を入力" & vbCrLf & _
            "② 発注数計算表作成" & vbCrLf & _
            "③ 予算・単価・希望数を入力" & vbCrLf & _
            "④ 予算チェック・自動計算" & vbCrLf & _
            "⑤ 必要に応じて発注数を手動修正" & vbCrLf & _
            "⑥ 手動調整・再計算" & vbCrLf & vbCrLf & _
            "セル色" & vbCrLf & _
            "青：入力セル" & vbCrLf & _
            "黄：手動調整セル" & vbCrLf & _
            "赤字：直近の変更・反映箇所"

        .Columns("A:Z").ColumnWidth = 13
    End With

End Sub

Private Sub AddCellMemo(targetCell As Range, memoText As String)

    On Error Resume Next
    targetCell.ClearComments
    On Error GoTo 0

    targetCell.AddComment memoText
    targetCell.Comment.Visible = False

End Sub

'========================
' 発注数計算表作成
'========================
Public Sub 発注数計算表作成()

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_CALC)

    Dim products As Variant
    Dim depts As Variant

    products = GetProductNames()
    depts = GetDeptNames()

    If GetProductCount() = 0 Or GetDeptCount() = 0 Then
        MsgBox "基本設定シートに商品名・課名を入力してください。", vbExclamation
        Exit Sub
    End If

    Dim errMsg As String
    errMsg = ValidateMasterNames()

    If errMsg <> "" Then
        MsgBox errMsg, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ws.Range("A13:Z300").Clear
    ws.Range(PENDING_TYPE_CELL & ":" & PENDING_COL_CELL).ClearContents

    CreateOneTable ws, HOPE_START_ROW, "希望数入力", products, depts, True
    CreateOneTable ws, GetOrderStartRow(), "発注数・最終調整", products, depts, False
    CreateStatusArea ws

    ws.Range(PENDING_TYPE_CELL & ":" & PENDING_COL_CELL).EntireColumn.Hidden = True

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "発注数計算表を作成しました。", vbInformation

End Sub

Private Sub CreateOneTable(ws As Worksheet, startRow As Long, title As String, products As Variant, depts As Variant, isHope As Boolean)

    Dim pCount As Long
    Dim dCount As Long

    pCount = GetProductCount()
    dCount = GetDeptCount()

    Dim totalQtyCol As Long
    Dim amountCol As Long

    totalQtyCol = DEPT_START_COL + dCount
    amountCol = totalQtyCol + 1

    Dim titleRow As Long
    Dim headerRow As Long

    titleRow = startRow
    headerRow = startRow + 1

    ws.Cells(titleRow, START_COL).Value = "【" & title & "】"
    ws.Cells(titleRow, START_COL).Font.Bold = True

    ws.Cells(headerRow, START_COL).Value = "商品名"
    ws.Cells(headerRow, PRICE_COL).Value = "単価"

    Dim j As Long
    For j = 1 To dCount
        ws.Cells(headerRow, DEPT_START_COL + j - 1).Value = depts(j)
    Next j

    If isHope Then
        ws.Cells(headerRow, totalQtyCol).Value = "希望数合計"
        ws.Cells(headerRow, amountCol).Value = "希望金額"
    Else
        ws.Cells(headerRow, totalQtyCol).Value = "発注数合計"
        ws.Cells(headerRow, amountCol).Value = "発注金額"
    End If

    Dim i As Long
    For i = 1 To pCount
        ws.Cells(headerRow + i, START_COL).Value = products(i)

        If isHope Then
            ws.Cells(headerRow + i, PRICE_COL).Interior.Color = RGB(221, 235, 247)
            ws.Range(ws.Cells(headerRow + i, DEPT_START_COL), ws.Cells(headerRow + i, DEPT_START_COL + dCount - 1)).Interior.Color = RGB(221, 235, 247)
        Else
            ws.Range(ws.Cells(headerRow + i, DEPT_START_COL), ws.Cells(headerRow + i, totalQtyCol)).Interior.Color = RGB(255, 242, 204)
        End If

        ws.Cells(headerRow + i, PRICE_COL).NumberFormat = MONEY_FORMAT
        ws.Cells(headerRow + i, amountCol).NumberFormat = MONEY_FORMAT
    Next i

    Dim totalRow As Long
    totalRow = headerRow + pCount + 1

    ws.Cells(totalRow, amountCol - 1).Value = "合計金額"
    ws.Cells(totalRow, amountCol).NumberFormat = MONEY_FORMAT

    ws.Cells(totalRow + 1, amountCol - 1).Value = "予算との差額"
    ws.Cells(totalRow + 1, amountCol).NumberFormat = MONEY_FORMAT

    With ws.Range(ws.Cells(headerRow, START_COL), ws.Cells(totalRow + 1, amountCol))
        .Borders.LineStyle = xlContinuous
    End With

    If isHope Then
        AddCellMemo ws.Cells(headerRow + 1, PRICE_COL), _
            "商品の単価を入力してください。" & vbCrLf & vbCrLf & _
            "入力ルール" & vbCrLf & _
            "・正の整数" & vbCrLf & _
            "・0、マイナス、小数、文字、空欄は不可" & vbCrLf & _
            "・単位は円" & vbCrLf & vbCrLf & _
            "入力例" & vbCrLf & _
            "500"

        AddCellMemo ws.Cells(headerRow + 1, DEPT_START_COL), _
            "各課の希望数量を入力してください。" & vbCrLf & vbCrLf & _
            "入力ルール" & vbCrLf & _
            "・0以上の整数" & vbCrLf & _
            "・マイナス、小数、文字、空欄は不可" & vbCrLf & vbCrLf & _
            "入力例" & vbCrLf & _
            "10"
    Else
        AddCellMemo ws.Cells(headerRow + 1, DEPT_START_COL), _
            "手動で編集可能です。" & vbCrLf & vbCrLf & _
            "入力ルール" & vbCrLf & _
            "・0以上の整数" & vbCrLf & _
            "・マイナス、小数、文字、空欄は不可" & vbCrLf & vbCrLf & _
            "変更後に「手動調整・再計算」を押すと、発注数合計を再計算します。"

        AddCellMemo ws.Cells(headerRow + 1, totalQtyCol), _
            "手動で編集可能です。" & vbCrLf & vbCrLf & _
            "入力ルール" & vbCrLf & _
            "・0以上の整数" & vbCrLf & _
            "・マイナス、小数、文字、空欄は不可" & vbCrLf & vbCrLf & _
            "変更後に「手動調整・再計算」を押すと、課別発注数を希望数比率で再配分します。"
    End If

    ws.Columns("A:Z").AutoFit

End Sub

Private Sub CreateStatusArea(ws As Worksheet)

    ws.Cells(13, STATUS_COL).Value = "【状態】"
    ws.Cells(14, STATUS_COL).Value = "予算"
    ws.Cells(15, STATUS_COL).Value = "発注金額"
    ws.Cells(16, STATUS_COL).Value = "予算との差額"

    ws.Cells(18, STATUS_COL).Value = "【エラー】"
    ws.Cells(19, STATUS_COL).Value = "なし"

    ws.Range(ws.Cells(14, STATUS_COL + 1), ws.Cells(16, STATUS_COL + 1)).NumberFormat = MONEY_FORMAT

    AddCellMemo ws.Cells(13, STATUS_COL), _
        "現在の計算結果を表示します。" & vbCrLf & vbCrLf & _
        "表示内容" & vbCrLf & _
        "・予算" & vbCrLf & _
        "・発注金額" & vbCrLf & _
        "・予算との差額" & vbCrLf & vbCrLf & _
        "予算との差額がマイナスの場合は、予算超過です。"

    AddCellMemo ws.Cells(18, STATUS_COL), _
        "入力ミスや予算超過を表示します。" & vbCrLf & vbCrLf & _
        "表示例" & vbCrLf & _
        "・予算は正の整数で入力してください。" & vbCrLf & _
        "・発注数は0以上の整数で入力してください。" & vbCrLf & _
        "・発注金額が予算を超過しています。" & vbCrLf & vbCrLf & _
        "「なし」の場合は問題ありません。"

    ws.Columns(STATUS_COL).ColumnWidth = 18
    ws.Columns(STATUS_COL + 1).ColumnWidth = 18

End Sub

'========================
' 予算チェック・自動計算
'========================
Public Sub 予算チェック_自動計算()

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_CALC)

    Dim errMsg As String
    errMsg = ValidateHopeInput(ws)

    If errMsg <> "" Then
        ShowError ws, errMsg
        MsgBox errMsg, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ClearManualMarks ws

    Dim budget As Long
    Dim hopeAmount As Long

    budget = CLng(ws.Range(BUDGET_CELL).Value)
    hopeAmount = CalcHopeAmount(ws)

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Dim mode As String

    If hopeAmount <= budget Then
        Dim ans As VbMsgBoxResult

        ans = MsgBox( _
            "希望金額：" & Format(hopeAmount, MONEY_FORMAT) & vbCrLf & _
            "予算：" & Format(budget, MONEY_FORMAT) & vbCrLf & _
            "予算との差額：" & Format(budget - hopeAmount, MONEY_FORMAT) & vbCrLf & vbCrLf & _
            "予算内で増やしますか？" & vbCrLf & _
            "はい：予算内で増やす" & vbCrLf & _
            "いいえ：希望数どおりにする", _
            vbYesNoCancel + vbQuestion)

        If ans = vbCancel Then Exit Sub

        If ans = vbYes Then
            mode = "increase"
        Else
            mode = "same"
        End If
    Else
        MsgBox _
            "希望金額：" & Format(hopeAmount, MONEY_FORMAT) & vbCrLf & _
            "予算：" & Format(budget, MONEY_FORMAT) & vbCrLf & _
            "予算を " & Format(hopeAmount - budget, MONEY_FORMAT) & " 超過しています。" & vbCrLf & vbCrLf & _
            "比率調整により発注数を自動計算します。", _
            vbInformation

        mode = "decrease"
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    OutputOptimizedOrder ws, mode
    UpdateAllTotals ws
    UpdateStatusAndErrors ws

    ws.Range(PENDING_TYPE_CELL & ":" & PENDING_COL_CELL).ClearContents

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'========================
' 手動調整・再計算
'========================
Public Sub 手動調整_再計算()

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_CALC)

    Dim pType As String
    Dim r As Long
    Dim c As Long

    pType = ws.Range(PENDING_TYPE_CELL).Value
    r = Val(ws.Range(PENDING_ROW_CELL).Value)
    c = Val(ws.Range(PENDING_COL_CELL).Value)

    If pType = "" Or r = 0 Or c = 0 Then
        MsgBox "変更された発注数セルがありません。", vbInformation
        Exit Sub
    End If

    Dim errMsg As String
    errMsg = ValidateOrderInput(ws, False)

    If errMsg <> "" Then
        ShowError ws, errMsg
        MsgBox errMsg, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim dCount As Long
    dCount = GetDeptCount()

    Dim totalQtyCol As Long
    totalQtyCol = DEPT_START_COL + dCount

    If c = totalQtyCol Then
        AllocateRowByHopeRatio ws, r, CLng(ws.Cells(r, totalQtyCol).Value)
        ws.Range(ws.Cells(r, DEPT_START_COL), ws.Cells(r, totalQtyCol)).Font.Color = RGB(255, 0, 0)

    ElseIf c >= DEPT_START_COL And c < totalQtyCol Then
        ws.Cells(r, totalQtyCol).Value = SumRow(ws, r, DEPT_START_COL, totalQtyCol - 1)
        ws.Range(ws.Cells(r, c), ws.Cells(r, totalQtyCol)).Font.Color = RGB(255, 0, 0)
    End If

    UpdateAllTotals ws
    UpdateStatusAndErrors ws

    ws.Range(PENDING_TYPE_CELL & ":" & PENDING_COL_CELL).ClearContents

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'========================
' 最適化処理
'========================
Private Sub OutputOptimizedOrder(ws As Worksheet, mode As String)

    Dim pCount As Long
    Dim dCount As Long

    pCount = GetProductCount()
    dCount = GetDeptCount()

    Dim hopeHeader As Long
    Dim orderHeader As Long

    hopeHeader = HOPE_START_ROW + 1
    orderHeader = GetOrderStartRow() + 1

    Dim totalQtyCol As Long
    totalQtyCol = DEPT_START_COL + dCount

    Dim hopeTotals() As Long
    Dim prices() As Long
    Dim orderTotals() As Long

    ReDim hopeTotals(1 To pCount)
    ReDim prices(1 To pCount)
    ReDim orderTotals(1 To pCount)

    Dim i As Long

    For i = 1 To pCount
        prices(i) = CLng(ws.Cells(hopeHeader + i, PRICE_COL).Value)
        hopeTotals(i) = SumRow(ws, hopeHeader + i, DEPT_START_COL, totalQtyCol - 1)
    Next i

    Dim budget As Long
    Dim hopeAmount As Long

    budget = CLng(ws.Range(BUDGET_CELL).Value)
    hopeAmount = CalcHopeAmount(ws)

    If mode = "same" Then
        For i = 1 To pCount
            orderTotals(i) = hopeTotals(i)
        Next i
    Else
        Dim ratio As Double
        ratio = budget / hopeAmount

        For i = 1 To pCount
            orderTotals(i) = Int(hopeTotals(i) * ratio)
        Next i

        Dim currentAmount As Long
        currentAmount = CalcAmountFromTotals(orderTotals, prices)

        Dim added As Boolean
        Do
            added = False
            For i = 1 To pCount
                If currentAmount + prices(i) <= budget Then
                    orderTotals(i) = orderTotals(i) + 1
                    currentAmount = currentAmount + prices(i)
                    added = True
                End If
            Next i
        Loop While added
    End If

    For i = 1 To pCount
        ws.Cells(orderHeader + i, START_COL).Value = ws.Cells(hopeHeader + i, START_COL).Value
        ws.Cells(orderHeader + i, PRICE_COL).Value = ws.Cells(hopeHeader + i, PRICE_COL).Value
        ws.Cells(orderHeader + i, PRICE_COL).NumberFormat = MONEY_FORMAT

        AllocateRowByHopeRatio ws, orderHeader + i, orderTotals(i)
    Next i

End Sub

Private Sub AllocateRowByHopeRatio(ws As Worksheet, orderRow As Long, newTotal As Long)

    Dim dCount As Long
    dCount = GetDeptCount()

    Dim orderHeader As Long
    Dim hopeHeader As Long

    orderHeader = GetOrderStartRow() + 1
    hopeHeader = HOPE_START_ROW + 1

    Dim productIndex As Long
    productIndex = orderRow - orderHeader

    Dim hopeRow As Long
    hopeRow = hopeHeader + productIndex

    Dim totalQtyCol As Long
    totalQtyCol = DEPT_START_COL + dCount

    Dim hopeTotal As Long
    hopeTotal = SumRow(ws, hopeRow, DEPT_START_COL, totalQtyCol - 1)

    Dim j As Long

    If hopeTotal = 0 Then
        For j = DEPT_START_COL To totalQtyCol - 1
            ws.Cells(orderRow, j).Value = 0
        Next j
        ws.Cells(orderRow, totalQtyCol).Value = 0
        Exit Sub
    End If

    Dim base() As Long
    Dim frac() As Double

    ReDim base(1 To dCount)
    ReDim frac(1 To dCount)

    Dim allocated As Long
    allocated = 0

    For j = 1 To dCount
        Dim raw As Double
        raw = newTotal * ws.Cells(hopeRow, DEPT_START_COL + j - 1).Value / hopeTotal
        base(j) = Int(raw)
        frac(j) = raw - base(j)
        allocated = allocated + base(j)
    Next j

    Dim remain As Long
    remain = newTotal - allocated

    Do While remain > 0
        Dim maxIdx As Long
        maxIdx = 1

        For j = 2 To dCount
            If frac(j) > frac(maxIdx) Then
                maxIdx = j
            End If
        Next j

        base(maxIdx) = base(maxIdx) + 1
        frac(maxIdx) = -1
        remain = remain - 1
    Loop

    For j = 1 To dCount
        ws.Cells(orderRow, DEPT_START_COL + j - 1).Value = base(j)
    Next j

    ws.Cells(orderRow, totalQtyCol).Value = newTotal

End Sub

'========================
' 集計
'========================
Private Sub UpdateAllTotals(ws As Worksheet)

    Dim pCount As Long
    Dim dCount As Long

    pCount = GetProductCount()
    dCount = GetDeptCount()

    UpdateOneTableTotals ws, HOPE_START_ROW, pCount, dCount
    UpdateOneTableTotals ws, GetOrderStartRow(), pCount, dCount

End Sub

Private Sub UpdateOneTableTotals(ws As Worksheet, startRow As Long, pCount As Long, dCount As Long)

    Dim headerRow As Long
    headerRow = startRow + 1

    Dim totalQtyCol As Long
    Dim amountCol As Long

    totalQtyCol = DEPT_START_COL + dCount
    amountCol = totalQtyCol + 1

    Dim i As Long
    Dim r As Long

    For i = 1 To pCount
        r = headerRow + i
        ws.Cells(r, totalQtyCol).Value = SumRow(ws, r, DEPT_START_COL, totalQtyCol - 1)
        ws.Cells(r, amountCol).Value = CLng(Val(ws.Cells(r, PRICE_COL).Value)) * CLng(Val(ws.Cells(r, totalQtyCol).Value))
        ws.Cells(r, amountCol).NumberFormat = MONEY_FORMAT
    Next i

    Dim totalRow As Long
    totalRow = headerRow + pCount + 1

    ws.Cells(totalRow, amountCol).Value = SumCol(ws, headerRow + 1, headerRow + pCount, amountCol)
    ws.Cells(totalRow, amountCol).NumberFormat = MONEY_FORMAT

    ws.Cells(totalRow + 1, amountCol).Value = CLng(Val(ws.Range(BUDGET_CELL).Value)) - CLng(Val(ws.Cells(totalRow, amountCol).Value))
    ws.Cells(totalRow + 1, amountCol).NumberFormat = MONEY_FORMAT

End Sub

Private Sub UpdateStatusAndErrors(ws As Worksheet)

    Dim budget As Long
    Dim orderAmount As Long

    budget = CLng(Val(ws.Range(BUDGET_CELL).Value))
    orderAmount = CalcOrderAmount(ws)

    ws.Cells(14, STATUS_COL + 1).Value = budget
    ws.Cells(15, STATUS_COL + 1).Value = orderAmount
    ws.Cells(16, STATUS_COL + 1).Value = budget - orderAmount

    ws.Range(ws.Cells(14, STATUS_COL + 1), ws.Cells(16, STATUS_COL + 1)).NumberFormat = MONEY_FORMAT

    Dim errMsg As String
    errMsg = ValidateOrderInput(ws, True)

    If errMsg = "" Then
        ws.Cells(19, STATUS_COL).Value = "なし"
    Else
        ws.Cells(19, STATUS_COL).Value = errMsg
    End If

End Sub

Private Sub ShowError(ws As Worksheet, msg As String)
    ws.Cells(19, STATUS_COL).Value = msg
End Sub

'========================
' エラーチェック
'========================
Private Function ValidateMasterNames() As String

    Dim msg As String
    msg = ""

    msg = msg & ValidateNameList(SHEET_SETTING, SET_PRODUCT_START_ROW, SET_PRODUCT_COL, "商品名")
    msg = msg & ValidateNameList(SHEET_SETTING, SET_DEPT_START_ROW, SET_DEPT_COL, "課名")

    ValidateMasterNames = msg

End Function

Private Function ValidateNameList(sheetName As String, startRow As Long, col As Long, labelName As String) As String

    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)

    Dim msg As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim v As String
    Dim foundBlankAfterInput As Boolean
    Dim hasInputAfterBlank As Boolean

    For i = 0 To MAX_ITEMS - 1
        v = Trim(ws.Cells(startRow + i, col).Value)

        If v = "" Then
            If seen.Count > 0 Then foundBlankAfterInput = True
        Else
            If foundBlankAfterInput Then hasInputAfterBlank = True

            If seen.Exists(v) Then
                msg = msg & "・" & labelName & "が重複しています：" & v & vbCrLf
            Else
                seen.Add v, True
            End If
        End If
    Next i

    If seen.Count = 0 Then
        msg = msg & "・" & labelName & "を1つ以上入力してください。" & vbCrLf
    End If

    If hasInputAfterBlank Then
        msg = msg & "・" & labelName & "の途中に空欄があります。上から詰めて入力してください。" & vbCrLf
    End If

    ValidateNameList = msg

End Function

Private Function ValidateHopeInput(ws As Worksheet) As String

    Dim msg As String
    msg = ""

    If Not IsPositiveInteger(ws.Range(BUDGET_CELL).Value) Then
        msg = msg & "・予算は正の整数で入力してください。" & vbCrLf
    End If

    Dim pCount As Long
    Dim dCount As Long

    pCount = GetProductCount()
    dCount = GetDeptCount()

    If pCount = 0 Or dCount = 0 Then
        ValidateHopeInput = "・基本設定シートに商品名・課名を入力し、発注数計算表を作成してください。"
        Exit Function
    End If

    Dim totalQtyCol As Long
    totalQtyCol = DEPT_START_COL + dCount

    Dim headerRow As Long
    headerRow = HOPE_START_ROW + 1

    Dim i As Long
    Dim j As Long
    Dim r As Long

    For i = 1 To pCount
        r = headerRow + i

        If Not IsPositiveInteger(ws.Cells(r, PRICE_COL).Value) Then
            msg = msg & "・" & ws.Cells(r, START_COL).Value & " の単価は正の整数で入力してください。" & vbCrLf
        End If

        For j = DEPT_START_COL To totalQtyCol - 1
            If Not IsZeroOrPositiveInteger(ws.Cells(r, j).Value) Then
                msg = msg & "・" & ws.Cells(r, START_COL).Value & " × " & ws.Cells(headerRow, j).Value & " の希望数は0以上の整数で入力してください。" & vbCrLf
            End If
        Next j
    Next i

    ValidateHopeInput = msg

End Function

Private Function ValidateOrderInput(ws As Worksheet, includeBudgetOver As Boolean) As String

    Dim msg As String
    msg = ""

    If Not IsPositiveInteger(ws.Range(BUDGET_CELL).Value) Then
        msg = msg & "・予算は正の整数で入力してください。" & vbCrLf
    End If

    Dim pCount As Long
    Dim dCount As Long

    pCount = GetProductCount()
    dCount = GetDeptCount()

    If pCount = 0 Or dCount = 0 Then
        ValidateOrderInput = "・商品名・課名を設定してください。"
        Exit Function
    End If

    Dim orderHeader As Long
    orderHeader = GetOrderStartRow() + 1

    Dim totalQtyCol As Long
    totalQtyCol = DEPT_START_COL + dCount

    Dim i As Long
    Dim j As Long
    Dim r As Long

    For i = 1 To pCount
        r = orderHeader + i

        For j = DEPT_START_COL To totalQtyCol
            If Not IsZeroOrPositiveInteger(ws.Cells(r, j).Value) Then
                msg = msg & "・" & ws.Cells(r, START_COL).Value & " の発注数は0以上の整数で入力してください。" & vbCrLf
                Exit For
            End If
        Next j
    Next i

    If includeBudgetOver Then
        If IsPositiveInteger(ws.Range(BUDGET_CELL).Value) Then
            If CalcOrderAmount(ws) > CLng(ws.Range(BUDGET_CELL).Value) Then
                msg = msg & "・発注金額が予算を " & Format(CalcOrderAmount(ws) - CLng(ws.Range(BUDGET_CELL).Value), MONEY_FORMAT) & " 超過しています。" & vbCrLf
            End If
        End If
    End If

    ValidateOrderInput = msg

End Function

Private Function IsPositiveInteger(v As Variant) As Boolean

    IsPositiveInteger = False

    If Trim(CStr(v)) = "" Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CDbl(v) <> Fix(CDbl(v)) Then Exit Function
    If CLng(v) <= 0 Then Exit Function

    IsPositiveInteger = True

End Function

Private Function IsZeroOrPositiveInteger(v As Variant) As Boolean

    IsZeroOrPositiveInteger = False

    If Trim(CStr(v)) = "" Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CDbl(v) <> Fix(CDbl(v)) Then Exit Function
    If CLng(v) < 0 Then Exit Function

    IsZeroOrPositiveInteger = True

End Function

'========================
' 手動変更セル管理
'========================
Public Sub MarkPendingManualCell(ws As Worksheet, target As Range)

    ClearManualMarks ws

    target.Font.Color = RGB(255, 0, 0)

    ws.Range(PENDING_TYPE_CELL).Value = "manual"
    ws.Range(PENDING_ROW_CELL).Value = target.Row
    ws.Range(PENDING_COL_CELL).Value = target.Column

    ws.Cells(19, STATUS_COL).Value = "未反映の手動変更があります。「手動調整・再計算」を押してください。"

End Sub

Private Sub ClearManualMarks(ws As Worksheet)

    Dim pCount As Long
    Dim dCount As Long

    pCount = GetProductCount()
    dCount = GetDeptCount()

    If pCount = 0 Or dCount = 0 Then Exit Sub

    Dim orderHeader As Long
    orderHeader = GetOrderStartRow() + 1

    Dim totalQtyCol As Long
    totalQtyCol = DEPT_START_COL + dCount

    ws.Range(ws.Cells(orderHeader + 1, DEPT_START_COL), ws.Cells(orderHeader + pCount, totalQtyCol)).Font.Color = RGB(0, 0, 0)

End Sub

'========================
' 名前・件数取得
'========================
Private Function GetProductNames() As Variant

    Dim arr(1 To MAX_ITEMS) As String
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTING)

    Dim i As Long
    Dim n As Long

    For i = 0 To MAX_ITEMS - 1
        If Trim(ws.Cells(SET_PRODUCT_START_ROW + i, SET_PRODUCT_COL).Value) <> "" Then
            n = n + 1
            arr(n) = ws.Cells(SET_PRODUCT_START_ROW + i, SET_PRODUCT_COL).Value
        End If
    Next i

    If n = 0 Then n = 1
    ReDim Preserve arr(1 To n)

    GetProductNames = arr

End Function

Private Function GetDeptNames() As Variant

    Dim arr(1 To MAX_ITEMS) As String
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTING)

    Dim i As Long
    Dim n As Long

    For i = 0 To MAX_ITEMS - 1
        If Trim(ws.Cells(SET_DEPT_START_ROW + i, SET_DEPT_COL).Value) <> "" Then
            n = n + 1
            arr(n) = ws.Cells(SET_DEPT_START_ROW + i, SET_DEPT_COL).Value
        End If
    Next i

    If n = 0 Then n = 1
    ReDim Preserve arr(1 To n)

    GetDeptNames = arr

End Function

Public Function GetProductCount() As Long

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTING)

    Dim i As Long
    Dim n As Long

    For i = 0 To MAX_ITEMS - 1
        If Trim(ws.Cells(SET_PRODUCT_START_ROW + i, SET_PRODUCT_COL).Value) <> "" Then
            n = n + 1
        End If
    Next i

    GetProductCount = n

End Function

Public Function GetDeptCount() As Long

    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_SETTING)

    Dim i As Long
    Dim n As Long

    For i = 0 To MAX_ITEMS - 1
        If Trim(ws.Cells(SET_DEPT_START_ROW + i, SET_DEPT_COL).Value) <> "" Then
            n = n + 1
        End If
    Next i

    GetDeptCount = n

End Function

Public Function GetOrderStartRow() As Long
    GetOrderStartRow = HOPE_START_ROW + GetProductCount() + 6
End Function

'========================
' 計算関数
'========================
Private Function SumRow(ws As Worksheet, r As Long, c1 As Long, c2 As Long) As Long

    Dim c As Long
    Dim s As Long

    For c = c1 To c2
        s = s + CLng(Val(ws.Cells(r, c).Value))
    Next c

    SumRow = s

End Function

Private Function SumCol(ws As Worksheet, r1 As Long, r2 As Long, c As Long) As Long

    Dim r As Long
    Dim s As Long

    For r = r1 To r2
        s = s + CLng(Val(ws.Cells(r, c).Value))
    Next r

    SumCol = s

End Function

Private Function CalcHopeAmount(ws As Worksheet) As Long

    Dim pCount As Long
    Dim dCount As Long

    pCount = GetProductCount()
    dCount = GetDeptCount()

    Dim totalQtyCol As Long
    totalQtyCol = DEPT_START_COL + dCount

    Dim headerRow As Long
    headerRow = HOPE_START_ROW + 1

    Dim i As Long
    Dim amount As Long

    For i = 1 To pCount
        amount = amount + CLng(ws.Cells(headerRow + i, PRICE_COL).Value) * _
                           SumRow(ws, headerRow + i, DEPT_START_COL, totalQtyCol - 1)
    Next i

    CalcHopeAmount = amount

End Function

Private Function CalcOrderAmount(ws As Worksheet) As Long

    Dim pCount As Long
    Dim dCount As Long

    pCount = GetProductCount()
    dCount = GetDeptCount()

    If pCount = 0 Or dCount = 0 Then Exit Function

    Dim totalQtyCol As Long
    totalQtyCol = DEPT_START_COL + dCount

    Dim headerRow As Long
    headerRow = GetOrderStartRow() + 1

    Dim i As Long
    Dim amount As Long

    For i = 1 To pCount
        amount = amount + CLng(Val(ws.Cells(headerRow + i, PRICE_COL).Value)) * _
                           SumRow(ws, headerRow + i, DEPT_START_COL, totalQtyCol - 1)
    Next i

    CalcOrderAmount = amount

End Function

Private Function CalcAmountFromTotals(totals() As Long, prices() As Long) As Long

    Dim i As Long
    Dim amount As Long

    For i = LBound(totals) To UBound(totals)
        amount = amount + totals(i) * prices(i)
    Next i

    CalcAmountFromTotals = amount

End Function