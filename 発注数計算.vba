Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ExitHandler

    If Target.CountLarge > 1 Then Exit Sub

    Dim pCount As Long
    Dim dCount As Long

    pCount = GetProductCount()
    dCount = GetDeptCount()

    If pCount = 0 Or dCount = 0 Then Exit Sub

    Dim orderHeader As Long
    orderHeader = GetOrderStartRow() + 1

    Dim totalQtyCol As Long
    totalQtyCol = DEPT_START_COL + dCount

    Dim editableRange As Range
    Set editableRange = Me.Range( _
        Me.Cells(orderHeader + 1, DEPT_START_COL), _
        Me.Cells(orderHeader + pCount, totalQtyCol) _
    )

    If Intersect(Target, editableRange) Is Nothing Then Exit Sub

    Application.EnableEvents = False

    MarkPendingManualCell Me, Target

ExitHandler:
    Application.EnableEvents = True

End Sub