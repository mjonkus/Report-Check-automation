Option Explicit

Sub Check_report()

Debug.Print "Begin"
Debug.Print Now


Application.ScreenUpdating = False

    Dim varSheetA As Variant
    Dim varSheetB As Variant
    Dim varSheetList As Variant
    Dim varRangetoCheck As Variant
    Dim varScope As Variant
    
    Dim SheetListNumber As Integer
    
    Dim errArr() As Variant
    ReDim errArr(0 To 5, 0 To 0)
    
        
    Dim pathA As String
    Dim pathB As String
    Dim iRow As Long
    Dim iCol As Long
    
    Dim iSheet As Long
    Dim iRange As String
    
    Dim wbkA As Workbook
    Dim wbkB As Workbook
    
    Dim start_Row As Integer
    Dim start_column As Integer
    
    Dim wbMacroFile As String
    Dim wsMacroFileSetup As String
    Dim wsMacroFileErrorList As String
    
    Dim ActiveSheetForCheck As String
    Dim ActiveRangeForCheck As String

    wbMacroFile = "CHECK_macro"
    wsMacroFileSetup = "Macro_setup"
    wsMacroFileErrorList = "Error_list"
    
    Workbooks(wbMacroFile).Activate

    pathA = Worksheets(wsMacroFileSetup).Range("E5").Value
    pathB = Worksheets(wsMacroFileSetup).Range("E7").Value

    
    Application.DisplayAlerts = False
    
    Set wbkA = Workbooks.Open(Filename:=pathA)
    Set wbkB = Workbooks.Open(Filename:=pathB)
    
    SheetListNumber = DimSheetListArray(wbkB, wbMacroFile)
    
    varScope = GetSheetList(wbkB, wbMacroFile, SheetListNumber)

    PrintArray varScope, Workbooks(wbMacroFile).Worksheets(wsMacroFileSetup).[D11]
    
    Set varSheetA = wbkA.Worksheets("Cover Region") ' or whatever sheet you need
    Set varSheetB = wbkB.Worksheets("Cover Region") ' or whatever sheet you need
    
    For iSheet = LBound(varScope, 1) To UBound(varScope, 1)
    
    Debug.Print iSheet & " is " & Now
    
        ActiveSheetForCheck = varScope(iSheet, 0)
        ActiveRangeForCheck = varScope(iSheet, 1)
        
        start_Row = Range(ActiveRangeForCheck).Row
        start_column = Range(ActiveRangeForCheck).Column
    
       ' Set varSheetA = wbkA.Worksheets(varScope(iSheet, 0)) ' or whatever sheet you need
        
        varSheetA = wbkA.Worksheets(ActiveSheetForCheck).Range(ActiveRangeForCheck)
        
        
       ' Set varSheetB = wbkB.Worksheets(varScope(iSheet, 0)) ' or whatever sheet you need
       
        varSheetB = wbkB.Worksheets(ActiveSheetForCheck).Range(ActiveRangeForCheck) ' or whatever your other sheet is.
             
        
        
        Application.DisplayAlerts = True
    
        For iRow = LBound(varSheetA, 1) To UBound(varSheetA, 1)
            For iCol = LBound(varSheetA, 2) To UBound(varSheetA, 2)
                If IsNumeric(varSheetA(iRow, iCol)) Then
                    varSheetA(iRow, iCol) = Round(varSheetA(iRow, iCol), 8)
                End If
                
                'On Error Resume Next
                If IsNumeric(varSheetB(iRow, iCol)) Then
                
                    varSheetB(iRow, iCol) = Round(varSheetB(iRow, iCol), 8)
                End If
                
                
                If varSheetA(iRow, iCol) = varSheetB(iRow, iCol) Then
                    ' Cells are identical.
                    ' Do nothing.
                Else
                    ' Cells are different.
                    ' Code goes here for whatever it is you want to do.
                    
                    errArr(0, UBound(errArr, 2)) = varScope(iSheet, 1)
                    errArr(1, UBound(errArr, 2)) = Cells(iRow + start_Row, iCol + start_column).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    errArr(2, UBound(errArr, 2)) = iRow
                    errArr(3, UBound(errArr, 2)) = iCol
                    errArr(4, UBound(errArr, 2)) = varSheetA(iRow, iCol)
                    errArr(5, UBound(errArr, 2)) = varSheetB(iRow, iCol)
                    ReDim Preserve errArr(0 To 5, 0 To (UBound(errArr, 2) + 1))
                   
                End If
            Next iCol
        Next iRow
    Next iSheet
        
    
    Workbooks(wbMacroFile).Activate
    Worksheets(wsMacroFileErrorList).Range("B2", Worksheets(wsMacroFileErrorList).Range("F2").End(xlDown)).Clear
    TransposeAndPrintArray errArr, ActiveWorkbook.Worksheets(wsMacroFileErrorList).[B2]

Debug.Print "End"
Debug.Print Now

Application.ScreenUpdating = True

End Sub


Sub PrintArray(Data As Variant, Cl As Range)
    Cl.Resize(UBound(Data, 1) + 1, UBound(Data, 2) + 1) = Data
End Sub

Sub TransposeAndPrintArray(Data As Variant, Cl As Range)
    Dim tData As Variant
    tData = TransposeArray(Data)
    Cl.Resize(UBound(tData, 1), UBound(tData, 2)) = tData
End Sub

Public Function TransposeArray(myarray As Variant) As Variant
Dim X As Long
Dim Y As Long
Dim Xupper As Long
Dim Yupper As Long
Dim tempArray As Variant
    Xupper = UBound(myarray, 2)
    Yupper = UBound(myarray, 1)
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = myarray(Y, X)
        Next Y
    Next X
    TransposeArray = tempArray
End Function

Function GetSheetList(reportWB As Workbook, wb As String, SheetListNumber As Integer) As Variant
    
    Dim ws As Worksheet
    Dim X As Integer
    Dim varSheetListGrab() As Variant
    ReDim varSheetListGrab(0 To SheetListNumber, 0 To 1)

    reportWB.Activate
    
    X = 0

    For Each ws In Worksheets
        
        varSheetListGrab(X, 0) = ws.Name
        varSheetListGrab(X, 1) = ws.PageSetup.PrintArea
        X = X + 1

    Next ws
    
    GetSheetList = varSheetListGrab

End Function

Function DimSheetListArray(reportWB As Workbook, wb As String) As Integer
    
    Dim ws As Worksheet
    Dim X As Integer
    
    reportWB.Activate
    
    X = -1

    For Each ws In Worksheets
        X = X + 1

    Next ws
    
    DimSheetListArray = X

End Function

Sub testArray()

Dim arrTest() As Variant
ReDim Preserve arrTest(0 To 1)
Dim i As Integer


For i = 0 To 5

    arrTest(i) = "test"
    
    Debug.Print arrTest(i)
Next i

End Sub

Sub testPrintAreaSelect()

    ActiveSheet.Range(ActiveSheet.PageSetup.PrintArea).Select

End Sub

Sub ErrorTest()
    
    Dim dblValue        As Double
      
    On Error GoTo ErrHandler1
    dblValue = 1 / 0
ErrHandler1:
    MsgBox "Exception Caught"
    On Error GoTo 0           'Comment this line to check the effect
    On Error GoTo ErrHandler2
    dblValue = 1 / 0
ErrHandler2:
    MsgBox "Again caught it."
        
End Sub
