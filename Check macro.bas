Option Explicit

Sub Check_report()
'On Error GoTo err
Debug.Print "Begin"
Debug.Print Now


Application.ScreenUpdating = False

    Dim StartTime As Date
    Dim EndTime As Date
    
    StartTime = Now

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
    
    Dim DecNumA As Integer
    Dim DecNumB As Integer
    Dim DecNumAnB As Integer
    Dim DecNumFinalRounding As Integer
    Dim DecNumLLimit As Integer
    Dim DecNumULimit As Integer
    
    DecNumA = 0
    DecNumB = 0
    DecNumLLimit = 4 ' limits rounding to at least 4 digits to avoid issues with percentages
    DecNumULimit = 7

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
    
    'is this really needed? looks like it works without. keeping for now
    'Set varSheetA = wbkA.Worksheets("Cover Region") ' can be any sheet in the file. Needed to asign object to variable
    'Set varSheetB = wbkB.Worksheets("Cover Region") ' can be any sheet in the file. Needed to asign object to variable
    
    CopyBrandlist wbkA, wbkB ' copy list of brands and markets to match report sorting
    
    For iSheet = LBound(varScope, 1) To UBound(varScope, 1)
    
    Debug.Print iSheet & " is " & Now
    
        ActiveSheetForCheck = varScope(iSheet, 0)
        ActiveRangeForCheck = varScope(iSheet, 1)
        
        start_Row = Range(ActiveRangeForCheck).Row
        start_column = Range(ActiveRangeForCheck).Column
    
       
        
        varSheetA = wbkA.Worksheets(ActiveSheetForCheck).Range(ActiveRangeForCheck) ' loads data from check file to excel memory
        
        
       
        varSheetB = wbkB.Worksheets(ActiveSheetForCheck).Range(ActiveRangeForCheck) ' loads data from report to excel memory
             
       
        
        Application.DisplayAlerts = True
    
        For iRow = LBound(varSheetA, 1) To UBound(varSheetA, 1)
            For iCol = LBound(varSheetA, 2) To UBound(varSheetA, 2)
            On Error Resume Next
                If IsNumeric(varSheetA(iRow, iCol)) And IsNumeric(varSheetB(iRow, iCol)) Then
                    
                    DecNumA = ((Len(CStr(varSheetA(iRow, iCol))) - InStr(CStr(varSheetA(iRow, iCol)), ".")))
                    
                    
                   
                    DecNumB = (Len(CStr(varSheetB(iRow, iCol))) - InStr(CStr(varSheetB(iRow, iCol)), "."))
                    
                    'Min function
                    'excel VBA does not have min or max function, hence using workaround
                    If DecNumA < DecNumB Then
                        DecNumAnB = DecNumA
                    Else
                        DecNumAnB = DecNumB
                    End If
                        
                    'Round to the shortest len of number but still within min and max limits
                    Select Case DecNumAnB
                        Case Is < DecNumLLimit
                            DecNumFinalRounding = DecNumLLimit
                        Case Is > DecNumULimit
                            DecNumFinalRounding = DecNumULimit
                        Case Else
                            DecNumFinalRounding = DecNumAnB
                    End Select
                    
                        
                    'rounding variable to the shortest number in comparison (but rounding not more than 4 digits after comma)
                    varSheetA(iRow, iCol) = WorksheetFunction.Round(varSheetA(iRow, iCol), DecNumFinalRounding)
                    varSheetB(iRow, iCol) = WorksheetFunction.Round(varSheetB(iRow, iCol), DecNumFinalRounding)
                    
                End If
                
                If varSheetA(iRow, iCol) <> "[IGNORE]" Then ' Skips marked cells in check file as was intended to be skipped
                
                
                    If varSheetA(iRow, iCol) = varSheetB(iRow, iCol) Then
                        ' Cells are identical.
                        ' Do nothing.
                    Else
                    
                    
                    
                        ' Cells are different.
                        ' Writes to array sheet name, location of difference (A1 cell type, row and column), and source and referrence values
                        errArr(0, UBound(errArr, 2)) = varScope(iSheet, 0) '
                        errArr(1, UBound(errArr, 2)) = Cells(iRow + start_Row - 1, iCol + start_column - 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                        errArr(2, UBound(errArr, 2)) = iRow
                        errArr(3, UBound(errArr, 2)) = iCol
                        errArr(4, UBound(errArr, 2)) = varSheetA(iRow, iCol)
                        errArr(5, UBound(errArr, 2)) = varSheetB(iRow, iCol)
                        'Adds additional line to the array for next addition to array
                        '(VBA does not allow to extend first dimension of array if # of dimension more than 2)
                        'Hence, values added to array horizontally
                        ReDim Preserve errArr(0 To 5, 0 To (UBound(errArr, 2) + 1))
                        
                    End If
                    
                End If
                 'debug
                    If iSheet = 3 And iRow = 24 And iCol = 1 Then
                       ' Stop
                    End If
                    'end debug
            Next iCol
        Next iRow
    Next iSheet
        
    'Clears previous error list and prints from Array "errArr" by transposing it (VBA does not allow to extend 1st dimension of array if # of dimension more than 2)
    Workbooks(wbMacroFile).Activate
    Worksheets(wsMacroFileErrorList).Range("B2", "G65000").Clear
    TransposeAndPrintArray errArr, ActiveWorkbook.Worksheets(wsMacroFileErrorList).[B2]
    Worksheets(wsMacroFileErrorList).Range("B2").Select

Debug.Print "End"
Debug.Print Now

Application.ScreenUpdating = True

EndTime = Now

MsgBox "Job's done." & vbCrLf & "Number of errors found " & UBound(errArr, 2) & vbCrLf & "It took " & WorksheetFunction.Text(EndTime - StartTime, "[hh]:mm:ss"), , "Done"

Done:
    Exit Sub

err:
    MsgBox "Something wrong happened", , "Done"
   

End Sub




Sub PrintArray(Data As Variant, Cl As Range)
    Cl.Resize(UBound(Data, 1) + 1, UBound(Data, 2) + 1) = Data
End Sub

Sub TransposeAndPrintArray(Data As Variant, Cl As Range)
    Dim tData As Variant
    tData = TransposeArray(Data)
    Cl.Resize(UBound(tData, 1), UBound(tData, 2) + 1) = tData
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
    Dim PrintAreaSize As Variant

    reportWB.Activate
    
    X = 0

    For Each ws In Worksheets
        
        varSheetListGrab(X, 0) = ws.Name
        
        If Range(ws.PageSetup.PrintArea).Rows.Count > 1000 Then
            varSheetListGrab(X, 1) = Range(ws.PageSetup.PrintArea).Resize(68).Address
        Else
            varSheetListGrab(X, 1) = ws.PageSetup.PrintArea
        End If
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

Sub browseFilePath(FileType As String)
    On Error GoTo err
    Dim fileExplorer As FileDialog
    Set fileExplorer = Application.FileDialog(msoFileDialogFilePicker)

    'To allow or disable to multi select
    fileExplorer.AllowMultiSelect = False

    With fileExplorer
        If .Show = -1 Then 'Any file is selected
            Select Case FileType
                Case Is = "report"
                    ActiveSheet.Range("E7") = .SelectedItems.Item(1)
                Case Is = "check_file"
                    ActiveSheet.Range("E5") = .SelectedItems.Item(1)
            End Select
             
        Else ' else dialog is cancelled
           ' I am fine if previous value stays
           ' MsgBox "You have cancelled the dialogue"
           ' [filePath] = "" ' when cancelled set blank as file path.
        End If
    End With
err:
    Exit Sub
End Sub

Sub CopyBrandlist(wbkA As Workbook, wbkB As Workbook)

    Dim shTopBrands As String
    Dim shBrands1 As String
    Dim shBrands2 As String
    Dim shBrands3 As String
    
    Dim list As Variant
    
    shTopBrands = "Vol Top 30 Brands"
    shBrands1 = "BF1 T30 Market Group"
    shBrands2 = "BF2 T30 Market Group"
    shBrands3 = "BF3 T30 Market Group"
    
    list = wbkB.Worksheets(shTopBrands).Range("C11:C40").Value
    
    wbkA.Worksheets(shTopBrands).Range("C11:C40") = list
    
    list = ""
    

End Sub

Sub hyperlinks()
'
' Macro4 Macro
'

'
    ActiveSheet.hyperlinks.Add Anchor:=Selection, Address:= _
        "Z:\a. Financial\1. Actual\2018\04_2018\02. Business Update\Volume\EU_region_PMI_report\test_on_March\check_files\EU Shipment Volume Report CHECK file MJ canvas_WITHOUT_spaces_on_area.xlsb" _
        , SubAddress:="Selection!A1", TextToDisplay:="PMI OB %"
End Sub

