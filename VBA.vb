Public oWS_Choice As Variant

Function browseFilePath() As Workbook
    On Error GoTo err
    Dim fileExplorer As FileDialog
    Set fileExplorer = Application.FileDialog(msoFileDialogFilePicker)

    'To allow or disable to multi select
    fileExplorer.AllowMultiSelect = False

    With fileExplorer
        If .Show = -1 Then 'Any file is selected
           Set browseFilePath = Workbooks.Open(.SelectedItems.Item(1))
        Else ' else dialog is cancelled
            MsgBox "You have cancelled the dialogue"
            ' [filePath] = "" ' when cancelled set blank as file path.
        End If
    End With
err:
End Function

Sub AFCAASUM()
    
    Application.ScreenUpdating = False
    
    Dim oWS As Worksheet, nWS As Worksheet
    Dim oWBSRange As Range, oHeadRange As Range, nheadrange As Range
    Dim i As Range, j As Range, oHeadIndex As Range
    Dim oWBSIndex As Long
    Dim oWB As Workbook, nWB As Workbook
    Dim TEST As Range, TEMP As Range
    Dim lBnd As Long, rBnd As Long, nHeadRowIndex As Long, oHeadRowIndex As Long
    
    ' IDENTIFY WORKSHEETS
    Set nWB = ThisWorkbook
    Set oWB = browseFilePath
    Do While (oWB Is Nothing)
        If (MsgBox("Not an acceptable filetype." & vbCrLf & "Do you want to select a new file?", vbYesNo) = vbYes) Then
            Set oWB = browseFilePath
        Else
            Exit Sub
        End If
    Loop
    
    For Each Sheet In oWB.Sheets
        SheetSelection.cbx_SheetList.AddItem (Sheet.Name)
    Next Sheet
    SheetSelection.Show
    Set oWS = oWB.Sheets(oWS_Choice)
    Set nWS = nWB.Sheets("TEMPLATE")
    
    Dim grpHeaders() As Range
    'grpFinder(nWS) arr:=grpheaders
    'If (grpHeaders.Count < 1) Then
    '    MsgBox ("idiot you need to define a group of headers")
    '    Exit Sub
    'End If
    
    ' IDENTIFY THE WBS NUMBER RANGE (oWBSRange)IN THE 1921-1 WORKSHEET (oWS)
    oWS.Activate
    Set TEST = Range("A:AZ").Find("WBS Number", , , , , , True)
    Do While (TEST Is Nothing)
        If (MsgBox("Worksheet does not contain expected data." & vbCrLf & "Do you want to select a new sheet?", vbYesNo) = vbYes) Then
            SheetSelection.Show
            Set oWS = oWB.Sheets(oWS_Choice)
            Set TEST = Range("A:AZ").Find("WBS Number", , , , , , True)
        Else
            Exit Sub
        End If
    Loop
    
    TEST.Select
    oWS.Cells(Selection.Row + 1, Selection.Column).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set oWBSRange = Selection
    Range(Selection.Cells(1, 1), Selection.Cells(Selection.Rows.Count, Selection.Columns.Count + 1)).Select
    Selection.Copy
    
    ' CYCLE THROUGH THE ROWS ON THE 1921-1 TABLE AND THE HEADDERS ON THE TEMPLATE AND PASTE THE APPROPRIATE CONTENT FROM THE ORIGINAL TO THE TEMPLATE
    oWBSIndex = 0
    
    nWS.Activate
    Range("A:AZ").Find("WBS Number", , , , , , True).Select
    nHeadRowIndex = Selection.Row
    nWS.Cells(nHeadRowIndex + 1, Selection.Column).Select
    ActiveSheet.Paste
    
    oWS.Activate
    Range("A:AZ").Find("WBS Number", , , , , , True).Select
    oHeadRowIndex = Selection.Row
    
    For Each i In oWBSRange ' FOR EACH ROW IN THE 1921-1 FIND CONTENT
        oWBSIndex = oWBSIndex + 1
        For Each k In grpFinder(nWS)
            ' IDENTIFY HEADDER RANGE (nHeadRange) IN THE TEMPLATE WORKSHEET (nWS)
            If (Not k Is Nothing) Then
                Set TEMP = Range(k.MergeArea.Address)
                lBnd = TEMP.Cells(1, 1).Column
                rBnd = lBnd + TEMP.Columns.Count - 1
                Set nheadrange = nWS.Range(nWS.Cells(nHeadRowIndex, lBnd), nWS.Cells(nHeadRowIndex, rBnd))
                
                ' IDENTIFY THE GROUP HEADDER RANGE (oGHeadRange) IN THE 1921-1 WORKSHEET (oWS)
                Set TEMP = oWS.Range("a:zz").Find(k.Value, , , , , , True)
                Set TEMP = oWS.Range(TEMP.MergeArea.Address)
                lBnd = TEMP.Cells(1, 1).Column
                rBnd = lBnd + TEMP.Columns.Count - 1
                Set oHeadRange = oWS.Range(oWS.Cells(oHeadRowIndex, lBnd), oWS.Cells(oHeadRowIndex, rBnd))
                
                For Each j In nheadrange ' FOR EACH COLUMN IN THE TEMPLATE COPY THE CONTENT TO THE TEMPLATE
                    With oHeadRange
                        If ((InStr(j.Value, "Total") >= 1 Or InStr(j.Value, "Plus") >= 1)) Then
                            X = 0
                            Index = j.Column - 1
                            leftcol = nheadrange.Columns(1).Column
                            Do
                                If (InStr(nWS.Cells(j.Row, Index).Value, "Hour") = 0) Then
                                    X = X + nWS.Cells(j.Row + oWBSIndex, Index).Value
                                End If
                                If (InStr(nWS.Cells(j.Row, Index).Value, "Total") >= 1 _
                                    Or InStr(nWS.Cells(j.Row, Index).Value, "Plus")) Then
                                    Exit Do
                                End If
                                Index = Index - 1
                            Loop While (Index >= leftcol)
                            nWS.Cells(j.Row + oWBSIndex, j.Column).Value = IIf(X = 0, "", X)
                        Else
                            Set oHeadIndex = .Find(j.Value, , , , , , True)
                            If (Not oHeadIndex Is Nothing) Then
                                If InStr(1, j.Value, "WBS") = 1 Then
                                    nWS.Cells(j.Row + oWBSIndex, j.Column).Value = oWS.Cells(i.Row, oHeadIndex.Column).Value
                                Else
                                    X = 0
                                    y = oHeadIndex.Address
                                    Do
                                        X = X + oWS.Cells(i.Row, oHeadIndex.Column).Value
                                        Set oHeadIndex = .FindNext(oHeadIndex)
                                    Loop While (Not oHeadIndex Is Nothing And oHeadIndex.Address <> y)
                                    nWS.Cells(j.Row + oWBSIndex, j.Column).Value = IIf(X = 0, "", X)
                                End If
                            End If
                        End If
                    End With
                Next j
            End If
        Next k
    Next i
    
    For Each k In grpFinder(nWS)
        If Not (k Is Nothing) Then
            Set TEMP = Range(k.MergeArea.Address)
            lBnd = TEMP.Cells(1, 1).Column
            rBnd = lBnd + TEMP.Columns.Count - 1
            Set nheadrange = nWS.Range(nWS.Cells(nHeadRowIndex, lBnd), nWS.Cells(nHeadRowIndex, rBnd))
                
            For Each j In nheadrange
                If ((InStr(j.Value, "Total") >= 1 Or InStr(j.Value, "Plus") >= 1) And _
                Not (InStr(j.Value, "Observed") >= 1 Or InStr(j.Value, "Expected") >= 1)) Then
                    j.EntireColumn.Insert xlShiftToRight, xlFormatFromRightOrBelow
                    nWS.Cells(j.Row, j.Column - 1).Value = "Observed " & j.Text
                    nWS.Cells(j.Row, j.Column - 1).EntireColumn.AutoFit
                    
                    Set TEMP = oWS.Range("a:zz").Find(k.Value, , , , , , True)
                    Set TEMP = oWS.Range(TEMP.MergeArea.Address)
                    lBnd = TEMP.Cells(1, 1).Column
                    rBnd = lBnd + TEMP.Columns.Count - 1
                    Set oHeadRange = oWS.Range(oWS.Cells(oHeadRowIndex, lBnd), oWS.Cells(oHeadRowIndex, rBnd))
                    With oWBSRange
                        X = .Rows(.Rows.Count).Row
                    End With
                    y = oHeadRange.Find(j.Value).Column
                    oWS.Activate
                    oWS.Range(oWS.Cells(oHeadRange.Row + 1, y), oWS.Cells(X, y)).Select
                    Selection.Copy
                    
                    nWS.Activate
                    nWS.Cells(j.Row + 1, j.Column - 1).Select
                    ActiveSheet.Paste
                    
                    j.Value = "Expected " & j.Text
                    j.EntireColumn.AutoFit
                End If
            Next j
        End If
    Next k
    
    ' IDENTIFY THE WBS RANGE (nWBSRange) IN THE TEMPLATE WORKSHEET (nWS) AND LEFT JUSTIFY IT
    nWS.Activate
    nWS.Name = "1921-1 Summary"
    nWS.Range("A:AZ").Find("WBS Number", , , , , , True).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Set nheadrange = Selection
    
    For Each j In nheadrange
        j.Value = Replace(j.Value, "* ", " ")
    Next j
    
    nWS.Copy After:=oWB.Sheets(oWS_Choice)
    
    
    Application.ScreenUpdating = True
    
    nWB.Close False

End Sub

Function grpFinder(WS As Worksheet) As Range()
    Dim rg As Range
    Dim grpHeaders() As Range
    Dim X As Integer
    
    X = 0
    ReDim Preserve grpHeaders(X)
    ' IDENTIFY NUMBER OF GROUPS
    
    ' GET THE RANGE FOR EACH GROUP FROM THE TEMPLATE
    WS.Activate
    Set rg = WS.Range("GRP_HEADER")
    If (rg Is Nothing) Then
        Exit Function
    End If
    For Each i In rg
        If (i.Value <> "") Then
            Set grpHeaders(X) = WS.Range("A:AZ").Find(i.Value, , , , , , True)
            X = X + 1
            ReDim Preserve grpHeaders(X)
        End If
    Next i
    grpFinder = grpHeaders
End Function

