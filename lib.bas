Attribute VB_Name = "lib"
Public NameCol, TaskCol, TimeCol, DateCol, MaterialCol, FirstLine, CodeCol, CostCol, PrjNameL, PrjNameC, PrjCodeL, PrjCodeC, FirstCol, RemarksCol, DateCelL, DateCelC As Integer
Public Sub OpenFR()
    If StartUI.Visible = False Then
        'Load StartUI
        'Application.OnTime Now + TimeValue("00:00:01"), "Rectangle1_Click"
        StartUI.Show
        'StartUI.Enabled = False
        'Unload StartUI
    End If
End Sub
Sub test()
    If StartUI.Visible = False Then
        StartUI.Show
        'Unload StartUI
    End If
End Sub
Sub SubSD(ByVal fd As Date, ByVal td As Date)
    Dim pctDone As Single
    Dim fromD, toD As Date
    Dim tsk As Task
    Dim rs As Resources
    Dim r As Resource
    Dim a As Assignment
    Dim TSV As TimeScaleValue
    Dim TSVS As TimeScaleValues
    'MsgBox ActiveProject.Name
    Dim fname As String
    Dim RunExcel, FirstTime As Boolean
    FirstTime = True
    RunExcel = False
    If Not IsAppRunning("Excel.Application") Then
        cPath = Application.ActiveProject.Path
        'x = Shell(cPath, vbNormalFocus)
        'CreateObject ("Excel.Application")
        RunExcel = True
    End If
    Dim xlApl As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlWs As Excel.Worksheet
    Set xlApl = New Excel.Application
    'xlApl.Visible = False
    'Application.DisplayAlerts = False
    'Application.ScreenUpdating = False
    Dim xlLogWb As Excel.Workbook
    Dim xlLogWs As Excel.Worksheet
    Set xlLogWb = xlApl.Workbooks.Open(Application.ActiveProject.Path & "\Log.xlsx", True, False)
    Set xlLogWs = xlLogWb.Worksheets("Log")
    
    FinalRow = xlLogWs.Cells(xlLogWs.Rows.Count, 1).End(xlUp).Row
    
    Set xlWb = xlApl.Workbooks.Open(Application.ActiveProject.Path & "\Start_Initial.xlsx", True, False)
    Set xlWs = xlWb.Worksheets("Params")
        FirstLine = xlWs.Cells(1, 2)
        NameCol = xlWs.Cells(2, 2)
        TaskCol = xlWs.Cells(3, 2)
        TimeCol = xlWs.Cells(4, 2)
        DateCol = xlWs.Cells(5, 2)
        MaterialCol = xlWs.Cells(6, 2)
        CostCode = xlWs.Cells(7, 2)
        CodeCol = xlWs.Cells(8, 2)
        PrjNameL = xlWs.Cells(9, 2)
        PrjNameC = xlWs.Cells(9, 3)
        PrjCodeL = xlWs.Cells(10, 2)
        PrjCodeC = xlWs.Cells(10, 3)
        FirstCol = xlWs.Cells(11, 2)
        RemarksCol = xlWs.Cells(12, 2)
        DateCelL = xlWs.Cells(13, 2)
        DateCelC = xlWs.Cells(13, 3)
    Set xlWs = xlWb.Worksheets(1)
    
    fromD = Format(Now(), "dd/mm/yyyy")
    toD = Format(Now(), "dd/mm/yyyy")
    
    fromD = fd
    toD = td
    'On Error GoTo ExSub
    cPrj = Application.ActiveProject.CustomDocumentProperties("project")
    GoTo ContSub
ExSub:
    ans = MsgBox("Error: Assign a code to project - " & Err & Error(Err), vbCritical, "ERROR")
    Exit Sub
ContSub:
    'fr = xlWs.Cells(xl.Rows.Count, 1).End(xlUp).Row
    Do While fromD <= toD
        If Not ActiveProject.Calendar.WeekDays(Weekday(fromD)).Working Then
            ans = MsgBox("Non working day " & fromD & " - " & Format(fromD, "dddd") & vbLf _
                & "Ok:Continue to next working day, Cancel:Create file for the day", vbOKCancel, "Select")
            If ans = 1 Then
                GoTo nextD
            End If
        End If
        sname = cPrj & "_" & Format(fromD, "dd_mm_yyyy") & ".xlsx"
        fname = Application.ActiveProject.Path & "\" & sname
        'ActiveProject.CurrentDate = fromD
        'ActiveProject.Calendar.WeekDays.Item.Working
        'If ActiveProject.Calendar.WeekDays(Weekday(fromD)).Working = False Then
        '    GoTo nextD
        'End If
        If Dir(fname) <> "" Then
           ans = MsgBox("The file " & fname & " allready exists" & crLF & "Try to Update or delete", vbCritical, "Error")
           GoTo nextD
        End If
        xlWb.SaveAs (fname)
        
        Call upd_status(sname)

        
        FR = xlWs.Cells(xlWs.Rows.Count, FirstCol).End(xlUp).Row
        FC = xlWs.Cells(FR, xlWs.Columns.Count).End(xlToLeft).Column
        xlWs.Range(xlWs.Cells(FR + 1, FC), xlWs.Cells(FirstLine + 1, FirstCol)).UnMerge
        xlWs.Range(xlWs.Cells(FR + 1, FC), xlWs.Cells(FirstLine + 1, FirstCol)).ClearContents
        Set rs = ActiveProject.Resources
        L = FirstLine
        m = FirstLine
        C = FirstLine
        m_sn = 0
        w_sn = 0
        counter = 0
        For Each r In rs
                counter = counter + 1
                pctDone = counter / rs.Count
                UpdateProgressBar pctDone
                
                 For Each a In r.Assignments
                    onoma = Mid(a.Project, InStrRev(a.Project, "\") + 1, Len(a.Project) - InStrRev(a.Project, "\"))
                    xlWs.Cells(PrjNameL, PrjNameC).Value = onoma
                    xlWs.Cells(PrjCodeL, PrjCodeC).Value = cPrj
                    xlWs.Cells(DateCelL, DateCelC).Value = fromD
                    If (a.Project = onoma) Then
                        If r.Type = pjResourceTypeWork Then
                            'Set TSVS = a.TimeScaleData(fromD, fromD, Type:=pjAssignmentTimescaledWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                            For Each TSV In a.TimeScaleData(fromD, fromD, Type:=pjAssignmentTimescaledWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                             If TSV.Value <> "" And TSV.Value > "0" Then
                                L = L + 1
                                w_sn = w_sn + 1
                                'xlWs.Cells(l, 1).Value = l - 7
                                xlWs.Cells(L, NameCol - 2).Value = w_sn
                                xlWs.Cells(L, NameCol - 1).Value = a.Task.OutlineParent.Name
                                xlWs.Cells(L, NameCol).Value = r.Name
                                xlWs.Cells(L, DateCol).Value = fromD 'TSV.StartDate
                                xlWs.Cells(L, TimeCol).Value = TSV.Value / 60
                                xlWs.Cells(L, TaskCol).Value = a.TaskName
                                
                                xlWs.Cells(L, 101).Value = a.TaskID
                                xlWs.Cells(L, 102).Value = r.ID
                                xlWs.Cells(L, 103).Value = a.UniqueID
                                xlWs.Cells(L, 104).Value = TSV.Value / 60
                                
                             End If
                            Next TSV
                        ElseIf r.Type = pjResourceTypeMaterial Then
                          For Each TSV In a.TimeScaleData(fromD, fromD, Type:=pjAssignmentTimescaledWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                            If TSV.Value <> "" And TSV.Value > "0" Then
                                m = m + 1
                                m_sn = m_sn + 1
                                xlWs.Cells(m, NameCol + MaterialCol - 1).Value = m_sn
                                xlWs.Cells(m, NameCol + MaterialCol).Value = r.Name
                                xlWs.Cells(m, DateCol + MaterialCol).Value = fromD 'TSV.StartDate
                                'xlWs.Cells(m, TimeCol + MaterialCol).Value = TSV.Value / 60
                                xlWs.Cells(m, TimeCol + MaterialCol).Value = r.Code
                                xlWs.Cells(m, TaskCol + MaterialCol).Value = a.TaskName
                                
                                xlWs.Cells(m, 111).Value = a.TaskID
                                xlWs.Cells(m, 112).Value = r.ID
                                xlWs.Cells(m, 113).Value = a.UniqueID
                                xlWs.Cells(m, 114).Value = TSV.Value
                             End If
                           Next TSV
                        Else
                           For Each TSV In a.TimeScaleData(fromD, fromD, Type:=pjAssignmentTimescaledWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                            If TSV.Value <> "" And TSV.Value > "0" Then
                                C = C + 1
                                xlWs.Cells(C, NameCol + CostCol).Value = r.Name
                                xlWs.Cells(C, DateCol + CostCol).Value = fromD 'TSV.StartDate
                                xlWs.Cells(C, TimeCol + CostCol).Value = TSV.Value
                                xlWs.Cells(C, TaskCol + CostCol).Value = a.TaskName
                                
                                xlWs.Cells(C, 121).Value = a.TaskID
                                xlWs.Cells(C, 122).Value = r.ID
                                xlWs.Cells(C, 123).Value = a.UniqueID
                                xlWs.Cells(C, 124).Value = TSV.Value / 60
                             End If
                           Next TSV
                        End If
                      End If
                 Next a
        Next r
        If w_sn > 0 Then
            Call SortbyTasks(xlWb.Worksheets(1)) ' sort edw na mhn xalaei ta mhxanhmata
        End If
        If FirstTime = True Then
            FirstTime = False
            'Add sheet Resources Assign Name for Resources
            Set xlWsN = xlWb.Sheets.Add(after:= _
                     xlWb.Sheets(xlWb.Sheets.Count))
            xlWsN.Name = "Resources"
            Set xlCell = xlWsN.Range("A1")
            Set rs = ActiveProject.Resources
            Lines = 0
            For Each r In rs
                If (Not r Is Nothing) And r.Type = pjResourceTypeWork Then
                        xlCell.Value = r.Name
                        xlCell.Offset(0, 1).Value = r.ID
                        Set xlCell = xlCell.Offset(1, 0)
                        Lines = Lines + 1
                End If
            Next r
            'xlWsN.Range("A1:A" & Lines & ")").Select
            xlWb.Names.Add Name:="Resources", RefersToR1C1:= _
                "=Resources!R1C1:R" & Lines & "C1"
            xlWb.Names.Add Name:="ResourcesAll", RefersToR1C1:= _
                "=Resources!R1C1:R" & Lines & "C2"
            'Add sheet Tasks Assign Name for Tasks
            Set xlWsN = xlWb.Sheets.Add(after:= _
                     xlWb.Sheets(xlWb.Sheets.Count))
            xlWsN.Name = "Tasks"
            Set xlCell = xlWsN.Range("A1")
            Set Project_Tasks = ActiveProject.Tasks
            Lines = 0
            For Each tsk In ActiveProject.Tasks
                    If (Not tsk Is Nothing) Then 'And t.OutlineLevel = 1 Then
                        If tsk.OutlineChildren.Count = 0 And tsk.Duration <> 0 Then
                            xlCell.Value = tsk.Name
                            xlCell.Offset(0, 1).Value = tsk.ID
                            xlCell.Offset(0, 2).Value = tsk.OutlineParent.Name
                            'xlCell.Value = Project_Task.OutlineLevel
                            Set xlCell = xlCell.Offset(1, 0)
                            Lines = Lines + 1
                         End If
                    End If
            Next tsk
            xlWb.Names.Add Name:="Tasks", RefersToR1C1:= _
                "=Tasks!R1C1:R" & Lines & "C1"
            xlWb.Names.Add Name:="TasksAll", RefersToR1C1:= _
                "=Tasks!R1C1:R" & Lines & "C3"
            'Add sheet Material Assign Name for Material
            Set xlWsN = xlWb.Sheets.Add(after:= _
                     xlWb.Sheets(xlWb.Sheets.Count))
            xlWsN.Name = "Material"
            Set xlCell = xlWsN.Range("A1")
            Set rs = ActiveProject.Resources
            Lines = 0
            For Each r In rs
                If (Not r Is Nothing) And r.Type = pjResourceTypeMaterial Then
                        xlCell.Value = r.Name
                        xlCell.Offset(0, 1).Value = r.ID
                        Set xlCell = xlCell.Offset(1, 0)
                        Lines = Lines + 1
                End If
            Next r
            'xlWsN.Range("A1:A" & Lines & ")").Select
            xlWb.Names.Add Name:="Material", RefersToR1C1:= _
                "=Material!R1C1:R" & Lines & "C1"
            xlWb.Names.Add Name:="MaterialAll", RefersToR1C1:= _
                "=Material!R1C1:R" & Lines & "C2"
                
            'xlWb.Sheets("Resources").Visible = False
            'xlWb.Sheets("Tasks").Visible = False
            'xlWb.Sheets("Material").Visible = False
       End If
            If L > 0 Then
                Call Valid(xlWb.Worksheets(1), 20)
            Else
                Call ValidN(xlWb.Worksheets(1), 20)
            End If
            xlWb.Save
            FinalRow = FinalRow + 1
            xlLogWs.Cells(FinalRow, 1).Value = sname
            xlLogWs.Cells(FinalRow, 2).Value = fromD
            xlLogWs.Cells(FinalRow, 3).Value = fname
            xlLogWs.Cells(FinalRow, 4).Value = Now()
        
        'xlWb.Close
        'xlApl.Visible = True
nextD:
        fromD = DateAdd("d", 1, fromD)
    Loop
        xlWb.Save
        For Each nm In xlLogWb.Names
            nm.Delete
        Next
        xlLogWb.Save
        'xlApl.DisplayAlerts = True
        'xlApl.DisplayAlerts = True
        'xlApl.ScreenUpdating = True
        If Not xlLogWb Is Nothing Then
            xlLogWb.Close
            Set xlLogWb = Nothing
        End If
        If Not xlWb Is Nothing Then
            xlWb.Close
            Set xlWb = Nothing
        End If
        If Not RunExcel Then
            xlApl.Quit
            Set xlApl = Nothing
        Else
            xlApl.DisplayAlerts = True
            xlApl.DisplayAlerts = True
            xlApl.ScreenUpdating = True
            xlApl.WindowState = xlMinimized
            xlApl.Visible = True
        End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    UpdateProgressBar 0
    MsgBox "Finish....."
End Sub
Sub Upd(ByVal fd As Date, ByVal td As Date)
    Dim pctDone As Single
    Dim fromD, toD, UpdDate As Date
    Dim rs As Resources
    Dim r As Resource
    Dim tsk As Task
    Dim tsks As Tasks
    Dim a, ass As Assignment
    Dim TSV As TimeScaleValue
    Dim TSVS As TimeScaleValues
    'MsgBox ActiveProject.Name
    Dim fname As String

    Dim RunExcel As Boolean
    RunExcel = False
    If Not IsAppRunning("Excel.Application") Then
        cPath = Application.ActiveProject.Path
        RunExcel = True
    End If
    Dim xlApl As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlWs As Excel.Worksheet
    Dim xlWsN As Excel.Worksheet
    Set xlApl = New Excel.Application
    'xlApl.Visible = False
    'Application.DisplayAlerts = False
    'Application.ScreenUpdating = False
    Dim xlLogWb As Excel.Workbook
    Dim xlLogWs, xlLogAll As Excel.Worksheet
    
        Set xlWb = xlApl.Workbooks.Open(Application.ActiveProject.Path & "\Start_Initial.xlsx", True, False)
        Set xlWs = xlWb.Worksheets("Params")
        FirstLine = xlWs.Cells(1, 2)
        NameCol = xlWs.Cells(2, 2)
        TaskCol = xlWs.Cells(3, 2)
        TimeCol = xlWs.Cells(4, 2)
        DateCol = xlWs.Cells(5, 2)
        MaterialCol = xlWs.Cells(6, 2)
        CostCode = xlWs.Cells(7, 2)
        CodeCol = xlWs.Cells(8, 2)
        PrjNameL = xlWs.Cells(9, 2)
        PrjNameC = xlWs.Cells(9, 3)
        PrjCodeL = xlWs.Cells(10, 2)
        PrjCodeC = xlWs.Cells(10, 3)
        FirstCol = xlWs.Cells(11, 2)
        RemarksCol = xlWs.Cells(12, 2)
        DateCelL = xlWs.Cells(13, 2)
        DateCelC = xlWs.Cells(13, 3)
        UpdDate = Now()
        xlWb.Close
        Set xlWb = Nothing
    Set xlLogWb = xlApl.Workbooks.Open(Application.ActiveProject.Path & "\Log.xlsx", True, False)
    Set xlLogAll = xlLogWb.Worksheets("All")
    Set xlLogWs = xlLogWb.Worksheets("Log")
    FinalRow = xlLogWs.Cells(xlLogWs.Rows.Count, 1).End(xlUp).Row
    
    xlLogWs.Sort.SortFields.Clear
    xlLogWs.Sort.SortFields.Add Key:=xlLogWs.Range("B2:B" & FinalRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With xlLogWs.Sort
        .SetRange xlLogWs.Range("A1:G" & FinalRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    fromD = Format(Now(), "dd/mm/yyyy")
    toD = Format(Now(), "dd/mm/yyyy")
    fromD = fd
    toD = td
    L = 2
    Do While xlLogWs.Cells(L, 2).Value < fd
       L = L + 1
    Loop
    counter = 0
    ola = IIf(toD > fd, toD - fd, 1)
    Do While L <= FinalRow And fromD >= fd And fromD <= toD
        'If xlLogWs.Cells(L, 2).Value < fromD And xlLogWs.Cells(L, 2).Value > toD Then
        '    GoTo cont
        'End If
        'If Not (xlLogWs.Cells(L, 5).Value = 1 And xlLogWs.Cells(L, 6).Value <> "") Then
        '    GoTo cont
        'End If
        If Not ActiveProject.Calendar.WeekDays(Weekday(fromD)).Working Then
         If Not (Dir(xlLogWs.Cells(L, 3).Value) <> "") Then
            'MsgBox "Non working day " & fromD & " - " & Format(fromD, "dddd")
            'GoTo Cont
         End If
        End If
        If Not (Dir(xlLogWs.Cells(L, 3).Value) <> "") Then
            MsgBox "The file " & xlLogWs.Cells(L, 3).Value & " was not found"
            GoTo Cont
        Else
            xlLogWs.Cells(L, 5).Value = "Upd"
            xlLogWs.Cells(L, 6).Value = UpdDate
        End If
        Set xlWb = xlApl.Workbooks.Open(xlLogWs.Cells(L, 3).Value, True, False)
        Set xlWs = xlWb.Worksheets(1)
        
        Call upd_status(xlLogWs.Cells(L, 3).Value)
        
        frow = FirstLine + 1
        FR = xlWs.Cells(xlWs.Rows.Count, NameCol).End(xlUp).Row
        Set rs = ActiveProject.Resources
        Set tsks = ActiveProject.Tasks
        
        counter = IIf(counter < (toD - fromD + 1), counter + 1, counter)
        pctDone = counter / ola
        UpdateProgressBar pctDone
        
        For i = frow To FR
            'ActiveProject.Tasks.UniqueID(xlWs.Cells(i, 101)).Assignments.UniqueID(xlWs.Cells(i, 103)).TimeScaleData(StartDate:=xlWs.Cells(i, 7), EndDate:=xlWs.Cells(i, 7), Type:=pjAssignmentTimescaledActualWork, TimeScaleUnit:=pjTimescaleDays, Count:=1).Item(1).Value = xlWs.Cells(i, 14).Value
            fromD = Format(xlLogWs.Cells(L, 2).Value, "dd/mm/yyyy") ' xlWs.Cells(i, DateCol)
            'If xlWs.Cells(i, 103) <> "" Then
            '    For Each r In rs
            '      If (r.Type = pjResourceTypeWork) And (r.Name = xlWs.Cells(i, NameCol).Value) Then '(r.ID = xlWs.Cells(i, 102).Value) Then
            '           For Each a In r.Assignments
            '               onoma = Mid(a.Project, InStrRev(a.Project, "\") + 1, Len(a.Project) - InStrRev(a.Project, "\"))
            '               If a.Project = onoma And a.TaskName = xlWs.Cells(i, TaskCol) Then 'a.TaskID = xlWs.Cells(i, 101) Then
            '                 For Each TSV In a.TimeScaleData(StartDate:=fromD, EndDate:=fromD, Type:=pjAssignmentTimescaledActualWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
            '                    TSV.Value = Val(TSV.Value) + xlWs.Cells(i, TimeCol).Value * 60
            '                    fnd = True
            '                 Next TSV
            '               End If
            '           Next a
            '       End If
            '    Next r
            'Else
                For Each r In rs
                    If r.Name = xlWs.Cells(i, NameCol).Value Then
                        For Each tsk In tsks
                            If tsk.Name = xlWs.Cells(i, TaskCol).Value Then
                                Set a = Nothing
                                For Each ass In tsk.Assignments
                                    If ass.ResourceName = r.Name Then
                                       Set a = ass
                                    End If
                                Next
                                If a Is Nothing Then
                                    Set a = tsk.Assignments.Add(ResourceID:=r.ID, Units:=1) 'Units:=(xlWs.Cells(i, TimeCol).Value / 8))
                                End If
                                For Each TSV In a.TimeScaleData(StartDate:=fromD, EndDate:=fromD, Type:=pjAssignmentTimescaledActualWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                                   TSV.Value = Val(TSV.Value) + xlWs.Cells(i, TimeCol).Value * 60
                                   fnd = True
                                Next TSV
                               Exit For
                           End If
                        Next
                        'a.Assignments.Add TaskID:=xlWs.Cells(i, 101).Value, ResourceID:=xlWs.Cells(i, 102).Value, Units:=xlWs.Cells(i, NameCol).Value
                    End If
                Next
           'End If
        Next i
        'Materials
        For i = frow To FR
            'ActiveProject.Tasks.UniqueID(xlWs.Cells(i, 101)).Assignments.UniqueID(xlWs.Cells(i, 103)).TimeScaleData(StartDate:=xlWs.Cells(i, 7), EndDate:=xlWs.Cells(i, 7), Type:=pjAssignmentTimescaledActualWork, TimeScaleUnit:=pjTimescaleDays, Count:=1).Item(1).Value = xlWs.Cells(i, 14).Value
            fromD = Format(xlLogWs.Cells(L, 2).Value, "dd/mm/yyyy")
            If xlWs.Cells(i, NameCol + MaterialCol) <> "" Then
                For Each r In rs
                  If (r.Type = pjResourceTypeMaterial) And (r.Name = xlWs.Cells(i, NameCol + MaterialCol).Value) Then '(r.ID = xlWs.Cells(i, 102).Value) Then
                       fnd = False
                       For Each a In r.Assignments
                           onoma = Mid(a.Project, InStrRev(a.Project, "\") + 1, Len(a.Project) - InStrRev(a.Project, "\"))
                           If a.Project = onoma And a.TaskName = xlWs.Cells(i, TaskCol + MaterialCol) Then 'a.TaskID = xlWs.Cells(i, 101) Then
                            fnd = True
                            For Each TSV In a.TimeScaleData(StartDate:=fromD, EndDate:=fromD, Type:=pjAssignmentTimescaledActualWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                                TSV.Value = Val(TSV.Value) + 1 'xlWs.Cells(i, TimeCol + MaterialCol).Value
                            Next TSV
                           End If
                       Next a
                       If fnd = False Then
                           Set a = tsk.Assignments.Add(ResourceID:=r.ID, Units:=1) 'Units:=(xlWs.Cells(i, TimeCol).Value / 8))
                           For Each TSV In a.TimeScaleData(StartDate:=fromD, EndDate:=fromD, Type:=pjAssignmentTimescaledActualWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                              TSV.Value = Val(TSV.Value) + 1 'xlWs.Cells(i, TimeCol).Value * 60
                           Next TSV
                       End If
                    End If
                Next r
           End If
        Next i
    xlWs.Range("D4").SpecialCells(xlCellTypeSameValidation).Select
    With xlApl.Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    xlWs.Range("F4").SpecialCells(xlCellTypeSameValidation).Select
    With xlApl.Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    xlWs.Range("N4").SpecialCells(xlCellTypeSameValidation).Select
    With xlApl.Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    xlWs.Range("P4").SpecialCells(xlCellTypeSameValidation).Select
    With xlApl.Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
        xlWs.Rows(frow & ":" & FR).Select
        xlApl.Selection.Copy
        xlLogAll.Cells(xlLogAll.Cells(xlLogAll.Rows.Count, NameCol).End(xlUp).Row + 1, 1).PasteSpecial _
         Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Cont:
        L = L + 1
        fromD = xlLogWs.Cells(L, 2).Value 'DateAdd("d", 1, fromD)
    Loop
    If Not xlWb Is Nothing Then
        xlWb.Save
        For Each nm In xlLogWb.Names
            nm.Delete
        Next
        xlLogWb.Save
    End If
    If Not xlWb Is Nothing Then
            xlWb.Close
            Set xlWb = Nothing
    End If
    If Not xlLogWb Is Nothing Then
        For j = xlApl.Workbooks.Count To 1 Step -1
            Tname = Application.ActiveProject.CustomDocumentProperties("project")
            If Left(xlApl.Workbooks(j).Name, Len(Trim(Tname))) = Tname Then
                xlApl.Workbooks(j).Save
                xlApl.Workbooks(j).Close
            End If
        Next j
        ews = xlLogWs.Cells(xlLogWs.Rows.Count, 5).End(xlUp).Row
        For t = 2 To ews
            If xlLogWs.Cells(t, 5).Value = "Upd" And xlLogWs.Cells(t, 6).Value = UpdDate Then
                If Dir(xlLogWs.Cells(t, 3).Value) <> "" Then
                    'On Error Resume Next
                    SetAttr xlLogWs.Cells(t, 3).Value, vbNormal
                    Kill (xlLogWs.Cells(t, 3).Value)
                    'On Error GoTo 0
                    xlLogWs.Cells(t, 5).Value = "Del"
                    xlLogWs.Cells(t, 7).Value = Now()
                End If
            End If
        Next t
        For Each nm In xlLogWb.Names
            nm.Delete
        Next
        xlLogWb.Save
        xlLogWb.Close
        Set xlLogWb = Nothing
    End If
    If Not RunExcel Then
            xlApl.Quit
            Set xlApl = Nothing
    Else
            xlApl.DisplayAlerts = True
            xlApl.DisplayAlerts = True
            xlApl.ScreenUpdating = True
            xlApl.WindowState = xlMinimized
            xlApl.Visible = True
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    UpdateProgressBar 0
    MsgBox "Finish....."
End Sub
Sub ttt()
    StartUF.Show
End Sub
Sub Hores(ByVal fd As Date)
    Dim fromD, toD As Date
    Dim rs As Resources
    Dim r As Resource
    Dim a As Assignment
    Dim TSV As TimeScaleValue
    Dim TSVS As TimeScaleValues
    'MsgBox ActiveProject.Name
    Dim fname As String
    Dim xlApl As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlWs As Excel.Worksheet
    Set xlApl = New Excel.Application
    Dim RunExcel As Boolean
    RunExcel = False
    If Not IsAppRunning("Excel.Application") Then
        cPath = Application.ActiveProject.Path
        RunExcel = True
    End If
    'xlApl.Visible = False
    'Application.DisplayAlerts = False
    'Application.ScreenUpdating = False
    Set xlWb = xlApl.Workbooks.Open(Application.ActiveProject.Path & "\month.xls", True, False)
    Set xlWs = xlWb.Worksheets(1)
    FinalRow = xlWs.Cells(xlWs.Rows.Count, 1).End(xlUp).Row
    'fromD = Format("1/" & Format(Now(), "mm/yyyy"), "dd/mm/yyyy")
    toD = Format(fd, "dd/mm/yyyy")
    
    fromD = fd
    toD = DateAdd("m", 1, fromD)
    
    sname = Application.ActiveProject.CustomDocumentProperties("project") & Format(fromD, "_mm_yyyy") & ".xls"
    fname = Application.ActiveProject.Path & "\" & sname
    If Dir(fname) <> "" Then
        MsgBox "The file " & fname & " allready exists"
        GoTo nextD
    End If
    xlWb.SaveAs (fname)
    'xlWb.Close
    
    Set xlWb = xlApl.Workbooks.Open(fname, True, False)
    Set xlWs = xlWb.Worksheets(1)
    
    Call upd_status(sname)
    
    mera = fromD
    ews = Day(DateAdd("d", -1, toD))
    For k = 1 To ews
        stili = k * 2 + 1
        xlWs.Cells(1, stili).Value = Format(CStr(k) & "/" & Format(fromD, "mm/yyyy"), "dd ddd") 'mera, "dd ddd")
        'mera = DateAdd("d", 1, fromD)
    Next k
    Dim counter, ola, pctDone As Single
    counter = 0
    ola = IIf(toD > fromD, toD - fromD, 1)
    Do While fromD <= toD
                counter = counter + 1
                pctDone = counter / ola
                UpdateProgressBar pctDone
        'xlApl.Visible = True
        'xlApl.DisplayAlerts = True
        'xlApl.ScreenUpdating = True
         'Header
        Set rs = ActiveProject.Resources
        'Headers 1st line
        L = 3
        For Each r In rs
            If r.Type = pjResourceTypeWork Then
                xlWs.Cells(L, 1).Value = r.Code
                xlWs.Cells(L, 2).Value = r.Name
                wres = 0
                For Each a In r.Assignments
                    onoma = Mid(a.Project, InStrRev(a.Project, "\") + 1, Len(a.Project) - InStrRev(a.Project, "\"))
                    If a.Project = onoma Then
                        For Each TSV In a.TimeScaleData(fromD, fromD, Type:=pjAssignmentTimescaledActualWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                            If TSV.Value <> "" And TSV.Value > "0" Then
                                    wres = wres + TSV.Value / 60
                            End If
                        Next TSV
                    End If
                Next
                stili = Day(fromD) * 2 + 1
                'stili = Day(fromD) + 2
                If wres > 0 Then
                    xlWs.Cells(L, stili).Value = Format(ActiveProject.DefaultStartTime, "hh:mm")
                    xlWs.Cells(L, stili + 1).Value = Format(DateAdd("h", wres, ActiveProject.DefaultStartTime), "hh:mm")
                    xlWs.Cells(L, stili + 100).Value = wres
                End If
            End If
            L = L + 1
        Next r
        fromD = DateAdd("d", 1, fromD)
        L = L + 1
    Loop
nextD:
    L = L - 4
    counter = 0
    ola = L
    For i = L To 3 Step -1
                counter = counter + 1
                pctDone = counter / ola
                UpdateProgressBar pctDone
        vrika = False
        For j = 3 To xlWs.Cells(i, xlWs.Columns.Count).End(xlToLeft).Column + 1
            If Val(xlWs.Cells(i, j).Value) <> 0 Then
                vrika = True
            End If
        Next j
        If vrika = False Then
            xlWs.Cells(i, 1).EntireRow.Delete
        End If
    Next
    xlWb.Save
    If Not xlWb Is Nothing Then
            xlWb.Close
            Set xlWb = Nothing
    End If
    If Not RunExcel Then
            xlApl.Quit
            Set xlApl = Nothing
    Else
            xlApl.DisplayAlerts = True
            xlApl.DisplayAlerts = True
            xlApl.ScreenUpdating = True
            xlApl.Visible = True
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    UpdateProgressBar 0
    MsgBox "Finish/Updated file:" & crLF & fname
End Sub
Sub tt()
PfromD = Format("1/" & CStr(Mines.Index) & "/" & Year(Now()), "dd/mm/yyyy")
Call Hores(PfromD)
End Sub
Public Sub ShowHideFR()
    If ActiveProject.ReadOnly Then
        ans = MsgBox("File was opened as read-only" & vbLf & "Try to Open Resources as Read Write", vbCritical, "Error")
    Else
        If StartUI.Visible = False Then
            StartUI.Show
        Else
            StartUI.Hide
        End If
    End If
End Sub
Private Sub RunApl()
    Dim x As Variant
    Dim cPath As String
    If Not IsAppRunning("Excel.Application") Then
        cPath = Application.ActiveProject.Path
        'x = Shell(cPath, vbNormalFocus)
    End If
End Sub
Function IsAppRunning(ByVal sAppName) As Boolean
    Dim oApp As Object
    On Error Resume Next
    Set oApp = GetObject(, sAppName)
    If Err Then
        Set oApp = CreateObject(sAppName)
    End If
    IsAppRunning = True
    'If Not oApp Is Nothing Then
    '    Set oApp = Nothing
    '    IsAppRunning = True
    'End If
End Function
Sub CreateXL(Project As String)
    Dim xlApl As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlWs As Excel.Worksheet
    Set xlApl = Excel.Application
    
    RunExcel = False
    If Not IsAppRunning("Excel.Application") Then
        cPath = Application.ActiveProject.Path
        RunExcel = True
    End If
    xlApl.Visible = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim lngCount As Long
    ' Open the file dialog
    'With xlApl.FileDialog(msoFileDialogFolderPicker)
    '    .AllowMultiSelect = False
    '    .Show
    '     nPath = .SelectedItems(1)
    'End With
    fname = Application.ActiveProject.Path & "\Log.xlsx"
    If Dir(fname) <> "" Then
        MsgBox "The file " & fname & " allready exists"
        GoTo nextch
    End If
    nPath = Application.ActiveProject.Path
    xlApl.Workbooks.Add
    Set xlWb = xlApl.ActiveWorkbook
    xlWb.SaveAs Application.ActiveProject.Path & "\Log.xlsx"
    'ChDir nPath
    xlWb.Sheets.Add after:=ActiveSheet
    xlWb.Sheets(1).Select
    xlWb.Sheets(1).Name = "Log"
    xlWb.Sheets("Log").Range("A1").Value = "Αρχείο"
    xlWb.Sheets("Log").Range("B1").Value = "Ημερομηνία"
    xlWb.Sheets("Log").Range("C1").Value = "Full path name"
    xlWb.Sheets("Log").Range("D1").Value = "CrTime'"
    xlWb.Sheets("Log").Range("E1").Value = "Status"
    xlWb.Sheets("Log").Range("F1").Value = "UpdTime"
    xlWb.Sheets("Log").Range("G1").Value = "DelTime"
    xlWb.Sheets(2).Select
    xlWb.Sheets(2).Name = "All"
    xlWb.Sheets("All").Range("D1").Value = "Εργο"
    xlWb.Sheets("All").Range("F1").Value = "Κωδ. Έργου:"
    xlWb.Sheets("All").Range("B3").Value = "α/α"
    xlWb.Sheets("All").Range("C3").Value = "Περιοχή εργασιων"
    xlWb.Sheets("All").Range("D3").Value = "Εργαζόμενος"
    xlWb.Sheets("All").Range("E3").Value = "Ωρες"
    xlWb.Sheets("All").Range("F3").Value = "Περιγραφή Εργασιών"
    xlWb.Sheets("All").Range("G3").Value = "Παρατηρήσεις"
    xlWb.Sheets("All").Range("H3").Value = "Ποσότητα"
    xlWb.Sheets("All").Range("I3").Value = "Αρ. Τιμ."
    xlWb.Sheets("All").Range("J3").Value = "Υπεύθυνος Ομάδας"
    xlWb.Sheets("All").Range("K3").Value = "Μηχανικός"
    xlWb.Sheets("All").Range("L3").Value = "Ημερομηνία"
    xlWb.Sheets("All").Range("M3").Value = "α/α"
    xlWb.Sheets("All").Range("N3").Value = "Όχημα - Μηχάνημα"
    xlWb.Sheets("All").Range("O3").Value = "Αρ. Κυκλ. / Αρ. Πλαισίου"
    xlWb.Sheets("All").Range("P3").Value = "Εργασία που Χρησιμοποιήθηκε"
    xlWb.Save
    xlWb.Close
nextch:
    fname = Application.ActiveProject.Path & "\Start_Initial.xlsx"
    If Dir(fname) <> "" Then
        MsgBox "The file " & fname & " allready exists"
        GoTo nextch1
    End If
    nPath = Application.ActiveProject.Path
    xlApl.Workbooks.Add
    Set xlWb = xlApl.ActiveWorkbook
    xlWb.SaveAs FileName:=nPath & "\Start_Initial.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    xlWb.Sheets.Add after:=ActiveSheet
    xlWb.Sheets(1).Select
    With xlWb.Sheets(1)
        .Range("D1").Value = "Εργο"
        .Range("F1").Value = "Κωδ. Έργου:"
        .Range("B3").Value = "α/α"
        .Range("C3").Value = "Περιοχή εργασιων"
        .Range("D3").Value = "Εργαζόμενος"
        .Range("E3").Value = "Ωρες"
        .Range("F3").Value = "Περιγραφή Εργασιών"
        .Range("G3").Value = "Παρατηρήσεις"
        .Range("H3").Value = "Ποσότητα"
        .Range("I3").Value = "Αρ. Τιμ."
        .Range("J3").Value = "Υπεύθυνος Ομάδας"
        .Range("K3").Value = "Μηχανικός"
        .Range("L3").Value = "Ημερομηνία"
        .Range("M3").Value = "α/α"
        .Range("N3").Value = "Όχημα - Μηχάνημα"
        .Range("O3").Value = "Αρ. Κυκλ. / Αρ. Πλαισίου"
        .Range("P3").Value = "Εργασία που Χρησιμοποιήθηκε"
    End With
    xlWb.Sheets.Add after:=ActiveSheet
    xlWb.Sheets(2).Select
    xlWb.Sheets(2).Name = "Params"
    With xlWb.Sheets("Params")
        .Cells(1, 1) = "FirstLine"
        .Cells(1, 2) = 3
        .Cells(2, 1) = "NameCol"
        .Cells(2, 2) = 4
        .Cells(3, 1) = "TaskCol"
        .Cells(3, 2) = 6
        .Cells(4, 1) = "TimeCol"
        .Cells(4, 2) = 5
        .Cells(5, 1) = "DateCol"
        .Cells(5, 2) = 12
        .Cells(6, 1) = "MaterialCol"
        .Cells(6, 2) = 10
        .Cells(7, 1) = "CostCode"
        .Cells(7, 2) = 20
        .Cells(8, 1) = "CodeCol"
        .Cells(8, 2) = 7
        .Cells(9, 1) = "PrjName"
        .Cells(9, 2) = 1
        .Cells(9, 3) = 5
        .Cells(10, 1) = "PrjCode"
        .Cells(10, 2) = 1
        .Cells(10, 3) = 7
        .Cells(11, 1) = "FirstCol"
        .Cells(11, 2) = 4
        .Cells(12, 1) = "RemarksCol"
        .Cells(12, 2) = 7
        .Cells(13, 1) = "DateCel"
        .Cells(13, 2) = 1
        .Cells(13, 3) = 11
    End With
    xlWb.Save
    xlWb.Close
nextch1:
    fname = Application.ActiveProject.Path & "\month.xls"
    If Dir(fname) <> "" Then
        MsgBox "The file " & fname & " allready exists"
        Exit Sub
    End If
    nPath = Application.ActiveProject.Path
    xlApl.Workbooks.Add
    ChDir nPath
    Set xlWb = xlApl.ActiveWorkbook
    xlWb.SaveAs FileName:=nPath & "\month.xls", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    xlWb.Sheets.Add after:=ActiveSheet
    xlWb.Sheets("Sheet1").Select
    With xlWb.Sheets("Sheet1")
        .Range("A1").Value = "Στοιχεία Εργαζόμενου"
        .Range("a2").Value = "Κωδικός"
        .Range("b3").Value = "Εργαζόμενος"
        For i = 3 To 63 Step 2
            .Cells(2, i).Value = "Είσ."
            .Cells(2, i + 1).Value = "Έξ."
        Next
    End With

    xlWb.Save
    If Not xlWb Is Nothing Then
            xlWb.Close
            Set xlWb = Nothing
    End If
    If Not RunExcel Then
            xlApl.Quit
            Set xlApl = Nothing
    Else
            xlApl.DisplayAlerts = True
            xlApl.DisplayAlerts = True
            xlApl.ScreenUpdating = True
            xlApl.Visible = True
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Finish...."
End Sub
Sub OpenFile()
   FileOpenEx Name:="E:\educ\VBA\Strikos\test\Prj1.mpp", ReadOnly:=False, FormatID:="MSProject.MPP"
End Sub
Sub CreateXL1()
    Dim strFileToOpen As String
    strFileToOpen = Application.GetOpenFilename _
    (Title:="Please choose a file to open", _
    FileFilter:="Excel Files *.xls* (*.xls*),")
    If strFileToOpen = False Then
        MsgBox "No file selected.", vbExclamation, "Sorry!"
        Exit Sub
    Else
        Workbooks.Open FileName:=strFileToOpen
    End If
End Sub
Sub UseFileDialogOpen()
    Dim lngCount As Long
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
         ' Display paths of each file selected
        For lngCount = 1 To .SelectedItems.Count
            MsgBox .SelectedItems(lngCount)
        Next lngCount
    End With
End Sub
Sub SortbyTasks(wsTarget As Worksheet)
    Dim rnTarget, rnSort As Range
    Dim fline, cline, sline As Integer
    Set rnSort = wsTarget.Range(wsTarget.Cells(FirstLine, TaskCol), wsTarget.Cells(FirstLine, TaskCol).End(xlDown))
    Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine, FirstCol - 1), wsTarget.Cells(FirstLine + 1, 104).End(xlDown))
    wsTarget.Sort.SortFields.Clear
    wsTarget.Sort.SortFields.Add Key:=rnSort, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wsTarget.Sort
        .SetRange rnTarget
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Valid(wsTarget As Worksheet, extendRng As Integer)
    Dim rnTarget As Range
    Dim fline, cline, sline As Integer
    'marge ta kelia me idia tasks
    fline = wsTarget.Cells(wsTarget.Rows.Count, 101).End(xlUp).Row
    bTask = wsTarget.Cells(FirstLine, TaskCol).Value
    sline = FirstLine
    cline = FirstLine + 1
    While cline <= fline + 1
        If wsTarget.Cells(sline, TaskCol).Value <> wsTarget.Cells(cline, TaskCol).Value Then
           wsTarget.Range(wsTarget.Cells(sline, RemarksCol), wsTarget.Cells(cline - 1, RemarksCol)).Merge
           sline = cline
        End If
        cline = cline + 1
    Wend
    
    Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, TaskCol), wsTarget.Cells(fline + extendRng, TaskCol))
    'StrRng = Left(rnTarget.Address, InStrRev(rnTarget.Address, "$")) _
    '& CStr(Right(rnTarget.Address, Len(rnTarget.Address) - InStrRev(rnTarget.Address, "$")) + extendRng)
    'Set rnTarget = wsTarget.Range(StrRng)
    With rnTarget
        '.ClearContents
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:="=Tasks"
            .ErrorTitle = "Error"
            .ErrorMessage = "You can only choose from the list."
        End With
    End With
    'Set rnTarget = wsTarget.Range(wsTarget.Range("d4"), wsTarget.Range("d4").End(xlDown))
    Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, NameCol), wsTarget.Cells(fline + extendRng, NameCol))
    'Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, NameCol), wsTarget.Cells(wsTarget.Rows.Count, NameCol).End(xlUp))
    'StrRng = Left(rnTarget.Address, InStrRev(rnTarget.Address, "$")) _
    '& CStr(Right(rnTarget.Address, Len(rnTarget.Address) - InStrRev(rnTarget.Address, "$")) + extendRng)
    'Set rnTarget = wsTarget.Range(StrRng)
    'Set rnTarget = wsTarget.Range(Range("b3"), Range("b3").End(xlDown))
    With rnTarget
        '.ClearContents
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:="=Resources"
            .ErrorTitle = "Value Error"
            .ErrorMessage = "You can only choose from the list."
        End With
    End With
    'Cars
    Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, TaskCol + MaterialCol), wsTarget.Cells(fline + extendRng, TaskCol + MaterialCol))
    'Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, TaskCol + MaterialCol), wsTarget.Cells(wsTarget.Rows.Count, TaskCol + MaterialCol).End(xlUp))
    'StrRng = Left(rnTarget.Address, InStrRev(rnTarget.Address, "$")) _
    '& CStr(Right(rnTarget.Address, Len(rnTarget.Address) - InStrRev(rnTarget.Address, "$")) + extendRng)
    'Set rnTarget = wsTarget.Range(StrRng)
    With rnTarget
        '.ClearContents
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:="=Tasks"
            .ErrorTitle = "Error"
            .ErrorMessage = "You can only choose from the list."
        End With
    End With
    Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, NameCol + MaterialCol), wsTarget.Cells(fline + extendRng, NameCol + MaterialCol))
    'Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, NameCol + MaterialCol), wsTarget.Cells(wsTarget.Rows.Count, NameCol + MaterialCol).End(xlUp))
    'StrRng = Left(rnTarget.Address, InStrRev(rnTarget.Address, "$")) _
    '& CStr(Right(rnTarget.Address, Len(rnTarget.Address) - InStrRev(rnTarget.Address, "$")) + extendRng)
    'Set rnTarget = wsTarget.Range(StrRng)
    'Set rnTarget = wsTarget.Range(Range("b3"), Range("b3").End(xlDown))
    'Clear out any artifacts from previous macro runs, then set up the target range with the validation data.
    On Error Resume Next
    With rnTarget
        '.ClearContents
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:="=Material"
            .ErrorTitle = "Value Error"
            .ErrorMessage = "You can only choose from the list."
        End With
    End With
    'Vlookup for TaskID and ResourceID stin faneri kai stin kryfi perioxi
    If wsTarget.Cells(FirstLine + 1, TaskCol) <> "" Then
        fline = wsTarget.Cells(FirstLine, TaskCol).End(xlDown).Row + 1
    Else
        fline = FirstLine + 1
    End If
    For i = fline To (fline + extendRng)
        Set mt = wsTarget.Cells(i, TaskCol)
        wsTarget.Cells(i, 101).FormulaLocal = "=Vlookup(" & mt.Address & ";TasksAll;2;FALSE)"
        Set mt = wsTarget.Cells(i, TaskCol)
        wsTarget.Cells(i, NameCol - 1).FormulaLocal = "=IFERROR(Vlookup(" & mt.Address & ";TasksAll;3;FALSE);" & Chr(34) & Chr(34) & ")"
        Set mr = wsTarget.Cells(i, NameCol)
        wsTarget.Cells(i, 102).FormulaLocal = "=Vlookup(" & mr.Address & ";ResourcesAll;2;FALSE)"
    Next i
    'Vlookup for materialID
    If wsTarget.Cells(FirstLine + 1, NameCol + MaterialCol) <> "" Then
        fline = wsTarget.Cells(FirstLine, NameCol + MaterialCol).End(xlDown).Row + 1
    Else
        fline = FirstLine + 1
    End If
    For i = fline To (fline + extendRng)
        Set mm = wsTarget.Cells(i, NameCol + MaterialCol)
        wsTarget.Cells(i, 111).FormulaLocal = "=Vlookup(" & mm.Address & ";MaterialAll;2;FALSE)"
    Next i
End Sub
Sub Costos()
    Dim xlApl As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim xlWs As Excel.Worksheet
    Set xlApl = Excel.Application
    If Not IsAppRunning("Excel.Application") Then
        cPath = Application.ActiveProject.Path
        'x = Shell(cPath, vbNormalFocus)
        CreateObject ("Excel.Application")
    End If
        'xlWb.Close
    xlApl.Visible = True
    Dim lngCount As Long
    ' Open the file dialog
    Set fd = xlApl.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Select File (Cost)"
        .Filters.Clear
        .Filters.Add "Excel files", "*.xls?"
        If .Show = True Then
            fn = Dir(.SelectedItems(1))
        End If
    End With
    xlApl.ScreenUpdating = False
    xlApl.DisplayAlerts = False
    xlApl.Visible = True
    Set xlWb = xlApl.Workbooks.Open(fn, True, False)
    Set xlWs = xlWb.Worksheets(1)
    'Workbooks.Open (fn)
    'Application.ScreenUpdating = True
    'Application.DisplayAlerts = True
    'Worksheets(1).Select
    FinalRow = xlWs.Cells(xlWs.Rows.Count, 11).End(xlUp).Row
    Kostos = xlWs.Cells(FinalRow, 11).Value
    
    On Error GoTo ErrCost
    ResourceAssignment Resources:="δαπάνες"
    ResourceAssignment Resources:="δαπάνες[0 ;;1.000,00 €]", Operation:=pjChange
    SelectTaskField Row:=1, Column:="Name"
    'electTaskField Row:=0, Column:="Start"
    'SelectTaskField Row:=0, Column:="Name"
    SetResourceField Field:="Cost", Value:=Kostos, AllSelectedResources:=True
    MsgBox "Assign Cost to Project : " & Kostos
    GoTo Cont:
ErrCost:
    xlApl.WindowState = xlMinimized
    MsgBox "Error with Cost UPD (Can't found Resource <δαπάνες>) "
Cont:
    If Not xlWb Is Nothing Then
            xlWb.Close
            Set xlWb = Nothing
    End If
    If Not RunExcel Then
            xlApl.Quit
            Set xlApl = Nothing
    Else
            xlApl.DisplayAlerts = True
            xlApl.ScreenUpdating = True
            xlApl.Visible = True
    End If
End Sub
Sub ValidN(wsTarget As Worksheet, extendRng As Integer)
    Dim rnTarget As Range
    Dim fline, cline, sline As Integer
    fline = FirstLine + 1 'wsTarget.Cells(wsTarget.Rows.Count, TaskCol).End(xlUp).Row
    Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, TaskCol), wsTarget.Cells(FirstLine + 1 + extendRng, TaskCol))
    With rnTarget
        '.ClearContents
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:="=Tasks"
            .ErrorTitle = "Error"
            .ErrorMessage = "You can only choose from the list."
        End With
    End With
    'Set rnTarget = wsTarget.Range(wsTarget.Range("d4"), wsTarget.Range("d4").End(xlDown))
    Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, NameCol), wsTarget.Cells(FirstLine + 1 + extendRng, NameCol))
    'Set rnTarget = wsTarget.Range(Range("b3"), Range("b3").End(xlDown))
    With rnTarget
        '.ClearContents
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:="=Resources"
            .ErrorTitle = "Value Error"
            .ErrorMessage = "You can only choose from the list."
        End With
    End With
    'Cars
    Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, TaskCol + MaterialCol), wsTarget.Cells(FirstLine + 1 + extendRng, TaskCol + MaterialCol))
    With rnTarget
        '.ClearContents
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:="=Tasks"
            .ErrorTitle = "Error"
            .ErrorMessage = "You can only choose from the list."
        End With
    End With
    
    Set rnTarget = wsTarget.Range(wsTarget.Cells(FirstLine + 1, NameCol + MaterialCol), wsTarget.Cells(FirstLine + 1 + extendRng, NameCol + MaterialCol))
    'Set rnTarget = wsTarget.Range(Range("b3"), Range("b3").End(xlDown))
    With rnTarget
        '.ClearContents
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Formula1:="=Material"
            .ErrorTitle = "Value Error"
            .ErrorMessage = "You can only choose from the list."
        End With
    End With
    'Vlookup for TaskID and ResourceID stin kryfi perioxi
    If wsTarget.Cells(FirstLine + 1, TaskCol) <> "" Then
        fline = wsTarget.Cells(FirstLine + 1, TaskCol).End(xlDown).Row + 1
    Else
        fline = FirstLine + 1
    End If
    For i = fline To (fline + extendRng)
        Set mt = wsTarget.Cells(i, TaskCol)
        wsTarget.Cells(i, 101).FormulaLocal = "=Vlookup(" & mt.Address & ";TasksAll;2;FALSE)"
        Set mt = wsTarget.Cells(i, TaskCol)
        wsTarget.Cells(i, NameCol - 1).FormulaLocal = "=IFERROR(Vlookup(" & mt.Address & ";TasksAll;3;FALSE);" & Chr(34) & Chr(34) & ")"
        Set mr = wsTarget.Cells(i, NameCol)
        wsTarget.Cells(i, 102).FormulaLocal = "=Vlookup(" & mr.Address & ";ResourcesAll;2;FALSE)"
    Next i
    'Vlookup for materialID
    If wsTarget.Cells(FirstLine + 1, TaskCol + MaterialCol) <> "" Then
        fline = wsTarget.Cells(FirstLine + 1, TaskCol + MaterialCol).End(xlDown).Row + 1
    Else
        fline = FirstLine + 1
    End If
    For i = fline To (fline + extendRng)
        Set mm = wsTarget.Cells(i, NameCol + MaterialCol)
        wsTarget.Cells(i, 111).FormulaLocal = "=Vlookup(" & mm.Address & ";MaterialAll;2;FALSE)"
    Next i
End Sub
Sub upd_status(ByVal xlsName As String)
    StartUI.Status.Caption = ActiveProject.Path & " : " & ActiveProject.Name & " : " & xlsName
End Sub
Sub UpdateProgressBar(pctDone As Single)
    pctDone = IIf(pctDone > 1, 1, pctDone)
    With StartUI
        .FrameProgress.Caption = Format(pctDone, "0%")
        .LabelProgress.Width = pctDone * (.FrameProgress.Width - 10)
    End With
    'DoEvents
End Sub

