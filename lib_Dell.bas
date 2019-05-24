Attribute VB_Name = "lib"
Public NameCol, TaskCol, TimeCol, DateCol, MaterialCol, FirstLine, CodeCol, CostCol As Integer
Public Sub OpenFR()
    If StartUI.Visible = False Then
        'Load StartUI
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
        'x = Shell(cPath, vbNormalFocus)
        'CreateObject ("Excel.Application")
        RunExcel = True
    End If
    xlApl.Visible = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim xlLogWb As Excel.Workbook
    Dim xlLogWs As Excel.Worksheet
    Set xlLogWb = xlApl.Workbooks.Open(Application.ActiveProject.Path & "\Log.xlsx", True, False)
    Set xlLogWs = xlLogWb.Worksheets("Log")
    
    finalrow = xlLogWs.Cells(xlLogWs.Rows.Count, 1).End(xlUp).Row
    
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
    Set xlWs = xlWb.Worksheets(1)
    
    fromD = Format(Now(), "dd/mm/yyyy")
    toD = Format(Now(), "dd/mm/yyyy")
    
    fromD = fd
    toD = td
    cPrj = Application.ActiveProject.CustomDocumentProperties("project")
    'fr = xlWs.Cells(xl.Rows.Count, 1).End(xlUp).Row
    Do While fromD <= toD
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
        'xlWb.Close
        'xlApl.Visible = True
        'xlApl.DisplayAlerts = True
        'xlApl.ScreenUpdating = True
         'Header
        Set rs = ActiveProject.Resources
        'Headers 1st line
        'xlWs.Cells(1, 1).Value = "Daily schedule"
        'xlWs.Cells(1, 3).Value = Format(fromD, "dddd")
        'xlWs.Cells(1, 4).Value = fromD
        'xlWs.Cells(1, 5).Value = "Project:"
        'xlWs.Cells(1, 6).Value = Application.ActiveProject.CustomDocumentProperties("project")
        'Headers 2st line
        'xlWs.Cells(2, 1).Value = "SN"
        'xlWs.Cells(2, 2).Value = "Name"
        'xlWs.Cells(2, 3).Value = "Task"
        'xlWs.Cells(2, 4).Value = "Start Date"
        'xlWs.Cells(2, 5).Value = "Work/h"
        'xlWs.Cells(2, 6).Value = "Actual Work/h"
        
        'L = 3
        'xlWs.Cells(L, 12).Value = Format(fromD, "dddd")
        'xlWs.Cells(1, 16).Value = cPrj
        'L = 4
        'xlWs.Cells(L, 12).Value = fromD
        L = FirstLine
        M = FirstLine
        C = FirstLine
        For Each r In rs
                 For Each a In r.Assignments
                    Onoma = Mid(a.Project, InStrRev(a.Project, "\") + 1, Len(a.Project) - InStrRev(a.Project, "\"))
                    If (a.Project = Onoma) Then
                        If r.Type = pjResourceTypeWork Then
                            'Set TSVS = a.TimeScaleData(fromD, fromD, Type:=pjAssignmentTimescaledWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                            For Each TSV In a.TimeScaleData(fromD, fromD, Type:=pjAssignmentTimescaledWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                             If TSV.Value <> "" And TSV.Value > "0" Then
                                L = L + 1
                                'xlWs.Cells(l, 1).Value = l - 7
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
                                M = M + 1
                                xlWs.Cells(M, NameCol + MaterialCol).Value = r.Name
                                xlWs.Cells(M, DateCol + MaterialCol).Value = fromD 'TSV.StartDate
                                xlWs.Cells(M, TimeCol + MaterialCol).Value = TSV.Value / 60
                                xlWs.Cells(M, TaskCol + MaterialCol).Value = a.TaskName
                                
                                xlWs.Cells(M, 111).Value = a.TaskID
                                xlWs.Cells(M, 112).Value = r.ID
                                xlWs.Cells(M, 113).Value = a.UniqueID
                                xlWs.Cells(M, 114).Value = TSV.Value / 60
                             End If
                           Next TSV
                        Else
                           For Each TSV In a.TimeScaleData(fromD, fromD, Type:=pjAssignmentTimescaledWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                            If TSV.Value <> "" And TSV.Value > "0" Then
                                C = C + 1
                                xlWs.Cells(C, NameCol + CostCol).Value = r.Name
                                xlWs.Cells(C, DateCol + CostCol).Value = fromD 'TSV.StartDate
                                xlWs.Cells(C, TimeCol + CostCol).Value = TSV.Value / 60
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
        xlWb.Save
        finalrow = finalrow + 1
        xlLogWs.Cells(finalrow, 1).Value = sname
        xlLogWs.Cells(finalrow, 2).Value = fromD
        xlLogWs.Cells(finalrow, 3).Value = fname
        xlLogWs.Cells(finalrow, 4).Value = Now()
        
        'xlWb.Close
        'xlApl.Visible = True
nextD:
        fromD = fromD + 1
    Loop
        xlWb.Save
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
            xlApl.Visible = True
        End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Finish....."
End Sub
Sub Upd(ByVal fd As Date, ByVal td As Date)
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
    xlApl.Visible = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim xlLogWb As Excel.Workbook
    Dim xlLogWs As Excel.Worksheet
    Set xlLogWb = xlApl.Workbooks.Open(Application.ActiveProject.Path & "\Log.xlsx", True, False)
    Set xlLogWs = xlLogWb.Worksheets("Log")
    finalrow = xlLogWs.Cells(xlLogWs.Rows.Count, 1).End(xlUp).Row
    fromD = Format(Now(), "dd/mm/yyyy")
    toD = Format(Now(), "dd/mm/yyyy")
    fromD = fd
    toD = td
    L = 2
    Do While L <= finalrow
        'If xlLogWs.Cells(L, 2).Value < fromD And xlLogWs.Cells(L, 2).Value > toD Then
        '    GoTo cont
        'End If
        'If Not (xlLogWs.Cells(L, 5).Value = 1 And xlLogWs.Cells(L, 6).Value <> "") Then
        '    GoTo cont
        'End If
        If Not (Dir(xlLogWs.Cells(L, 3).Value) <> "") Then
            MsgBox "The file " & xlLogWs.Cells(L, 3).Value & " was not found"
            GoTo cont
        End If
        Set xlWb = xlApl.Workbooks.Open(xlLogWs.Cells(L, 3).Value, True, False)
        Set xlWs = xlWb.Worksheets(1)
        frow = xlWs.Cells(1, 101).End(xlDown).Row
        fr = xlWs.Cells(xlWs.Rows.Count, 101).End(xlUp).Row
        Set rs = ActiveProject.Resources
        
        For i = frow To fr
            'ActiveProject.Tasks.UniqueID(xlWs.Cells(i, 101)).Assignments.UniqueID(xlWs.Cells(i, 103)).TimeScaleData(StartDate:=xlWs.Cells(i, 7), EndDate:=xlWs.Cells(i, 7), Type:=pjAssignmentTimescaledActualWork, TimeScaleUnit:=pjTimescaleDays, Count:=1).Item(1).Value = xlWs.Cells(i, 14).Value
             fromD = xlWs.Cells(i, 7) ' stili G?
             For Each r In rs
               If (r.Type = pjResourceTypeWork) And (r.ID = xlWs.Cells(i, 102).Value) Then
                 For Each a In r.Assignments
                   Onoma = Mid(a.Project, InStrRev(a.Project, "\") + 1, Len(a.Project) - InStrRev(a.Project, "\"))
                   If a.Project = Onoma And a.TaskID = xlWs.Cells(i, 101) Then
                     For Each TSV In a.TimeScaleData(StartDate:=fromD, EndDate:=fromD, Type:=pjAssignmentTimescaledActualWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                        TSV.Value = xlWs.Cells(i, 8).Value * 60
                     Next TSV
                   End If
                 Next a
              End If
        Next r
        Next i
cont:
        L = L + 1
    Loop
    xlWb.Save
    xlLogWb.Save
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
            xlApl.Visible = True
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
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
    xlApl.Visible = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Set xlWb = xlApl.Workbooks.Open(Application.ActiveProject.Path & "\month.xls", True, False)
    Set xlWs = xlWb.Worksheets(1)
    finalrow = xlWs.Cells(xlWs.Rows.Count, 1).End(xlUp).Row
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
    
    mera = fromD
    ews = Day(DateAdd("d", -1, toD))
    For k = 1 To ews
        stili = k * 2 + 1
        xlWs.Cells(1, stili).Value = Format(CStr(k) & "/" & Format(fromD, "mm/yyyy"), "dd ddd") 'mera, "dd ddd")
        'mera = DateAdd("d", 1, fromD)
    Next k
    Do While fromD <= toD
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
                    Onoma = Mid(a.Project, InStrRev(a.Project, "\") + 1, Len(a.Project) - InStrRev(a.Project, "\"))
                    If a.Project = Onoma Then
                        For Each TSV In a.TimeScaleData(fromD, fromD, Type:=pjAssignmentTimescaledWork, TimeScaleUnit:=pjTimescaleDays, Count:=1)
                            If TSV.Value <> "" And TSV.Value > "0" Then
                                    wres = ores + TSV.Value / 60
                            End If
                        Next TSV
                    End If
                Next
                stili = Day(fromD) * 2 + 1
                'stili = Day(fromD) + 2
                If wres > 0 Then
                    xlWs.Cells(L, stili).Value = wres
                End If
            End If
            L = L + 1
        Next r
        fromD = DateAdd("d", 1, fromD)
        L = L + 1
    Loop
nextD:
    L = L - 4
    For i = L To 3 Step -1
        vrika = False
        For j = 3 To xlWs.Cells(i, Columns.Count).End(xlToLeft).Column + 1
            If Val(Cells(i, j).Value) <> 0 Then
                vrika = True
            End If
        Next
        If vrika = False Then
            Cells(i, 1).EntireRow.Delete
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
    If Not oApp Is Nothing Then
        Set oApp = Nothing
        IsAppRunning = True
    End If
End Function
Sub tttt()
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
End Sub
