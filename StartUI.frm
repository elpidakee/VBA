VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartUI 
   Caption         =   "Schedule - Update Project"
   ClientHeight    =   4800
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6825
   OleObjectBlob   =   "StartUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StartUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AssignName_Click()
    Dim fnd As Boolean
    'On Error Resume Next
    'Application.ActiveProject.CustomDocumentProperties.Add ("project")
    fnd = False
    For i = 1 To ActiveProject.CustomDocumentProperties.Count
        If ActiveProject.CustomDocumentProperties(i).Name = "project" Then
           fnd = True
        End If
    Next
    If Not fnd Then
        With ActiveProject.CustomDocumentProperties
            .Add Name:="project", _
                LinkToContent:=False, _
                Type:=msoPropertyTypeString, _
                Value:=Project.Text
        End With
    End If
    Application.ActiveProject.CustomDocumentProperties("project") = Project.Text
End Sub

Private Sub CbEtos_Change()
    For i = 1 To Mines.ListCount
        Mines.RemoveItem 0
    Next
    For i = 1 To 12
        Mines.AddItem (Format("1/" & CStr(i) & "/" & CbEtos.Text, "mmm yyyy"))
    Next i
End Sub

Private Sub Cost_Click()
    Call Costos
End Sub

Private Sub CrProject_Click()
    If IsEmpty(Project.Text) Then
        a = MsgBox("Empty Project Name", vbCritical, "Error")
    Else
        Call CreateXL(Project.Text)
    End If
End Sub
Private Sub DTPSfromD_Change()
    DTPStoD.Value = DTPSfromD.Value
End Sub

Private Sub DTPUfromD_Change()
    DTPUtoD.Value = DTPUfromD.Value
End Sub

Private Sub Payroll_Click()
    PfromD = Format("1/" & CStr(Mines.ListIndex + 1) & "/" & Year(Now()), "dd/mm/yyyy")
    Call Hores(PfromD)
End Sub
Private Sub Schedule_Click()
    'Call SubSD(CDate(SfromD.Text), CDate(StoD.Text))
    Call SubSD(Format(CDate(DTPSfromD.Value), "dd/mm/yyyy"), Format(CDate(DTPStoD.Value), "dd/mm/yyyy"))
End Sub

Private Sub Update_Click()
    'Call Upd(CDate(SfromD.Text), CDate(StoD.Text))
    Call Upd(Format(CDate(DTPUfromD.Value), "dd/mm/yyyy"), Format(CDate(DTPUtoD.Value), "dd/mm/yyyy"))
End Sub
Private Sub UserForm_Initialize()
        LabelProgress.Width = 0
        DTPSfromD.MinDate = Application.ActiveProject.ProjectStart
        DTPSfromD.MaxDate = Application.ActiveProject.ProjectFinish
        DTPSfromD.Value = IIf(Now() > DTPSfromD.MaxDate, DTPSfromD.MaxDate, Now())
        DTPStoD.MinDate = Application.ActiveProject.ProjectStart
        DTPStoD.MaxDate = Application.ActiveProject.ProjectFinish
        DTPStoD.Value = IIf(Now() > DTPSfromD.MaxDate, DTPSfromD.MaxDate, Now())
        DTPUfromD.MinDate = Application.ActiveProject.ProjectStart
        DTPUfromD.MaxDate = Application.ActiveProject.ProjectFinish
        DTPUfromD.Value = IIf(Now() > DTPSfromD.MaxDate, DTPSfromD.MaxDate, Now())
        DTPUtoD.MinDate = Application.ActiveProject.ProjectStart
        DTPUtoD.MaxDate = Application.ActiveProject.ProjectFinish
        DTPUtoD.Value = IIf(Now() > DTPSfromD.MaxDate, DTPSfromD.MaxDate, Now())

        'SfromD.Text = Format(Now(), "dd/mm/yyyy")
        'StoD.Text = Format(Now(), "dd/mm/yyyy")
        'UfromD.Text = Format(Now(), "dd/mm/yyyy")
        'UtoD.Text = Format(Now(), "dd/mm/yyyy")
        For i = Year(Application.ActiveProject.ProjectStart) To Year(Application.ActiveProject.ProjectFinish)
            CbEtos.AddItem i
        Next i
        For i = 1 To 12
            Mines.AddItem (Format("1/" & CStr(i) & "/" & Year(Now()), "mmm yyyy"))
        Next i
        On Error Resume Next
        Project.Text = Application.ActiveProject.CustomDocumentProperties("project")
        FirstLine = 7
        NameCol = 4
        TaskCol = 9
        TimeCol = 8
        DateCol = 7
        MaterialCol = 10
        CostCode = 20
        CodeCol = 21
        FirstCol = 4
        Call upd_status("initialize")
End Sub

