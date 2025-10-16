VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_00(1)eSysCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright: Allen Browne's Database And Training, 2007.
'Author:    Allen Browne. allen@allenbrowne.com
Option Compare Database
Option Explicit

Private Const conMod = "frmCalendar"

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
    'Purpose:   Close without transferring date back to calling text box.
    
    DoCmd.Close acForm, Me.Name, acSaveNo

Exit_Handler:
    Exit Sub

Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".cmdCancel_Click")
    Resume Exit_Handler
End Sub

Private Sub cmdMonthDown_Click()
    Call SetDate("M", -1)
End Sub
Private Sub cmdMonthUp_Click()
    Call SetDate("M", 1)
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
    'Purpose:   Transfer the result back to the calling text box (if there is one), and close.
    
    If Me.cmdOk.Enabled Then
        If gBsCalTarget = Me.txtDate Then
            'do nothing
        Else
            gBsCalTarget = Me.txtDate
        End If
    End If
    gBsCalTarget.SetFocus
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdToday_Click()
'On Error GoTo Err_Handler
    'Purpose:   Set today.
    
    Me.txtDate = Date
    Call ShowCal
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".cmdToday_Click")
    Resume Exit_Handler
End Sub

Private Sub cmdYearDown_Click()
    Call SetDate("YYYY", -1)
End Sub
Private Sub cmdYearUp_Click()
    Call SetDate("YYYY", 1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo Err_Handler
    'Purpose:
    
    With Me.txtDate
        Select Case KeyCode
        Case vbKeyLeft              '1 day left or right.
            .Value = .Value - 1
            KeyCode = 0
            Call ShowCal
        Case vbKeyRight
            .Value = .Value + 1
            KeyCode = 0
            Call ShowCal
        Case vbKeyUp                '1 week up or down.
            .Value = .Value - 7
            KeyCode = 0
            Call ShowCal
        Case vbKeyDown
            .Value = .Value + 7
            KeyCode = 0
            Call ShowCal
        Case vbKeyHome              'Home/End = first/last of this month.
            .Value = .Value - Day(.Value) + 1
            KeyCode = 0
            Call ShowCal
        Case vbKeyEnd
            .Value = DateSerial(Year(.Value), Month(.Value) + 1, 0)
            KeyCode = 0
            Call ShowCal
        Case vbKeyPageUp            'PgUp/PgDn = previous/next month.
            .Value = DateAdd("m", -1, .Value)
            KeyCode = 0
            Call ShowCal
        Case vbKeyPageDown
            .Value = DateAdd("m", 1, .Value)
            KeyCode = 0
            Call ShowCal
        Case vbKeyT, vbKeyT + 32    'T or t = today
            .Value = Date
            KeyCode = 0
            Call ShowCal
        End Select
    End With
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".Form_KeyDown")
    Resume Exit_Handler
End Sub

Private Sub Form_Load()
    '------------------------------------------------------------
    'Define posição do formulário centralizado sobre o formulário que acionou o [ SysDialogBox ]
    '------------------------------------------------------------
    clObjFormOpenParams.bSetPosition = True
    clObjFormOpenParams.sTrgtForm = Me.Name
    Set clObjFormOpenParams.fTrgtForm = Me
    
    Call FormLoad08_SetFormPosition(clObjFormOpenParams)
End Sub

Private Sub Form_Open(Cancel As Integer)
'On Error GoTo Form_Open_Err
    Dim bEnabled As Boolean
    
    'Initialize to the existing date, or today if null.
    If IsDate(gBsCalTarget) Then
        Me.txtDate = gBsCalTarget.Value
    Else
        Me.txtDate = Date
    End If
    
    'Lock the Ok button if the text box is locked or disabled.
    bEnabled = (gBsCalTarget.Enabled) And (Not gBsCalTarget.Locked)
    With Me.cmdOk
        If .Enabled <> bEnabled Then
            .Enabled = bEnabled
        End If
    End With
   
    'Set the title
    If Len(Me.OpenArgs) > 0& Then
        Me.Caption = Me.OpenArgs
    End If
    
    'Set up the calendar for this month.
    Call ShowCal

Form_Open_Exit:
    Exit Sub

Form_Open_Err:
    MsgBox Err.Description, vbCritical, "frmCalendar.Form_Open"
    Resume Form_Open_Exit
End Sub

Private Function SetSelected(ctlName As String)
On Error GoTo Err_Handler

    Me.txtDate = DateSerial(Year(txtDate), Month(txtDate), CLng(Me(ctlName).Caption))
    Call ShowHighligher(ctlName)

Exit_Handler:
    Exit Function

Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".SetSelected")
    Resume Exit_Handler
End Function

Private Function SelectDate(ctlName As String)
    Call SetSelected(ctlName)
    Call cmdOk_Click
End Function

Private Function SetDate(Subject As String, Optional intStep As Integer = 1)
On Error GoTo Err_Handler
    
    Me.txtDate = DateAdd(Subject, intStep, Me.txtDate)
    Call ShowCal

Exit_Handler:
    Exit Function

Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".SetDate")
    Resume Exit_Handler
End Function

Private Function ShowCal() As Boolean
On Error GoTo Err_Handler
    'Purpose:
    Dim dtStartDate As Date     'First of month
    Dim iDays As Integer        'Days in month
    Dim iOffset As Integer      'Offset to first label for month.
    Dim i As Integer            'Loop controller.
    Dim iDay As Integer         'Day under consideration.
    Dim bShow As Boolean        'Flag: show label
    
    dtStartDate = Me.txtDate - Day(Me.txtDate) + 1  'First of month
    iDays = Day(DateAdd("m", 1, dtStartDate) - 1)   'Days in month.
    iOffset = Weekday(dtStartDate, vbSunday) - 2    'Offset to first label for month.
    
    For i = 0 To 41
        With Me("lblDay" & Format(i, "00"))
            iDay = i - iOffset
            bShow = ((iDay > 0) And (iDay <= iDays))
            If .Visible <> bShow Then
                .Visible = bShow
            End If
            If (bShow) And (.Caption <> iDay) Then
                .Caption = iDay
            End If
        End With
    Next
    
    Call ShowHighligher("lblDay" & Format(Day(Me.txtDate) + iOffset, "00"))
    
Exit_Handler:
    Exit Function

Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".ShowCal")
    Resume Exit_Handler
End Function

Private Function ShowHighligher(ctlName As String)
On Error GoTo Err_Handler
    Const lngcVOffset As Long = -83

    With Me(ctlName)
        Me.lblHighlight.Left = .Left
        Me.lblHighlight.Top = .Top + lngcVOffset
    End With

Exit_Handler:
    Exit Function

Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".ShowHighligher")
    Resume Exit_Handler
End Function

Private Sub lblHelp_Click()
    MsgBox Me.lblHelp.ControlTipText, vbInformation, "Calendar help"
End Sub
