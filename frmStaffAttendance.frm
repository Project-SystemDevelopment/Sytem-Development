VERSION 5.00
Begin VB.Form frmStaffAttendance 
   Caption         =   "Staff Attendance"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   6855
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   495
         Left            =   4560
         TabIndex        =   6
         Top             =   155
         Width           =   1095
      End
      Begin VB.CommandButton cmdClockIn 
         Caption         =   "Clock In"
         Default         =   -1  'True
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   155
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   5640
         TabIndex        =   1
         Top             =   155
         Width           =   1095
      End
   End
   Begin VB.TextBox txtReason 
      Height          =   1095
      Left            =   120
      MaxLength       =   500
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1920
      Width           =   6855
   End
   Begin VB.Label lblLate 
      Caption         =   "You are late! Please be in office before "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label lblMorning 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Staff Attendance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7125
   End
End
Attribute VB_Name = "frmStaffAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpTime As String
Dim tmpDateTime As String
Dim tmpTimeformat As String
Dim tmpTimeonly As String
Dim tmpDateonly As String
Dim tmpWeekDay As String

Private Sub ServerTime()
    Dim lpbuffer As Long
    Dim t_struct As TIME_OF_DAY_INFO
    Dim ret As Long
    Dim bServer() As Byte
    
    If Trim(SvrDBPath) = "" Then
        'Local machine
        ret = NetRemoteTOD(vbNullString, lpbuffer)
    Else
        'Check the syntax of the ServerName string
        If InStr(SvrDBPath, "\\") = 1 Then
            bServer = SvrDBPath & vbNullChar
        Else
            bServer = SvrDBPath & vbNullChar
        End If
        ret = NetRemoteTOD(bServer(0), lpbuffer)
    End If
    CopyMem t_struct, ByVal lpbuffer, Len(t_struct)
    If lpbuffer Then
        Call NetApiBufferFree(lpbuffer)
    End If
    
    tmpDateTime = DateSerial(t_struct.tod_year, t_struct.tod_month, t_struct.tod_day) + TimeSerial(t_struct.tod_hours, t_struct.tod_mins - t_struct.tod_timezone, t_struct.tod_secs)
    tmpTime = Format(tmpDateTime, "hh:mm")
    tmpTimeonly = Format(tmpTime, "hh:mm")
    tmpTimeformat = Format(tmpTime, "AM/PM")
    tmpDateonly = Format(tmpDateTime, "dd/mm/yyyy")
    tmpWeekDay = WeekdayName(Weekday(tmpDateonly))
End Sub

Public Sub cmdClockIn_Click()
Dim strsql As String
Dim rst As New ADODB.Recordset
Dim LateStatus As Boolean
Dim UsrCLKTimeAMPM As String
Dim tmplblLate As String
    lblLate = ""
    tmplblLate = ""
    UsrCLKTimeAMPM = IIf(USRCLKTime < "12:00", " AM", " PM")
    LateStatus = False
    tmplblLate = "You are late! Please be at office before "
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

        Call ServerTime
        strsql = ""
        strsql = "Select StaffLogon,StaffName,ClockInTime,ClockInDate, TimeFormat,WeekDay from StaffAttendance where StaffLogon = '" & Logon & "' and ClockInDate = '" & Format(tmpDateonly, "mm/dd/yyyy") & "'"
        rst.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            If rst.recordCount Then
                lblMorning = "Hi, " & rst!StaffName & ". " & " You had clocked in at " & "" & Format(rst!ClockInTime, "hh:mm") & rst!TimeFormat & " on " & rst!Weekday & ", " & rst!ClockInDate
                lblLate.Visible = False
            Else
                rst.AddNew
                With rst
                    !StaffLogon = Trim(Logon)
                    !StaffName = LogonName
                    !ClockInTime = tmpTimeonly
                    !ClockInDate = tmpDateonly
                    !TimeFormat = tmpTimeformat
                    !Weekday = tmpWeekDay
                End With
                
                rst.Update
            'End If
                rst.Close
                Set rst = Nothing
                
                If Format(tmpTimeonly, "hh:mm") > Format(USRCLKTime, "hh:mm") Then
                    lblLate = tmplblLate & USRCLKTime & UsrCLKTimeAMPM & " and enter reason for being late."
                    LateStatus = True
                    frmStaffAttendance.Height = 4395
                    Frame1.Top = 3120
                    txtReason.Visible = False
                    cmdExit.Enabled = True
                    cmdUpdate.Visible = False
                Else
                    LateStatus = False
                    frmStaffAttendance.Height = 4395
                    Frame1.Top = 3120
                    txtReason.Visible = True
                    cmdExit.Enabled = False
                    cmdUpdate.Visible = True
                End If
                
                If tmpTimeformat = "AM" Then
                    If LateStatus = False Then
                        lblMorning = "Good Morning, " & "" & LogonName & "." & " Now is " & "" & tmpWeekDay & ", " & tmpDateTime
                        lblLate.Visible = False
                        frmStaffAttendance.Height = 3210
                        Frame1.Top = 1920
                        txtReason.Visible = False
                        cmdExit.Enabled = True
                        cmdUpdate.Visible = False
                    Else
                        lblMorning = "Good Morning, " & "" & LogonName & "." & " Now is " & "" & tmpWeekDay & ", " & tmpDateTime
                        lblLate = tmplblLate & USRCLKTime & UsrCLKTimeAMPM & " and enter reason for being late."
                        lblLate.Visible = True
                        frmStaffAttendance.Height = 4395
                        Frame1.Top = 3120
                        txtReason.Visible = True
                        cmdExit.Enabled = False
                        cmdUpdate.Visible = True
                    End If
                Else
                    If tmpTimeonly >= "19:00" And tmpTimeonly <= "23:59" Then
                    
                        lblMorning = "Good Evening, " & "" & LogonName & "." & " Now is " & "" & tmpWeekDay & ", " & tmpDateTime
                    
                    Else
                    
                        lblMorning = "Good Afternoon, " & "" & LogonName & "." & " Now is " & "" & tmpWeekDay & ", " & tmpDateTime
                    
                    End If
                    
                    If LateStatus = False Then
                        lblLate.Visible = False
                        frmStaffAttendance.Height = 3210
                        Frame1.Top = 1920
                        txtReason.Visible = False
                        cmdExit.Enabled = True
                        cmdUpdate.Visible = False
                    Else
                        lblLate.Visible = True
                        lblLate = tmplblLate & USRCLKTime & UsrCLKTimeAMPM & " and enter reason for being late."
                        frmStaffAttendance.Height = 4395
                        Frame1.Top = 3120
                        txtReason.Visible = True
                        cmdExit.Enabled = False
                        cmdUpdate.Visible = True
                    End If
                    
                End If
            End If
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
Dim rst As New ADODB.Recordset
Dim strsql As String
    If txtReason <> "" Then
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
            strsql = ""
            strsql = "Select StaffLogon,StaffName,ClockInTime,ClockInDate, TimeFormat,WeekDay,Reason from StaffAttendance where StaffLogon = '" & Logon & "' and ClockInDate = '" & Format(tmpDateonly, "mm/dd/yyyy") & "'"
            rst.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            If rst.recordCount > 0 Then
                rst!Reason = txtReason
                rst.Update
                MsgBox "Record Updated", vbCritical
                cmdExit.Enabled = True
            End If
    Else
        MsgBox "Please Enter Remark", vbCritical
        'medixmasterconnMaint.Close
        'Set medixmasterconnMaint = Nothing
    End If
End Sub

Private Sub txtReason_LostFocus()
    txtReason = ChkStr(txtReason)
End Sub
