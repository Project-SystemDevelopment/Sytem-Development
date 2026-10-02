VERSION 5.00
Begin VB.Form frmPrintScreen 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleMode       =   0  'User
   ScaleWidth      =   30156.92
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPrintScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mcsClassName As String = "frmPrintScreen"

Public Sub PrintBitMap(pbmpBitMap As Variant) ', ByVal pintHeight As Integer, ByVal pintWidth As Integer)
    Const csMethodName As String = "PrintBitMap"

    On Error GoTo ErrorHandler

    With Me
        .Visible = False
        .Height = Screen.Height + 300 'pintHeight + 120
        .Width = Screen.Width + 150 'pintWidth + 120
        .Picture = pbmpBitMap
        .PrintForm
    End With
    Unload Me

    Exit Sub

ErrorHandler:
    '--- Add your error handler here:
    MsgBox ERR.Number & " " & ERR.Source & vbCrLf & ERR.Description, vbExclamation, mcsClassName & "." & csMethodName
End Sub



