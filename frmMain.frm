VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "j3DWorld"
   ClientHeight    =   6570
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo error_h
    
    'This code could just as easily been implemented in jDXEngine.Render
    Select Case KeyCode
        Case vbKeyEscape
            jDX.EndRender
        Case vbKeyF1
            jDX.Camera.DisplayVector jDX.Camera.position
    End Select
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "frmMain_KeyDown(" & KeyCode & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    jDX.EndRender       'Form is going away, so make sure the engine stops
End Sub



