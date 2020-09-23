VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Used for keyboard input
Private DI As DirectInput8
Private DIDevice As DirectInputDevice8
Private DIState As DIKEYBOARDSTATE
Private KeyState(0 To 255) As Boolean
Private Const BufferSize As Long = 20
Dim pBuffer(0 To 20) As DIDEVICEOBJECTDATA


'This will poll the keyboard for any keys pressed
Public Sub GetKeys()
    On Error GoTo error_h
    
    DIDevice.GetDeviceStateKeyboard DIState
    On Error Resume Next
    DIDevice.GetDeviceData pBuffer, DIGDD_DEFAULT
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsKeyboard.GetKeys")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'The initialization routine for the keyboard
'DX:       The DirectX8 object used by jDXEngine
'hWnd:     The handle to the window used to accept the keyboard input
Public Sub Init(DX As DirectX8, hWnd As Long)
    On Error GoTo error_h
    
    Dim I As Long
    Dim DevProp As DIPROPLONG
    Dim DevInfo As DirectInputDeviceInstance8
    
    Set DI = DX.DirectInputCreate
    Set DIDevice = DI.CreateDevice("GUID_SysKeyboard")
    
    DIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIDevice.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    
    DevProp.lHow = DIPH_DEVICE
    DevProp.lData = BufferSize
    
    DIDevice.Acquire
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsKeyboard.Init")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub


'This function is used to determine if a specific key is pressed
Public Function IsKeyDown(dxKeyCode As Long) As Boolean
    On Error GoTo error_h
    
    If DIState.Key(dxKeyCode) Then
        IsKeyDown = True
    Else
        IsKeyDown = False
    End If
    
    Exit Function
error_h:
    Select Case ErrMsg(Err, "clsKeyboard.IsKeyDown(" & dxKeyCode & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Function
    End Select
End Function


