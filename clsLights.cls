VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'An array of D3DLIGHT8 objects (All the lights to be used)
Private Lights() As D3DLIGHT8

'Types of lights
Public Enum jDXLightType
    Point = D3DLIGHT_POINT
    Directional = D3DLIGHT_DIRECTIONAL
    Spotlight = D3DLIGHT_SPOT
End Enum

'The D3DDevice that is used by the jDXEngine object
Private D3DDevice As Direct3DDevice8

'This will add another light object (NOT initialized in any way) to the array, and return the
'array element associated with it
Public Function AddNewLight() As Integer
    On Error Resume Next
    
    ReDim Preserve Lights(UBound(Lights) + 1)   'Extend the light array by one
    If Err Then         'If there's an error, then there are NO lights existing, so make the first one
        ReDim Lights(0)
        Err.Clear
    End If
    
    AddNewLight = UBound(Lights)
    
End Function


'The only initialization needed for this object is a reference to the D3DDevice used by
'the jDXEngine object
Public Sub Init(Device As Direct3DDevice8)
    Set D3DDevice = Device
End Sub

'This function is used to configure a defined light
'LightIndex:        The array element to refer to a specific light
'vtrLightPosition:  A vector to position the light in World Space
'color:             The color of the light
'range:             How far the light travels
'blnActive:         Is this light turned on or off?
'Please note - This function is currently configured for POINT lights only.  It would need to
'accept values for many other parameters (including the type) for other types of lights
Public Sub SetupLight(LightIndex As Integer, vtrLightPosition As D3DVECTOR, color As D3DCOLORVALUE, range As Single, blnActive As Boolean)
    On Error GoTo error_h
    
    Dim Mtrl As D3DMATERIAL8    'The material used for drawing by default
    Dim Col As D3DCOLORVALUE    'A color value used with the material
    
    With Col
        .a = 1  'Ambient color
        .r = 1  'Red color
        .g = 1  'Green color
        .b = 1  'Blue color
    End With
    
    'Set the material's color values
    Mtrl.Ambient = Col
    Mtrl.diffuse = Col
    
    'Set the material to be used with D3DDevice
    D3DDevice.SetMaterial Mtrl
    
    'Configure this particular light
    With Lights(LightIndex)
        .Type = D3DLIGHT_POINT          'Currently only supporting POINT type lights
        .diffuse = color                'Color of the light
        .range = range                  'How far the light will travel
        .position = vtrLightPosition    'The position of the light
        .Attenuation1 = 0.05            'I'm not quite sure what the attenuations(0,1,2) do, but I do know
                                        'that they should not ALL be 0.0
    End With
    
    'Store the light in the D3DDevice
    D3DDevice.SetLight LightIndex, Lights(LightIndex)

    'Turn on/off the light
    D3DDevice.LightEnable LightIndex, IIf(blnActive, 1, 0)

    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsLights.SetupLights")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub
'This will turn a light on/off depending on the value of blnActive
'LightIndex:    Array element index of the light to set
'blnActive:     True=ON, False=OFF
Public Sub ActivateLight(LightIndex As Integer, blnActive As Boolean)
    On Error GoTo error_h
    
    D3DDevice.LightEnable LightIndex, IIf(blnActive, 1, 0)
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsLights.ActivateLight(" & LightIndex & "," & blnActive & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub


'This routine will change a light's color regardless if its on or off
Public Sub ChangeLightColor(LightIndex As Integer, r As Integer, g As Integer, b As Integer)
    On Error GoTo error_h
    
    Dim sngR As Single
    Dim sngG As Single
    Dim sngB As Single
    
    'Convert 0-255 values to 0.0-1.0 values
    sngR = r / 255
    sngG = g / 255
    sngB = b / 255
    
    'Change the light's color
    With Lights(LightIndex)
        .diffuse.r = sngR
        .diffuse.g = sngG
        .diffuse.b = sngB
    End With
    
    'Store the light
    D3DDevice.SetLight 0, Lights(LightIndex)
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsLights.ChangeLightColor")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub
'This will change a light's position
'LightIndex:        Array element index of light to configure
'vtrNewPosition:    The new position of the light in World Space
Public Sub MoveLight(LightIndex As Integer, vtrNewPosition As D3DVECTOR)
    On Error GoTo error_h
    
    'This is quite simple, really.  Just apply the new position vector to the light and reset the light
    Lights(LightIndex).position = vtrNewPosition
    D3DDevice.SetLight 0, Lights(LightIndex)
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsLights.MoveLight")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub


