VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "jDXEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Ahhh... the powerful and might jDXEngine. ;)
'This is meant to be the mother of all the other 'helper' objects that are contained in this project.
'Note that the other objects CAN be used independently, but will probably need a little bit of re-coding
'to make them work alone.  All helper objects should be able to be accessed through jDXEngine, so declaring
'them seperately is quite unneccessary.  If you have any questions/comments on this object, please direct
'them to jhicks@hsadallas.com.  In case anone is wondering where the 'j' in jDXEngine comes from, my name
'is 'j'oe, and I've gotten into the habit of prefixing my own personal object classes with j. :)

'Declare DirectX objects needed
Dim DX As DirectX8                  'The mother of all DirectX objects
Dim D3D As Direct3D8                'Handles 3D stuffs
Dim D3DX As D3DX8                   'Helper library for meshes and such (will be implemented next release)
Dim D3DDevice As Direct3DDevice8    'Device to draw upon... (almost anything with a .hWnd property)

'Local class variables for use in rendering
Dim bRunning As Boolean             'Quit rendering when False
Dim BackColor As Long               'Background color of the drawing surface

'Camera object
Dim mCamera As New clsCamera

'Keyboard object
Dim mKeyboard As New clsKeyboard

'Lights object(s)
Dim mLights As New clsLights

'Value of pi (4*atn(1)) for any geometry calculations that may need to be performed
Const mPI = 3.14159265358979

'Vertex type definitions to use
Const FVF_LVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)    'Use without lighting
Const FVF_UNLITVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)                    'Use with lighting

'Matrices
Dim matWorld As D3DMATRIX   'World matrix... when you rotate this, you rotate the entire world
Dim matProj As D3DMATRIX    'Projection matrix... used to define the 'lens' for the camera
                            'matProj SHOULD be with the camera object, but until I understand it's use
                            'better, it will remain here.
Dim matTemp As D3DMATRIX    'Temporary matrix used in calculations

'Storage for all textures that may be loaded
Dim colTextures As New collection

'Unlit geometry object collection.  Each is it's own 'entity'
Private UnlitGeometry As New collection


'A simple sub to add a clsUnlitObject object to the collection
Public Sub AddGeometry(UnlitGeo As clsUnlitObject)
    UnlitGeometry.Add UnlitGeo
End Sub

'This function will load a texture from a file, and store it into a collection using a programmer-provided
'key to reference it by.  If the function is successful, it will return the reference name of the texture.
'Otherwise, it will return an empty string to indicate failure
'strFileName:       Full path and filename of the image to be loaded
'strTextureName:    Name used to reference this texture
Public Function AddTexture(strFileName As String, strTextureName As String) As String
    On Error GoTo error_h
    
    'Temporary storage for the texture
    Dim Texture As Direct3DTexture8
    
    'First make sure the file exists...
    If Dir(strFileName) = "" Then Exit Function
    
    'Validate the texture name
    If strTextureName = "" Or IsInCollection(colTextures, strTextureName) Then Exit Function
        
    'Create the texture
    Set Texture = D3DX.CreateTextureFromFile(D3DDevice, strFileName)
    
    'Make sure that the texture was created properly
    If Texture Is Nothing Then Exit Function
    
    'Add this texture to the collection
    colTextures.Add Texture, strTextureName
    
    'If all went well, return the name of the texture
    AddTexture = strTextureName
    
    Exit Function
error_h:
    Select Case ErrMsg(Err, "jDXEngine.AddTexture(" & strFileName & "," & strTextureName & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Function
    End Select
End Function

'This allows programmer-access to the camera object
Property Get Camera() As clsCamera
    Set Camera = mCamera
End Property

'Returns a texture when given the reference name of an existing texture
'This function will return NOTHING if the reference name provided does not exist
Public Function GetTexture(strTexture As String) As Direct3DTexture8
    On Error GoTo error_h
    
    If IsInCollection(colTextures, strTexture) Then
        Set GetTexture = colTextures(strTexture)
    Else
        Set GetTexture = Nothing
    End If
    
    Exit Function
error_h:
    Select Case ErrMsg(Err, "jDXEngine.GetTexture(" & strTexture & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Function
    End Select
End Function

'This function will determine if an item with a specific key exists within a collection
Private Function IsInCollection(collection As collection, strKey As String) As Boolean
    On Error GoTo error_h
    
    On Error Resume Next
    Dim X As Direct3DTexture8
    Set X = collection(strKey)
    If Err Then
        Err.Clear
        IsInCollection = False
    Else
        IsInCollection = True
    End If
    
    Exit Function
error_h:
    Select Case ErrMsg(Err, "jDXEngine.IsInCollection")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Function
    End Select
End Function

'r=red value(0-255),g=green,b=blue
Public Function jRGB(r As Integer, g As Integer, b As Integer) As Long
    On Error GoTo error_h
    
    'I've noticed when working with DX8 that the long color code values it uses
    'are exactly the same as VB's RGB() function except that the R and B values
    'are transposed, so I wrote this little helper function to help out with that.
    jRGB = RGB(b, g, r)
    
    Exit Function
error_h:
    Select Case ErrMsg(Err, "jDXEngine.DXRGB(" & r & "," & g & "," & b & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Function
    End Select
End Function

Public Sub EndRender()
    On Error GoTo error_h
    
    'All we need to do to stop rendering is set this little variable here to false. :)
    bRunning = False
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "jDXEngine.EndRender")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'hWnd = the handle to the window that you want to draw in (i THINK it can be anything with a hWnd property, but don't quote me on that)
'HardwareDevice = true if you want to use a 3d accelorator card... false if not
'vtrCamPos = a vector desribing the initial position of the camera
'lngBackColor = the background color when rendering
'blnUseLighting = True to enable lighting
'blnCullCCW = True to enable back-face removal
Public Function InitWindowed(hWnd As Long, HardwareDevice As Boolean, vtrCamPos As D3DVECTOR, lngBackColor As Long, blnUseLighting As Boolean, blnCullCCW As Boolean) As Boolean
    On Error GoTo error_h
    
    'Boolean values default to false, but just to be safe...
    InitWindowed = False
    
    'Declare variables to store display information
    Dim d3dpp As D3DPRESENT_PARAMETERS
    Dim DispMode As D3DDISPLAYMODE
    
    'Create the main 3 objects (D3DX8 is NOT one of the main 3... D3DDevice is the 3rd, but we must configure it first)
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate
    
    'The D3DX8 object is NOT required for use in DX8, but it contains a LOT of helper functions you may need
    'so it doesn't hurt to declare it for when you need it
    Set D3DX = New D3DX8

    'Here we are getting the current display mode from the default diaplay adapter (your video card)
    'The current display information is stored in the second parameter "DispMode"
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    'D3DPRESENT_PARAMETERS objects show which options are (dis/en)abled and/or how they are configured.
    With d3dpp
        .BackBufferFormat = DispMode.Format     'The BackBuffer is what is actually drawn upon
        .Windowed = 1                           'We want to draw in a window as opposed to full screen
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC  'This says to re-draw when the monitor does
        .EnableAutoDepthStencil = 1             'Enable auto-depth perception
        .AutoDepthStencilFormat = D3DFMT_D16    '16bit depth
    End With
    
    'This creates the 3rd main DX8 object
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, IIf(HardwareDevice, D3DDEVTYPE_HAL, D3DDEVTYPE_REF), hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    
    'Now we can configure the D3DDevice
    With D3DDevice
        .SetVertexShader IIf(blnUseLighting, FVF_UNLITVERTEX, FVF_LVERTEX)  'This is the vertex definition we're using
        .SetRenderState D3DRS_LIGHTING, IIf(blnUseLighting, 1, 0)           'Turn lighting on/off (off is default, but just to be sure)
                                                                            '(We're turning the lights off because the vertices are
                                                                            ' already lit)
        If blnCullCCW Then
            .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW                     'Enable culling (don't show back-side of geometry)
        Else
            .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE                    'Disable culling (show back-sides)
        End If
        
        .SetRenderState D3DRS_ZENABLE, 1                                    'Enable depth perception
        
        If blnUseLighting Then
            .SetRenderState D3DRS_AMBIENT, &H202020                         'Dark grey ambient light
        End If
    End With
    
    'I won't go into what a matrix is, but when you Identity one, it's like a 'reset' to all default values
    D3DXMatrixIdentity matWorld
    D3DDevice.SetTransform D3DTS_WORLD, matWorld    'If you were to rotate this matrix first, the entire
                                                    'drawing would be rotated as well
    
    'Store D3DDevice in the camera
    mCamera.SetDevice D3DDevice
    
    'Setup the camera with default values
    mCamera.SetupCamera MakeVector(0, 0, 1), MakeVector(0, 1, 0), vtrCamPos
    
    'Setup the keyboard
    mKeyboard.Init DX, hWnd
    
    'Setup the lights
    mLights.Init D3DDevice
    
    'This matrix is for the 'lens' of the camera
    D3DXMatrixPerspectiveFovLH matProj, mPI / 4, 1, 0.1, 1000
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
        
    'Just a long to store the user-chosen background color
    BackColor = lngBackColor
    
    'Everything went well
    InitWindowed = True
    
    'Since everything's configured, we can start to Render() whenever the programmer says so
    bRunning = True
    
    Exit Function
error_h:
    Select Case ErrMsg(Err, "jDXEngine.InitWindowed(" & hWnd & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Function
    End Select
End Function
'Allow programmer access to the keyboard object
Property Get Keyboard() As clsKeyboard
    Set Keyboard = mKeyboard
End Property

'Allow programmer-access to the lights
Property Get Lights() As clsLights
    Set Lights = mLights
End Property

'Helper function to createa D3DCOLORVALUE for functions that require one
Public Function MakeColorValue(a As Integer, r As Integer, g As Integer, b As Integer) As D3DCOLORVALUE
    On Error GoTo error_h
    
    With MakeColorValue
        .a = a / 255
        .r = r / 255
        .g = g / 255
        .b = b / 255
    End With
    
    Exit Function
error_h:
    Select Case ErrMsg(Err, "jDX.MakeColorValue(" & r & "," & g & "," & b & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Function
    End Select
End Function

'ErrObj = the actual error
'strProc = a string containing the name of the procedure the error occured in
'(sometimes i send along parameter information as well if applicable...)
Private Function ErrMsg(ErrObj As ErrObject, strProc As String)
    'This is something of my own invention.  I use it for debugging purposes, and you should see a call
    'to it in almost every routine save the simplest.  What it does is LOG the error (because users never
    'bother to remember what the error said, or even it's error number), and present a message box with
    'the error information, and the ROUTINE that the error occurred in! (VERY useful for debugging).  All
    'of this information is put into a .txt file called "error.log" which can be viewed in notepad.
    
    Dim intFreeFile As Integer
    
    intFreeFile = FreeFile
    
    Open App.Path & "\error.log" For Append As #intFreeFile
        Print #intFreeFile, Date
        Print #intFreeFile, Time
        Print #intFreeFile, " "
        Print #intFreeFile, ErrObj.Number
        Print #intFreeFile, ErrObj.Description
        Print #intFreeFile, "(" & strProc & ")"
    Close intFreeFile
    
    Select Case MsgBox(ErrObj.Number & vbCrLf & ErrObj.Description, vbExclamation + vbAbortRetryIgnore, strProc)
        Case vbRetry
            ErrMsg = vbRetry
        Case vbIgnore
            ErrMsg = vbIgnore
        Case Else
            ErrMsg = vbAbort
    End Select
    
End Function

Public Function MakeVector(X As Single, Y As Single, z As Single) As D3DVECTOR
    'Again, copy the supplied values to a D3DVECTOR structure
    With MakeVector
        .X = X
        .Y = Y
        .z = z
    End With
End Function


Property Get PI() As Single
    'In case the programmer needs to play with PI...
    PI = mPI
End Property


Public Sub Render()
    On Error GoTo error_h
    
    Dim geo As clsUnlitObject
    
    'Keep running this loop until the programmer says not to (bRunning=false will stop the loop)
    Do While bRunning
        'Clear the screen so we can redraw
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, BackColor, 1#, 0
        
        'Get the keyboard state
        mKeyboard.GetKeys
        
        'This would be better if it were somewhere else because i don't like hardcoding in the render
        'routine, but anywhere else and it was going OH so slow....
        If mKeyboard.IsKeyDown(DIK_UP) Then mCamera.MoveForward 0.2
        If mKeyboard.IsKeyDown(DIK_DOWN) Then mCamera.MoveForward -0.2
        If mKeyboard.IsKeyDown(DIK_RIGHT) Then mCamera.RotateY 0.04
        If mKeyboard.IsKeyDown(DIK_LEFT) Then mCamera.RotateY -0.04
        If mKeyboard.IsKeyDown(DIK_PERIOD) Then mCamera.StrafeRight 0.2
        If mKeyboard.IsKeyDown(DIK_COMMA) Then mCamera.StrafeRight -0.2
        
        'Make the original light eminate from the camera.  That way you only need one light
        'This will cause an error if lighting is disabled
        mLights.MoveLight 0, mCamera.position
        
        'Begin scene
        D3DDevice.BeginScene
            'Rendering begins here:
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            'Draw each geometry as a separate entity
            For Each geo In UnlitGeometry
                geo.Draw D3DDevice
            Next geo
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            'End of rendering
        'End of scene
        D3DDevice.EndScene
        
        'Present the scene
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        
        'Free up the processor so the computer don't choke! :)
        DoEvents
    Loop
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "jDXEngine.Renter")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            bRunning = False
            Exit Sub
    End Select
End Sub

Public Sub ResetWorld()
    On Error GoTo error_h
    
    'To reset a matrix to it's default orientation, you just need to 'identity' it:
    D3DXMatrixIdentity matWorld
    
    'Now that it's been reset, it can now be used as the new world matrix
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "jDXEngine.ResetWorld")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

Public Sub RotateWorldX(sngAngle As Single)
    On Error GoTo error_h
    
    'This sub will rotate the entire world around the X axis (vertical rotation)
    'The reason for matTemp as that if there were any OTHER rotations, then we need to
    ' multiply this temporary matrix with the original world matrix to get a combined
    ' effect.  Note that this is DIFFERENT from changing the camera position.
    D3DXMatrixRotationX matTemp, -sngAngle
    
    'Multiply the world matrix by the temporary matrix to combine this effect with
    ' any other effects that may have occured
    D3DXMatrixMultiply matWorld, matWorld, matTemp
    
    'Set the new world matrix:
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "RotateWorldX(" & sngAngle & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

Public Sub RotateWorldY(sngAngle As Single)
    On Error GoTo error_h
    
    'This sub will rotate the entire world around the Y axis (horizontal rotation)
    'The reason for matTemp as that if there were any OTHER rotations, then we need to
    ' multiply this temporary matrix with the original world matrix to get a combined
    ' effect.  Note that this is DIFFERENT from changing the camera position.
    D3DXMatrixRotationY matTemp, -sngAngle
    
    'Multiply the world matrix by the temporary matrix to combine this effect with
    ' any other effects that may have occured
    D3DXMatrixMultiply matWorld, matWorld, matTemp
    
    'Set the new world matrix:
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "RotateWorldY(" & sngAngle & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select

End Sub


Private Sub Class_Terminate()
    'Theoretically, VB will do all of this for you, BUT, just to be safe...
    
    On Error Resume Next       'This is so if the class goes out of scope before the InitWindowed function is called
    Set D3DX = Nothing
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set DX = Nothing
End Sub


