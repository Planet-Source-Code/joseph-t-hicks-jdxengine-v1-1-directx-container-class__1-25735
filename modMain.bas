Attribute VB_Name = "modMain"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Programmed by Joseph Hicks
'
'jDXEngine v1.1
'
'Canges from v1.0:
'   - Most camera, light, keyboard, and geometry functions
'     have been coded as separate objects and can be accessed
'     through the jDXEngine's properties/methods
'   - The use of textures, keyboard and lights have been
'     implemented
'   - Comments have been revised
'
'To all who find this useful/interesting:
'   If you could just take the time to send me a quick
'   email to jhicks@hsadallas.com and say if you liked
'   this or not, and maybe suggest ways to improve it,
'   it would make all the time spent coding these objects
'   more than worth my effort. :)

'This is the main engine's object
Public jDX As New jDXEngine

' This sub doesn't HAVE to be off on it's own, but was done that way to keep the geometry creation
' separate from the jDXEngine initialization
Public Sub GenerateWorld()
    On Error GoTo error_h
    
    'A clsUnlitObject can store any number of primitives(triangles and such) together as a
    'single entity, so if you wanted to perform a transformation on a complete object, but
    'not the entire world, then it should be a much simpler matter.
    '(Ok, so X wasn't a good variable name, but you get the idea.)
    Dim X As New clsUnlitObject
    
    With X
        .PrimitiveType = TriangleList       'These are the same primitive types that DX8 uses
        .AddRect -3, 1, 0, -1, -1, 0, 0, 0  'Routines to add rectangles at the specified corners
        .AddRect 1, 1, 0, 3, -1, 0, 0, 0    '(The parameters will be explained in greater detail
        .AddRect -3, 1, 6, -3, -1, 4, 6, 4  ' under the .AddRect definition)
        .AddRect -3, 1, 4, -3, -1, 2, 4, 2
        .AddRect -3, 1, 2, -3, -1, 0, 2, 0
        .AddRect 3, 1, 0, 3, -1, 2, 0, 2
        .AddRect 3, 1, 2, 3, -1, 4, 2, 4
        .AddRect 3, 1, 4, 3, -1, 6, 4, 6
        .AddRect -3, 3, 6, -3, 1, 4, 6, 4
        .AddRect -3, 3, 4, -3, 1, 2, 4, 2
        .AddRect -3, 3, 2, -3, 1, 0, 2, 0
        .AddRect 3, 3, 0, 3, 1, 2, 0, 2
        .AddRect 3, 3, 2, 3, 1, 4, 2, 4
        .AddRect 3, 3, 4, 3, 1, 6, 4, 6
        .AddRect 3, 1, 6, 1, -1, 6, 6, 6
        .AddRect 1, 1, 6, -1, -1, 6, 6, 6
        .AddRect -1, 1, 6, -3, -1, 6, 6, 6
        .AddRect 3, 3, 6, 1, 1, 6, 6, 6
        .AddRect 1, 3, 6, -1, 1, 6, 6, 6
        .AddRect -1, 3, 6, -3, 1, 6, 6, 6
        .SetTexture jDX.GetTexture("brownbrickwall")    'This will set the texture to be used when
                                                        'drawing this geometry.  jDX.GetTexture will
                                                        'return a previously loaded texture, or NOTHING
                                                        'if a texture by that name hasn't been loaded.
    End With
    jDX.AddGeometry X       'Add this 'object' to the main collection.  After it is added, it WILL be
                            'drawn when the engine is told to .Render
    
    Set X = New clsUnlitObject  'The previous object has already been added, so this one can be reset anew
    With X
        .PrimitiveType = D3DPT_TRIANGLELIST     '(This is just to show that the primitive type IS the same)
        .AddRect -1, 1, 0, 1, -1, 0, 0, 0
        .AddRect 1, 1, 0, -1, -1, 0, 0, 0
        .SetTexture jDX.GetTexture("arch")
    End With
    jDX.AddGeometry X
    
    Set X = New clsUnlitObject
    With X
        .PrimitiveType = TriangleList
        .AddRect -3, 3, 0, -1, 1, 0, 0, 0
        .AddRect -1, 3, 0, 1, 1, 0, 0, 0
        .AddRect 1, 3, 0, 3, 1, 0, 0, 0
        .SetTexture jDX.GetTexture("vertical")
    End With
    jDX.AddGeometry X
    
    Set X = New clsUnlitObject
    With X
        .PrimitiveType = TriangleList
        .AddRect -5, -1, 0, -1, -1, -2, -2, 0
        .AddRect 1, -1, 0, 5, -1, -2, -2, 0
        .AddRect -5, -1, 8, -3, -1, 0, 0, 8
        .AddRect 3, -1, 8, 5, -1, 0, 0, 8
        .AddRect -5, -1, 8, 5, -1, 6, 6, 8
        .SetTexture jDX.GetTexture("grass")
    End With
    jDX.AddGeometry X
    
    Set X = New clsUnlitObject
    With X
        .PrimitiveType = TriangleList
        .AddRect -1, -1, 0, 1, -1, -6, -6, 0
        .SetTexture jDX.GetTexture("sidewalk")
    End With
    jDX.AddGeometry X
    
    Set X = New clsUnlitObject
    With X
        .PrimitiveType = TriangleFan                'This is different than specifying a TriangleList
        .AddVertex -1, 4, 0, 0, 0, -1, 0, 0.5       'Vertices are added manually to create other geometry
        .AddVertex -0.33, 5, 0, 0, 0, -1, 0.33, 0   'besides rectangles and cubes.  Functions could be
        .AddVertex 0.33, 5, 0, 0, 0, -1, 0.66, 0    'written to facilitate hexagons, octagons, and such
        .AddVertex 1, 4, 0, 0, 0, -1, 1, 0.5        'but for time's sake, they were added manually.  The
        .AddVertex 0.33, 3, 0, 0, 0, -1, 0.66, 1    'parameters will be explained in the .AddVertex function
        .AddVertex -0.33, 3, 0, 0, 0, -1, 0.33, 1   'definition.
        .SetTexture jDX.GetTexture("sign")
    End With
    jDX.AddGeometry X
    
    Set X = New clsUnlitObject
    With X
        .PrimitiveType = D3DPT_TRIANGLEFAN          '(This will be the backside of the hexagon sign on top
        .AddVertex -0.33, 3, 0, 0, 0, 1, 0.33, 1    ' of the building)
        .AddVertex 0.33, 3, 0, 0, 0, 1, 0.66, 1
        .AddVertex 1, 4, 0, 0, 0, 1, 1, 0.5
        .AddVertex 0.33, 5, 0, 0, 0, 1, 0.66, 0
        .AddVertex -0.33, 5, 0, 0, 0, 1, 0.33, 0
        .AddVertex -1, 4, 0, 0, 0, 1, 0, 0.5
        .SetTexture jDX.GetTexture("wood")
    End With
    jDX.AddGeometry X
    
    Set X = New clsUnlitObject
    With X
        .PrimitiveType = D3DPT_TRIANGLELIST
        .AddRect -3, -1, 6, -2, -1, 5, 5, 6
        .AddRect -2, -1, 6, -1, -1, 5, 5, 6
        .AddRect -1, -1, 6, 0, -1, 5, 5, 6
        .AddRect 0, -1, 6, 1, -1, 5, 5, 6
        .AddRect 1, -1, 6, 2, -1, 5, 5, 6
        .AddRect 2, -1, 6, 3, -1, 5, 5, 6
        .AddRect -3, -1, 5, -2, -1, 4, 4, 5
        .AddRect -2, -1, 5, -1, -1, 4, 4, 5
        .AddRect -1, -1, 5, 0, -1, 4, 4, 5
        .AddRect 0, -1, 5, 1, -1, 4, 4, 5
        .AddRect 1, -1, 5, 2, -1, 4, 4, 5
        .AddRect 2, -1, 5, 3, -1, 4, 4, 5
        .AddRect -3, -1, 4, -2, -1, 3, 3, 4
        .AddRect -2, -1, 4, -1, -1, 3, 3, 4
        .AddRect -1, -1, 4, 0, -1, 3, 3, 4
        .AddRect 0, -1, 4, 1, -1, 3, 3, 4
        .AddRect 1, -1, 4, 2, -1, 3, 3, 4
        .AddRect 2, -1, 4, 3, -1, 3, 3, 4
        .AddRect -3, -1, 3, -2, -1, 2, 2, 3
        .AddRect -2, -1, 3, -1, -1, 2, 2, 3
        .AddRect -1, -1, 3, 0, -1, 2, 2, 3
        .AddRect 0, -1, 3, 1, -1, 2, 2, 3
        .AddRect 1, -1, 3, 2, -1, 2, 2, 3
        .AddRect 2, -1, 3, 3, -1, 2, 2, 3
        .AddRect -3, -1, 2, -2, -1, 1, 1, 2
        .AddRect -2, -1, 2, -1, -1, 1, 1, 2
        .AddRect -1, -1, 2, 0, -1, 1, 1, 2
        .AddRect 0, -1, 2, 1, -1, 1, 1, 2
        .AddRect 1, -1, 2, 2, -1, 1, 1, 2
        .AddRect 2, -1, 2, 3, -1, 1, 1, 2
        .AddRect -3, -1, 1, -2, -1, 0, 0, 1
        .AddRect -2, -1, 1, -1, -1, 0, 0, 1
        .AddRect -1, -1, 1, 0, -1, 0, 0, 1
        .AddRect 0, -1, 1, 1, -1, 0, 0, 1
        .AddRect 1, -1, 1, 2, -1, 0, 0, 1
        .AddRect 2, -1, 1, 3, -1, 0, 0, 1
        .SetTexture jDX.GetTexture("rug")
    End With
    jDX.AddGeometry X
    
    Set X = New clsUnlitObject
    With X
        .PrimitiveType = D3DPT_TRIANGLELIST
        .AddRect -3, 3, 0, -3, -1, 3, 0, 3
        .AddRect -3, 3, 3, -3, -1, 6, 3, 6
        .AddRect -3, 3, 6, 0, -1, 6, 6, 6
        .AddRect 0, 3, 6, 3, -1, 6, 6, 6
        .AddRect 3, 3, 6, 3, -1, 3, 6, 3
        .AddRect 3, 3, 3, 3, -1, 0, 3, 0
        .AddRect 3, 3, 0, 1, -1, 0, 0, 0
        .AddRect -1, 3, 0, -3, -1, 0, 0, 0
        .AddRect 1, 3, 0, -1, 1, 0, 0, 0
        .SetTexture jDX.GetTexture("wallpaper")
    End With
    jDX.AddGeometry X
            
    Set X = New clsUnlitObject
    With X
        .PrimitiveType = TriangleList
        .SetTexture jDX.GetTexture("grass")
        .AddRect -5, 0, -2, -3, -1, -2, -2, -2
        .AddRect -3, 0, -2, -1, -1, -2, -2, -2
        .AddRect 1, 0, -2, 3, -1, -2, -2, -2
        .AddRect 3, 0, -2, 5, -1, -2, -2, -2
        .AddRect 5, 0, -2, 3, -1, -2, -2, -2
        .AddRect 3, 0, -2, 1, -1, -2, -2, -2
        .AddRect -1, 0, -2, -3, -1, -2, -2, -2
        .AddRect -3, 0, -2, -5, -1, -2, -2, -2
        .AddRect -5, 0, -2, -5, -1, 0, -2, 0
        .AddRect -5, 0, 0, -5, -1, 2, 0, 2
        .AddRect -5, 0, 2, -5, -1, 4, 2, 4
        .AddRect -5, 0, 4, -5, -1, 6, 4, 6
        .AddRect -5, 0, 6, -5, -1, 8, 6, 8
        .AddRect -5, 0, 8, -3, -1, 8, 8, 8
        .AddRect -3, 0, 8, -1, -1, 8, 8, 8
        .AddRect -1, 0, 8, 1, -1, 8, 8, 8
        .AddRect 1, 0, 8, 3, -1, 8, 8, 8
        .AddRect 3, 0, 8, 5, -1, 8, 8, 8
        .AddRect 5, 0, 8, 5, -1, 6, 8, 6
        .AddRect 5, 0, 6, 5, -1, 4, 6, 4
        .AddRect 5, 0, 4, 5, -1, 2, 4, 2
        .AddRect 5, 0, 2, 5, -1, 0, 2, 0
        .AddRect 5, 0, 0, 5, -1, -2, 0, -2
        .AddRect 5, 0, -2, 5, -1, 0, -2, 0
        .AddRect 5, 0, 0, 5, -1, 2, 0, 2
        .AddRect 5, 0, 2, 5, -1, 4, 2, 4
        .AddRect 5, 0, 4, 5, -1, 6, 4, 6
        .AddRect 5, 0, 6, 5, -1, 8, 6, 8
        .AddRect 5, 0, 8, 3, -1, 8, 8, 8
        .AddRect 3, 0, 8, 1, -1, 8, 8, 8
        .AddRect 1, 0, 8, -1, -1, 8, 8, 8
        .AddRect -1, 0, 8, -3, -1, 8, 8, 8
        .AddRect -3, 0, 8, -5, -1, 8, 8, 8
        .AddRect -5, 0, 8, -5, -1, 6, 8, 6
        .AddRect -5, 0, 6, -5, -1, 4, 6, 4
        .AddRect -5, 0, 4, -5, -1, 2, 4, 2
        .AddRect -5, 0, 2, -5, -1, 0, 2, 0
        .AddRect -5, 0, 0, -5, -1, -2, 0, -2
    End With
    jDX.AddGeometry X
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "GenerateWorld")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

Public Sub LoadTextures()
    On Error GoTo error_h
    
    Dim strPath As String
    
    strPath = App.Path & "\textures\"
    
    With jDX
        'The purpose of .AddTexture is to load an image (.jpg, .bmp, and i THINK .gif) to be placed
        'on your geometry.  The filename and a unique name must be supplied.  The .AddTexture function
        'will return the name supplied if successful and will return an empty string if unsuccessful.
        'You will use the name supplied (2nd parameter) to refer to the texture when needed.
        .AddTexture strPath & "brick_wall.jpg", "brickwall"
        .AddTexture strPath & "brownbrick_wall.jpg", "brownbrickwall"
        .AddTexture strPath & "facade_arch.jpg", "arch"
        .AddTexture strPath & "facade_steel.jpg", "steel"
        .AddTexture strPath & "facade_stone.jpg", "stone"
        .AddTexture strPath & "facade_vertical.jpg", "vertical"
        .AddTexture strPath & "grass.bmp", "grass"
        .AddTexture strPath & "parkgarage.jpg", "garage"
        .AddTexture strPath & "parkinglot.jpg", "parkinglot"
        .AddTexture strPath & "paving.jpg", "pavement"
        .AddTexture strPath & "road_lined.jpg", "linedroad"
        .AddTexture strPath & "road_plain.jpg", "road"
        .AddTexture strPath & "road_sidewalk.jpg", "sidewalkroad"
        .AddTexture strPath & "sidewalk.jpg", "sidewalk"
        .AddTexture strPath & "stone_wall.jpg", "stonewall"
        .AddTexture strPath & "water.jpg", "water"
        .AddTexture strPath & "sign.bmp", "sign"
        .AddTexture strPath & "wood.jpg", "wood"
        .AddTexture strPath & "rug.jpg", "rug"
        .AddTexture strPath & "wallpaper.jpg", "wallpaper"
    End With
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "LoadTextures")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

Public Sub Main()
    On Error GoTo error_h       'Begin basic error handling.  This is beyond the scope of this code.
    
    frmMain.Show                'Make sure the form is visible before trying to do anything with it.
    DoEvents
    
    frmMain.Print "Loading images..."   'Simple alert to the user
    
    With jDX
        If Not .InitWindowed(frmMain.hWnd, True, .MakeVector(0, -0.5, -15), vbBlack, True, True) Then
            MsgBox "Initialization failed"
            Stop
        End If
        
        'Add all textures BEFORE creating geometry
        LoadTextures
        
        'Start generating world (geometry)
        GenerateWorld
        
        'Init some lights
        .Lights.SetupLight .Lights.AddNewLight, .MakeVector(0, 0, -8), .MakeColorValue(255, 255, 255, 255), 20, True
        
        'Start rendering
        frmMain.Cls 'This is so the "Loading images..." doesn't get stuck on the screen
        .Render
    End With
    
    Unload frmMain
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "Main")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

' Public generic error handler...
'
Public Function ErrMsg(ErrObj As ErrObject, strProc As String)
    
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


'This sub is meant to be run from the immediate window.  It will open a Windows Explorer window to the
'App.Path so you don't have to navigate to it. :)
Public Sub OpenRoot()
    Shell "explorer " & App.Path, vbNormalFocus
End Sub


