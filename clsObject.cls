VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnlitObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This object's purpose is to group specific geometry together as a single entity

'The texture to be used when drawn
Dim Texture As Direct3DTexture8

'The primitive type to be used when drawn
Dim mPrimitiveType As PRIMITIVES

'The orientation of this geometry-set
'Please note that this is not the 'World Space' matrix itself, but
'a 'personal' World Space for this particular geometry-set
Dim matWorld As D3DMATRIX

'Unlit Vertex Type used by DirectX (Same as D3DVERTEX)
Private Type VERTEX
    X As Single
    Y As Single
    z As Single
    nx As Single
    ny As Single
    nz As Single
    tu As Single
    tv As Single
End Type

'Array of vertices to use when drawing
Private Vertices() As VERTEX

'Vertex type definition for DirectX
Private Const FVF_UNLITVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)

'Types of Primitives
Public Enum PRIMITIVES
    TriangleList = D3DPT_TRIANGLELIST
    TriangleFan = D3DPT_TRIANGLEFAN
    TraingleStrip = D3DPT_TRIANGLESTRIP
End Enum

'This sub will add a cube given the left-top-back corner, and the right-bottom-front corners
'x1,y1,z1,x2,y2,z2:     The two corners of the cube
'z3,z4:                 The Z coordinates of the left-bottom-front corner and right-top-back corner to determine orientation
Public Sub AddCube(x1 As Single, y1 As Single, z1 As Single, x2 As Single, y2 As Single, z2 As Single, z3 As Single, z4 As Single)
    On Error GoTo error_h
    
    'Just make the faces with the appropriate values
    AddRect x1, y1, z2, x2, y2, z2, z3, z3  'Side 1
    AddRect x1, y1, z1, x2, y1, z2, z3, z4  'Side 2
    AddRect x1, y2, z1, x2, y1, z1, z1, z4  'Side 3
    AddRect x1, y2, z2, x2, y2, z1, z4, z2  'Side 4
    AddRect x1, y1, z1, x1, y2, z2, z4, z3  'Side 5
    AddRect x2, y1, z2, x2, y2, z1, z2, z4  'Side 6

    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsUnlitObject.AddCube")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'This will create a 2D rectangle out of 2 triangles given the top-left and bottom-right coordinates
'x1,y1,z1,x2,y2,z2:     top-left, and bottom-right coords of the rectangle
'z3,z4:                 Z coordinates of bottom-left and top-right corners for orientation
'[strTexture]:          Unused at this time.  Was originally meant for a texture, but now is obsolete.
'                       (This argument WILL be removed in the next version)
Public Sub AddRect(x1 As Single, y1 As Single, z1 As Single, x2 As Single, y2 As Single, z2 As Single, z3 As Single, z4 As Single, Optional strTexture As String)
    On Error GoTo error_h
    
    If IsMissing(strTexture) Then
        AddTriangle x1, y1, z1, x2, y2, z2, x1, y2, z3, 0
        AddTriangle x1, y1, z1, x2, y1, z4, x2, y2, z2, 0
    Else
        AddTriangle x1, y1, z1, x2, y2, z2, x1, y2, z3, 0, strTexture
        AddTriangle x1, y1, z1, x2, y1, z4, x2, y2, z2, 1, strTexture
    End If
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsUnlitObject.AddRect")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'This will create a 2D triangle based off 3 points
'x(1-3),y(1-3),z(1-3):  xyz coordinates of each point in the triangle
'tValue:                1 if this triangle is the top-half of a rectangle. Used for texture mapping
'[strTexture]:          Unused.  Will be removed in next update
Public Sub AddTriangle(x1 As Single, y1 As Single, z1 As Single, x2 As Single, y2 As Single, z2 As Single, x3 As Single, y3 As Single, z3 As Single, tValue As Integer, Optional strTexture As String)
    On Error GoTo error_h
    
    'Vertices to be created
    Dim v1 As VERTEX, v2 As VERTEX, v3 As VERTEX
    'A 'normal' is a vector that describes which way the vertices are facing
    Dim Normal As D3DVECTOR
    
    If IsMissing(strTexture) Then       'Untextured triangle (Not gonna happen with unlit objects)
        'v1 = MakeVertex(x1, y1, z1, 0, 0, 0, 0, 0)
        'v2 = MakeVertex(x2, y2, z2, 0, 0, 0, 0, 0)
        'v3 = MakeVertex(x3, y3, z3, 0, 0, 0, 0, 0)
    ElseIf tValue = 1 Then
        'Top half of the rect (if this is even PART of a rect...)
        v1 = MakeVertex(x1, y1, z1, 0, 0, 0, 0, 0)
        v2 = MakeVertex(x2, y2, z2, 0, 0, 0, 1, 0)
        v3 = MakeVertex(x3, y3, z3, 0, 0, 0, 1, 1)
    Else
        'Bottom half of the theoretical rect...
        v1 = MakeVertex(x1, y1, z1, 0, 0, 0, 0, 0)
        v2 = MakeVertex(x2, y2, z2, 0, 0, 0, 1, 1)
        v3 = MakeVertex(x3, y3, z3, 0, 0, 0, 0, 1)
    End If
    
    'This will get the direction that the triangle is facing ASSUMING that the vertices are
    'in clockwise order from v1 to v2 to v3.  If they are in counter-clockwise order, the
    'normal will point in the opposite direction
    Normal = GetTriangleNormal(v1, v2, v3)
    
    'Copy the normal values to all the created vertices
    With Normal
        v1.nx = .X
        v1.ny = .Y
        v1.nz = .z
        v2.nx = .X
        v2.ny = .Y
        v2.nz = .z
        v3.nx = .X
        v3.ny = .Y
        v3.nz = .z
    End With
    
    'Add these vertices to the array
    mAddVertex v1
    mAddVertex v2
    mAddVertex v3
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsUnlitObject.AddTriangle")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub
'This will add a vertex object to the containing array to be drawn
'x,y,z:     World Space coordinates of this vertex
'nx,ny,nz:  World Space coordinates of which way this vertex is pointing
'tu,tv:     Texture map coordinates of this vertex
'Please note that this sub is meant to add custom shapes to your app, and is NOT used when adding
'cubes/rects to World Space.  It can also be used for custom texture-mapping.
Public Sub AddVertex(X As Single, Y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single)
    On Error GoTo error_h
    
    'Create a vertex
    Dim v As VERTEX
    
    'Assign values
    With v
        .X = X
        .Y = Y
        .z = z
        .nx = nx
        .ny = ny
        .nz = nz
        .tu = tu
        .tv = tv
    End With
    
    'Add the vertex
    mAddVertex v
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsUnlitObject.AddVertex")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub


'This sub's purpose is to take a created vertex, and store it in the Vertices() array
'v: The VERTEX object to store
Private Sub mAddVertex(v As VERTEX)
    On Error GoTo error_h
    
    On Error Resume Next
    ReDim Preserve Vertices(UBound(Vertices) + 1)   'Increase the array by one
    If Err Then     'There were no vertices in the array, so make the first
        Err.Clear
        ReDim Vertices(0)
    End If
    Vertices(UBound(Vertices)) = v  'Set the vertex passed to this sub as the last vertex in the list
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsUnlitObject.AddVertex")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'This sub will actually perform the drawing of all the geometry for this object
'D3DDevice:     Passed from jDXEngine, this is the device to draw upon
Public Sub Draw(D3DDevice As Direct3DDevice8)
    On Error GoTo error_h
    
    'The count of primitives (triangles)
    Dim intCount As Integer
    
    'Use the proper texture for this object.
    'If no texture exists, use default white
    If Not Texture Is Nothing Then
        D3DDevice.SetTexture 0, Texture
    Else
        D3DDevice.SetTexture 0, Nothing
    End If
    
    'Set the world matrix to this object's own matrix so that it can be rotated independantly of the others
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    'Calculate the number of primitives
    Select Case mPrimitiveType
        Case D3DPT_TRIANGLELIST
            intCount = (UBound(Vertices) + 1) / 3
        Case D3DPT_TRIANGLEFAN
            intCount = UBound(Vertices) - 1
        Case D3DPT_TRIANGLESTRIP
            intCount = UBound(Vertices) + 1 - 2
        Case Else
            Stop
            'Need to figure out primitive count for this mPrimitiveType
    End Select
    
    'Perform the draw operation
    D3DDevice.DrawPrimitiveUP mPrimitiveType, intCount, Vertices(0), Len(Vertices(0))
            
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsUnlitObject.Draw")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'This function will compute the 'normal' of a triangle given the 3 vertices.
'Please note that the vertices are assumed to be in clockwise order.
'If they are NOT clockwise, then the vector returned will point in the opposite direction
Private Function GetTriangleNormal(v1 As VERTEX, v2 As VERTEX, v3 As VERTEX) As D3DVECTOR
    On Error GoTo error_h
    
    'Variables used to calculate the normal
    Dim v01 As D3DVECTOR
    Dim v02 As D3DVECTOR
    Dim Normal As D3DVECTOR
    
    'I'm not going to pretend that I understand the mathmatics behind this but , again,
    'this has proven effective from the tutorials that I've read. :)
    
    D3DXVec3Subtract v01, MakeVector(v2.X, v2.Y, v2.z), MakeVector(v1.X, v1.Y, v1.z)
    D3DXVec3Subtract v02, MakeVector(v3.X, v3.Y, v3.z), MakeVector(v1.X, v1.Y, v1.z)
    
    D3DXVec3Cross Normal, v01, v02
    
    D3DXVec3Normalize Normal, Normal
    
    GetTriangleNormal = Normal
    
    Exit Function
error_h:
    Select Case ErrMsg(Err, "clsUnlitObjec.GetTriangleNormal")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Function
    End Select
End Function

'A simple function to create a vector from coordinate data
Private Function MakeVector(X As Single, Y As Single, z As Single) As D3DVECTOR
    With MakeVector
        .X = X
        .Y = Y
        .z = z
    End With
End Function


'Another simple function to create a VERTEX object by copying the values
'Please note that the VERTEX type is EXACTLY the same as D3DVERTEX
'(I had already implemented the VERTEX type before I knew that, and didn't feel like re-coding)
Private Function MakeVertex(X As Single, Y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As VERTEX
    With MakeVertex
        .X = X
        .Y = Y
        .z = z
        .nx = nx
        .ny = ny
        .nz = nz
        .tu = tu
        .tv = tv
    End With
End Function

Property Get PrimitiveType() As PRIMITIVES
    PrimitiveType = mPrimitiveType
End Property

Property Let PrimitiveType(lngNewType As PRIMITIVES)
    mPrimitiveType = lngNewType
End Property
'This sub will rotate the entire object as a single entity
Public Sub RotateY(sngAngle As Single)
    On Error GoTo error_h
    
    Dim matTemp As D3DMATRIX
    
    'Set the temp matrix to default
    D3DXMatrixIdentity matTemp
    
    'Rotate the matrix on the Y axis
    D3DXMatrixRotationY matTemp, sngAngle
    
    'Combine this transformation with any existing transformations
    D3DXMatrixMultiply matWorld, matWorld, matTemp
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsUnlitObject.RotateX")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'This sub stores the texture to be used when drawing this object
Public Sub SetTexture(newTexture As Direct3DTexture8)
    Set Texture = newTexture
End Sub

Property Get WorldMatrix() As D3DMATRIX
    WorldMatrix = matWorld
End Property
Property Let WorldMatrix(matNewWorld As D3DMATRIX)
    matWorld = matNewWorld
End Property

'Make sure the object starts at default orientation
Private Sub Class_Initialize()
    D3DXMatrixIdentity matWorld
End Sub


