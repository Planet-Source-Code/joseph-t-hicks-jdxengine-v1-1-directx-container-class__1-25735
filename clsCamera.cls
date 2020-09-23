VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'These 4 vectors will describe the camera's orientation.
'Please note that the axis functions for the camera object
'will operate on THESE vectors, and not the World Space axis.
Private vtrPosition As D3DVECTOR    'The camera's position in World Space
Private vtrUp As D3DVECTOR          'Which way is up? (Y Axis)
Private vtrLook As D3DVECTOR        'Which way is forward? (Z Axis)
Private vtrRight As D3DVECTOR       'Which way is right? (X Axis)

'This is a temporary handle to jDXEngine's Direct3DDevice8
Private D3DDevice As Direct3DDevice8


'This function will take into account any floating point math errors that may occur from
'rotating the camera's axis-vectors
Private Sub Maintenance()
    On Error GoTo error_h
    
    D3DXVec3Normalize vtrLook, vtrLook      'By normalizing a vector, it's length is set to 1
    D3DXVec3Cross vtrRight, vtrUp, vtrLook  'A cross product vector will be perpendicular to the other two
    D3DXVec3Normalize vtrRight, vtrRight    'but it's length will be a combination of both so
    D3DXVec3Cross vtrUp, vtrLook, vtrRight  'each vector must be re-normalized to ensure that all
    D3DXVec3Normalize vtrUp, vtrUp          'of the vectors are perpendicular to each other as they
                                            'should be.
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsCamera.Maintenance")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'sngSpeed: The amount to move forward (use a negative value to move backward)
Public Sub MoveForward(sngSpeed As Single)
    On Error GoTo error_h:
    
    'vtrLook (Z axis) should always point forward.  If this vector is normalized, we can multiply each component
    'by the sngSpeed argument to get the distance traveled along each axis
    With vtrPosition
        .X = .X + sngSpeed * vtrLook.X
        .Y = .Y + sngSpeed * vtrLook.Y
        .z = .z + sngSpeed * vtrLook.z
    End With
    
    'This updates the camera position
    Update
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsCamera.MoveForward(" & sngSpeed & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'Just a simple property get in case you need to know where the camera is
Property Get position() As D3DVECTOR
    position = vtrPosition
End Property
'This will set a new position for the camera
Property Let position(newVector As D3DVECTOR)
    vtrPosition = newVector
    Update
End Property

'This sub lets you 'spin' the camera in place to the left/right
'sngAngle:  The angle, in radians, to rotate the camera (use a negative value to spin left)
Public Sub RotateY(sngAngle As Single)
    On Error GoTo error_h
    
    'The rotation matrix to be applied to the other vectors
    Dim matRotation As D3DMATRIX
    
    'Creates a rotation matrix based on whatever vector is currently 'up' at the moment
    D3DXMatrixRotationAxis matRotation, vtrUp, sngAngle
    
    'Apply the rotation matrix to the other vectors so that they will still 'point' forward and right
    D3DXVec3TransformCoord vtrLook, vtrLook, matRotation
    D3DXVec3TransformCoord vtrRight, vtrRight, matRotation
    
    'Correct any floating point math errors
    Maintenance
    
    'Set the new camera position/orientation
    Update
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsCamera.RotateY(" & sngAngle & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub
'Same as the RotateY sub, but will 'roll' the camera
Public Sub RotateZ(sngAngle As Single)
    On Error GoTo error_h
    
    Dim matRotation As D3DMATRIX
    
    D3DXMatrixRotationAxis matRotation, vtrLook, sngAngle
    
    D3DXVec3TransformCoord vtrUp, vtrUp, matRotation
    D3DXVec3TransformCoord vtrRight, vtrRight, matRotation
    
    Maintenance
    
    Update
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsCamera.RotateZ")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'Same as RotateY/RotateZ but will angle the camera up/down
Public Sub RotateX(sngAngle As Single)
    On Error GoTo error_h
    
    Dim matRotation As D3DMATRIX
    
    D3DXMatrixRotationAxis matRotation, vtrRight, sngAngle
    
    D3DXVec3TransformCoord vtrUp, vtrUp, matRotation
    D3DXVec3TransformCoord vtrLook, vtrLook, matRotation
    
    Maintenance
    
    Update
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsCamera.RotateX(" & sngAngle & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub


'Store the D3DDevice that jDXEngine is using
'newD3DDevice:  a valid Direct3DDevice8 object created by jDXEngine
Public Sub SetDevice(newD3DDevice As Direct3DDevice8)
    Set D3DDevice = newD3DDevice
End Sub

'This is to give the camera initial values
'LOOK:      A vector describing which direction to look in
'UP:        A vector describing which direction is 'up'
'position:  A vector containing the camera's coordinates in world space
'Please note that there is no RIGHT vector at this time.  That vector will be created in the
'maintenance() function as the DotProduct of LOOK and UP
Public Sub SetupCamera(LOOK As D3DVECTOR, UP As D3DVECTOR, position As D3DVECTOR)
    On Error GoTo error_h
    
    vtrLook = LOOK          'Set all 3 member variables
    vtrUp = UP
    vtrPosition = position
    
    Maintenance             'Fix all directional vectors
    
    Update                  'Set the new data
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "SetupCamera2")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'I used this sub for debugging purposes to find the camera's position in World Space
Public Sub DisplayVector(v As D3DVECTOR)
    Debug.Print v.X & "," & v.Y & "," & v.z
End Sub

'Same as the MoveForward sub, except that we are moving along the RIGHT vector instead of the LOOK vector
Public Sub StrafeRight(sngSpeed As Single)
    On Error GoTo error_h
    
    With vtrPosition
        .X = .X + sngSpeed * vtrRight.X
        .Y = .Y + sngSpeed * vtrRight.Y
        .z = .z + sngSpeed * vtrRight.z
    End With
    
    'Since no vectors were rotated, no maintenance need be performed
    
    Update
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsCamera.StrafeRight(" & sngSpeed & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'This sub will create a ViewMatrix out of the 3 directional vectors of the camera
'I don't particularly understand the 3DMath behind this operation, but it has proven
'itself to work.
Private Sub Update()
    On Error GoTo error_h
    
    'Make sure that a valid device exists as to prevent errors
    If D3DDevice Is Nothing Then Exit Sub
    
    'The matrix to be created
    Dim matView As D3DMATRIX
    
    'Reset the matrix to default values
    D3DXMatrixIdentity matView
    
    'Put the values into the matrix.  I got the procedure for this from a tutorial on the web, but
    'I don't remember which one.  From what I understand, this is basic matrix mathmatics and
    'I'd be very grateful if someone could point me to a tutorial/faq on matrix math. :)
    With matView
        .m11 = vtrRight.X
        .m21 = vtrRight.Y
        .m31 = vtrRight.z
        .m41 = -D3DXVec3Dot(vtrPosition, vtrRight)
        .m12 = vtrUp.X
        .m22 = vtrUp.Y
        .m32 = vtrUp.z
        .m42 = -D3DXVec3Dot(vtrPosition, vtrUp)
        .m13 = vtrLook.X
        .m23 = vtrLook.Y
        .m33 = vtrLook.z
        .m43 = -D3DXVec3Dot(vtrPosition, vtrLook)
        .m14 = 0
        .m24 = 0
        .m34 = 0
        .m44 = 1
    End With

    'Now, send the camera data to the D3DDevice
    D3DDevice.SetTransform D3DTS_VIEW, matView
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "clsCamera.Update")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub
