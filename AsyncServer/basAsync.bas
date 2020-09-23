Attribute VB_Name = "basAsync"
Option Explicit


Private Declare Function CoLockObjectExternal Lib "ole32" (ByVal pUnk As IUnknown, _
    ByVal fLock As Long, _
    ByVal fLastUnlockReleases As Long) As Long

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
    ByVal nIDEvent As Long) As Long

Private msProcName  As String
Private mCallType   As cbnCallTypes
Private mArgs()     As Variant
Private mCol        As Collection

Public Function InitializeClass(ByVal Class As String, _
    ByVal Server As String, _
    ByVal ProcName As String, _
    ByVal CallType As cbnCallTypes, _
    ParamArray Arguments()) As Object
    
    Dim CTargetClass As Object

    If mCol Is Nothing Then
        Set mCol = New Collection
    End If
    
    If Len(Server) Then
        Set CTargetClass = CreateObject(Server & "." & Class)
    Else
        Set CTargetClass = CreateObject(Class)
    End If

    msProcName = ProcName
    mCallType = CallType
    mArgs = Arguments(0)

    If Err.Number = 0 Then                            'object successfully created
        Set InitializeClass = CTargetClass            'return object reference
        'defer continued execution of this class to windows so the created object is
        'returned to the caller immediately...timer is set to 0 so the caller receives
        'the object reference and the Instanciate function will run at the "same time"
        'this is where the async server forces the caller and object operations to
        'multi-process with complete independence of eachother
        mCol.Add CTargetClass, CStr(SetTimer(0, 0, 1, AddressOf Instanciate))
        Set CTargetClass = Nothing
    End If
End Function

Private Sub Instanciate(ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal TimerID As Long, _
    ByVal TimeIntVal As Long)

    On Error GoTo ErrMsg
    
    Call CoLockObjectExternal(mCol(CStr(TimerID)), 1, 1)
    'lock - object reference has been returned to the caller...
    KillTimer 0, TimerID
    
    'lock prevents caller from using it until
    'the CallByNameFixParamArray function has invoked the caller's first desired action
    CallByNameFixParamArray mCol(CStr(TimerID)), msProcName, mCallType, mArgs
    'CallByName in vb6 doesn't treat params correctly
    Call CoLockObjectExternal(mCol(CStr(TimerID)), 0, 1)
    'unlock - release system's lock on class object
    'so caller has control from here on in
    'variable cleanup
    'if the Arguments() paramarray contained an object
    'the Erase function will set it to nothing
    
    mCol.Remove CStr(TimerID)
    Erase mArgs
    Exit Sub

ErrMsg:
    Err.Raise Err.Number, "basAsync.Instantiate", Err.Description
End Sub
