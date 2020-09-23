Attribute VB_Name = "basAsync"
' ************************************************************************
' Copyright:    All rights reserved.  © 2004
' Project:      AsyncServer
' Module:       basAsync
' Author:       james b tollan
' Date:         15-Sep-2004 14:33
' Revisions **************************************************************
' Date          Name                Description
' ************************************************************************

Option Explicit

Private Declare Function CoLockObjectExternal Lib "ole32" _
    (ByVal pUnk As IUnknown, _
    ByVal fLock As Long, _
    ByVal fLastUnlockReleases As Long) As Long

Public Declare Function SetTimer Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal nIDEvent As Long) As Long

Public gICallBackInterface  As ICallBackInterface
Public gstrProcName         As String
Public gCallType            As cbnCallTypes
Public gArgs()              As Variant
Public gcClass              As Object
Public glngTimerID          As Long
Public gCallbackObj         As AsyncClass


Public Sub Instanciate _
    (ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal TimerID As Long, _
    ByVal TimeIntVal As Long)

    Dim lngLoop As Long
    Dim lngMax  As Long

    On Error GoTo ErrMsg

    Call CoLockObjectExternal(gcClass, 1, 1)    'lock - object reference has been returned to the caller...
    KillTimer 0, glngTimerID                    'lock prevents caller from using it until
                                                'the CallByNameFixParamArray function
                                                'has invoked the caller's first desired action
    CallByNameFixParamArray gcClass, gstrProcName, gCallType, gArgs
    'CallByName in vb6 doesn't treat params correctly
    Call CoLockObjectExternal(gcClass, 0, 1)
    'unlock - release system's lock on class object
    'so caller has control from here on in

    lngMax = UBound(gArgs)

    For lngLoop = 0 To lngMax
        If VarType(gArgs(lngMax - lngLoop)) = vbObject Then
            Set gArgs(lngMax - lngLoop) = Nothing
        End If
    Next

    Erase gArgs

    If Not gCallbackObj Is Nothing And Not gcClass Is Nothing Then
        gCallbackObj.Finished
        Set gCallbackObj = Nothing
    End If

    Exit Sub

ErrMsg:
    gCallbackObj.ErrorMessage "[" & TypeName(gcClass) & "." & gstrProcName & "] " _
        & vbCrLf & Err.Description, Err.Number
    Set gCallbackObj = Nothing
    Set gcClass = Nothing
    Erase gArgs
 End Sub
