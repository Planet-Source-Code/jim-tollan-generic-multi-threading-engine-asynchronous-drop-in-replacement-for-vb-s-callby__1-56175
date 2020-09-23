VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsyncClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ************************************************************************
' Copyright:    All rights reserved.  © 2004
' Project:      AsyncServer
' Module:       AsyncClass
' Author:       james b tollan
' Date:         15-Sep-2004 14:34
' Revisions **************************************************************
' Date          Name                Description
' 16-Sep-2004                       Changed FactoryObject to
'                                   PreProcessFactoryObject
' ************************************************************************

Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Purpose:
'
'   To instanciate any COM object Out-Of-Process. Single
'   Interface - ReturnAsyncObject - uses five
'   parameters which capitalize on CreateObject and
'   CallByName VB functions to return the created object
'   to the calling application - "Caller" - and then act on the object
'   with the DLL entry-point and args the user specified.
'
'       ClassName - ClassName param in CreateObject function
'       ServerName - ServerName param in CreateObject function
'                       (may be ""/vbNullString for local instance)
'       ProcName - Used as IDispatch param in CallByName function
'       CallType - Used to identify Interface type (Method, or
'                       Property Let, Get, Set) for CallByName function
'       Arguments() - Interface arguments matching ProcName IDispatch Arguments
'
'   Arguments are used in conjunction with modular variables to
'   avoid further over-head since this obj is intended to be implemented
'   as a SingleUse/"Thread Per Object" COM exe server.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' enums
Enum cbnCallTypes
    VbLet = DISPATCH_PROPERTYPUT
    VbGet = DISPATCH_PROPERTYGET
    VbSet = DISPATCH_PROPERTYPUTREF
    VbMethod = DISPATCH_METHOD
End Enum

' events
Public Event ErrorMessage _
    (ByVal ErrorMessage As String, _
    ErrorNumber As Long)

Public Event Initialized _
    (ByVal ProgramID As String, _
    ByVal Status As Boolean)

Public Event Finished _
    (ByVal ProgramID As String, _
    ByVal Status As Boolean)

Public Event PreProcessFactoryObject(ByRef RequestedObject As Object)

Public Property Set ClientInterface(ByRef pClientInterface As ICallBackInterface)
    Set gICallBackInterface = pClientInterface
End Property
Friend Sub ErrorMessage(ByVal ErrorMessage As String, ErrorNumber As Long)
    If Not gICallBackInterface Is Nothing Then
        Call gICallBackInterface.ErrorMessage(ErrorMessage, ErrorNumber)
    Else
        RaiseEvent ErrorMessage(ErrorMessage, ErrorNumber)
    End If
End Sub
Friend Sub Finished()
    If Not gICallBackInterface Is Nothing Then
        Call gICallBackInterface.Finished(GetProgramID(gcClass), Not (gcClass Is Nothing))
    Else
        RaiseEvent Finished(GetProgramID(gcClass), Not (gcClass Is Nothing))
    End If
    Set gcClass = Nothing
    Set gICallBackInterface = Nothing
End Sub

Public Function ReturnAsyncObject _
    (ByRef pobjSource As Object, _
    ByVal ProcName As String, _
    ByVal CallType As cbnCallTypes, _
    ParamArray Arguments()) As Object
    
    Dim strProgID As String

    On Error GoTo ErrOut

    strProgID = GetProgramID(pobjSource)
    Set gCallbackObj = Me

    gstrProcName = ProcName
    gCallType = CallType
    gArgs = Arguments
    
    Set gcClass = CreateObject(strProgID)

    If Not gICallBackInterface Is Nothing Then
        gICallBackInterface.PreProcessFactoryObject gcClass
    Else
        RaiseEvent PreProcessFactoryObject(gcClass)
    End If

    Set ReturnAsyncObject = gcClass

    If Not gICallBackInterface Is Nothing Then
        Call gICallBackInterface.Initialized(strProgID, Not (gcClass Is Nothing))
    Else
        RaiseEvent Initialized(strProgID, Not (gcClass Is Nothing))
    End If

    If Err.Number = 0 Then
        glngTimerID = SetTimer(0, 0, 1, AddressOf Instanciate)
    End If

    ' not happy with this but keeps multiple calls
    ' in check a bit better (suggestions please)!!
    DoEvents

    Exit Function

ErrOut:
    If Not gICallBackInterface Is Nothing Then
        Call gICallBackInterface.ErrorMessage("[ReturnAsyncObject: " _
            & strProgID & " ] " & vbCrLf _
            & Err.Description, Err.Number)
    Else
        RaiseEvent ErrorMessage("[ReturnAsyncObject: " _
            & strProgID & " ] " & vbCrLf _
            & Err.Description, Err.Number)
    End If
    Set gCallbackObj = Nothing
    Set gcClass = Nothing
    Set gICallBackInterface = Nothing

End Function

Private Sub Class_Terminate()
    Set gCallbackObj = Nothing
    Set gcClass = Nothing
    Set gICallBackInterface = Nothing
    Erase gArgs
    DoEvents
End Sub

Private Function GetProgramID(pOBject As Object) As String
    Dim oTypeLib As TLIApplication

    GetProgramID = vbNullString

    On Error Resume Next
    
    Set oTypeLib = New TLIApplication
    With oTypeLib
        ' Hope This object has a parent.......
        If Not .InterfaceInfoFromObject(pOBject).Parent Is Nothing Then
            If Err = 0 Then
                GetProgramID = .InterfaceInfoFromObject(pOBject).Parent _
                & "." & TypeName(pOBject)
            Else
                GetProgramID = TypeName(pOBject)
            End If
        End If
        Print GetProgramID
    End With

    Set oTypeLib = Nothing
    Err.Clear
End Function
