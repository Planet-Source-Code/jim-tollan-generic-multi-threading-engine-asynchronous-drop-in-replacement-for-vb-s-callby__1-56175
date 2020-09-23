VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICallBackInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ************************************************************************
' Copyright:    All rights reserved.  © 2004
' Project:      AsyncServer
' Module:       ICallBackInterface
' Author:       james b tollan
' Date:         15-Sep-2004 14:34
' Revisions **************************************************************
' Date          Name                Description
' 16-Sep-2004                       Changed FactoryObject to
'                                   PreProcessFactoryObject
' ************************************************************************

Option Explicit

Public Sub ErrorMessage _
    (ByVal ErrorMessage As String, _
    ErrorNumber As Long)
End Sub

Public Sub Initialized _
    (ByVal ProgramID As String, _
    ByVal Status As Boolean)
End Sub

Public Sub Finished _
    (ByVal ProgramID As String, _
    ByVal Status As Boolean)
End Sub

Public Sub PreProcessFactoryObject(ByRef RequestedObject As Object)
End Sub
