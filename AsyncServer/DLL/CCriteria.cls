VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "CreatedWithClassBuilder" ,"Yes"
Option Explicit


'************************************************************
'* Class Name:      CCriteria.cls
'* Copyright:       itv © 2004
'* Created:         15 Sep 2004 09:25
'* Authored by:     james tollan
'* Description:     <--here-->
'* Last Change:     15 Sep 2004
'************************************************************

' Private variables used internally by CCriteria

Private mstrFileName    As String
Private mlngIterations  As Long
Private mlngModCheck    As Long
Private mstrLineTag     As String

Public Property Let FileName(ByVal pstrFileName As String)
    mstrFileName = pstrFileName
End Property
Public Property Get FileName() As String
    FileName = mstrFileName
End Property

Public Property Let Iterations(ByVal plngIterations As Long)
    mlngIterations = plngIterations
End Property
Public Property Get Iterations() As Long
    Iterations = mlngIterations
End Property

Public Property Let ModCheck(ByVal plngModCheck As Long)
    mlngModCheck = plngModCheck
End Property
Public Property Get ModCheck() As Long
    ModCheck = mlngModCheck
End Property

Public Property Let LineTag(ByVal pstrLineTag As String)
    mstrLineTag = pstrLineTag
End Property
Public Property Get LineTag() As String
    LineTag = mstrLineTag
End Property
Friend Sub Clone(pCCriteria As CCriteria)
    
    With pCCriteria
        FileName = .FileName
        Iterations = .Iterations
        LineTag = .LineTag
        ModCheck = .ModCheck
    End With

End Sub
