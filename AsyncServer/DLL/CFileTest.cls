VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "CreatedWithClassBuilder" ,"Yes"
Option Explicit

Private Declare Function GetQueueStatus Lib "user32" _
    (ByVal fuFlags As Long) As Long

Const QS_KEY = &H1
Const QS_HOTKEY = &H80
Const QS_MOUSEMOVE = &H2
Const QS_MOUSEBUTTON = &H4
Const QS_POSTMESSAGE = &H8
Const QS_TIMER = &H10
Const QS_PAINT = &H20
Const QS_SENDMESSAGE = &H40
Const QS_UI_INPUT = (QS_MOUSEBUTTON Or QS_KEY Or QS_SENDMESSAGE)
Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or _
    QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Event Message(ByVal Message As String)
Event CurrentWrite(ByVal LineNumber As Long, ByRef CancelJob As Boolean)
Event SystemActive()

Private mCICallback     As ICFileTest
Private mstrResult      As String
Private mstrFilename    As String
Private mlngIterations  As Long
Private mlngModBreak    As Long

Public Property Set CallbackInterface(ByRef pICFileTest As ICFileTest)
    Set mCICallback = pICFileTest
End Property
Public Property Let Iterations(ByVal plngIterations As Long)
    mlngIterations = plngIterations
End Property
Public Property Get Iterations() As Long
    Iterations = mlngIterations
End Property

Public Property Let ModBreak(ByVal plngModBreak As Long)
    mlngModBreak = plngModBreak
End Property
Public Property Get ModBreak() As Long
    ModBreak = mlngModBreak
End Property

Public Property Let Filename(ByVal pstrFilename As String)
    mstrFilename = pstrFilename
End Property
Private Property Get Filename() As String
    Filename = mstrFilename
End Property

Public Function DoBigJobNoParams()
    ' declare vars
    Dim lngLoop     As Long
    Dim T           As Double
    Dim blnCancel   As Boolean

    On Error Resume Next
    
    If Filename = vbNullString Then Exit Function
    T = Timer

    For lngLoop = 1 To Iterations
        Open Filename For Append As #1
        Write #1, Now
        Close #1
'        CheckEvents
        If lngLoop Mod ModBreak = 0 Or lngLoop = Iterations Then
            If Not mCICallback Is Nothing Then
                mCICallback.CurrentWrite lngLoop, blnCancel
            Else
                RaiseEvent CurrentWrite(lngLoop, blnCancel)
            End If
        End If
        If blnCancel Then Exit For
    Next

    mstrResult = FileLen(Filename)
    If Not mCICallback Is Nothing Then
        mCICallback.Message "Done and took: " & Timer - T & " [new iterated " & lngLoop & " times]"
    Else
        RaiseEvent Message("Done and took: " & Timer - T & " [new iterated " & lngLoop & " times]")
    End If

End Function
Public Sub DoBigJob _
    (ByVal plngMod As Long, _
    ByVal plngIterations As Long)

    ' declare vars
    Dim lngLoop     As Long
    Dim T           As Double
    Dim blnCancel   As Boolean

    On Error Resume Next
    
    If Filename = vbNullString Then Exit Sub
    T = Timer

    For lngLoop = 1 To plngIterations
        Open Filename For Append As #1
        Write #1, Now
        Close #1
'        CheckEvents
        If lngLoop Mod plngMod = 0 Or lngLoop = plngIterations Then
            If Not mCICallback Is Nothing Then
                mCICallback.CurrentWrite lngLoop, blnCancel
            Else
                RaiseEvent CurrentWrite(lngLoop, blnCancel)
            End If
        End If
        If blnCancel Then Exit For
    Next

    mstrResult = FileLen(Filename)
    If Not mCICallback Is Nothing Then
        mCICallback.Message "Done and took: " & Timer - T & " [new iterated " & lngLoop & " times]"
    Else
        RaiseEvent Message("Done and took: " & Timer - T & " [new iterated " & lngLoop & " times]")
    End If

End Sub
Public Property Get Result() As String
    Result = mstrResult
End Property

Private Sub CheckEvents()
    ' GetInputState returns 1 if a button has been clicked or a key
    ' has been pressed.  Allow a DoEvents only if the GetInputState
    ' is not 0.
    If GetQueueStatus(QS_UI_INPUT) Then
        If Not mCICallback Is Nothing Then
            mCICallback.SystemActive
        Else
            RaiseEvent SystemActive
        End If
        DoEvents
    End If
End Sub
