VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlTest 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   ScaleHeight     =   405
   ScaleWidth      =   8985
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8100
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H0080FF80&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "ctlTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
' Copyright:    All rights reserved.  © 2004
' Project:      Project1
' Module:       ctlTest
' Author:       james b tollan
' Date:         15-Sep-2004 15:39
' Revisions **************************************************************
' Date          Name                Description
' ************************************************************************

Option Explicit

' constants
Const CAPTION_TEXT = "Async File Save Test - "
Const FILE_1_MAX = 33400
Const MOD_TEST = 25

' events
Event Caption(ByVal NewCaption As String)
Event ErrorMessage(ByVal Message As String, ByVal ErrorNumber As Long)
Event Change(ByVal NewText As String)

' module level variables
' (note mAsyncFile and mCAsyncServer are dimmed WithEvents)
Private WithEvents mAsyncFile       As Asynctest.CFileTest
Attribute mAsyncFile.VB_VarHelpID = -1
Private WithEvents mCAsyncServer    As AsyncServer.AsyncClass
Attribute mCAsyncServer.VB_VarHelpID = -1
Private mblnUseCallBackInterface    As Boolean
Private mblnUseExeThread            As Boolean
Private mstrFile                    As String
Private mblnCancel                  As Boolean

' implemented interfaces
Implements ICallBackInterface
Implements ICFileTest

' write only properties
Public Property Let UseExeThread(ByVal pblnUseExeThread As Boolean)
    mblnUseExeThread = pblnUseExeThread
End Property
Public Property Let UseCallBackInterface(ByVal pblnUseCallBackInterface As Boolean)
    mblnUseCallBackInterface = pblnUseCallBackInterface
End Property
Public Property Let Filename(ByVal pstrFile As String)
    mstrFile = pstrFile
End Property

Private Sub cmdCancel_Click()
    mblnCancel = True
    RaiseEvent Change("Cancelled..." & vbCrLf)
End Sub
Private Sub cmdGo_Click()
    mblnCancel = False
    If mblnUseExeThread Then
        Set mAsyncFile = New Asynctest.CFileTest
        With pb1
            .Min = 0
            .Max = FILE_1_MAX
            .Value = 1
        End With
        RaiseEvent Caption(CAPTION_TEXT & "running... " & mstrFile)
        With mAsyncFile
            .Filename = mstrFile
            .DoBigJob MOD_TEST, FILE_1_MAX
        End With
        pb1.Value = 0
    Else
        ' this proc
        ' purely instanciates the objects
        ' and kicks off the required proc
        Set mCAsyncServer = New AsyncServer.AsyncClass
        ' set our ICallBackInterface to the asyncserver
        If mblnUseCallBackInterface Then
            Set mCAsyncServer.ClientInterface = Me
        End If
        ' it is important to set mAsyncFile to nothing
        ' in order to keep the ole object reference count
        ' in sync. this is by design and is true of many ole
        ' servers (Winword and Excel being good examples)
        Set mAsyncFile = Nothing
        ' notice we HAVE to use the New declaration
        ' purely just to identify the GUID to the typlib
        ' in asyncserver. this means we don't get caught out
        ' with typo's now.
        Set mAsyncFile = mCAsyncServer.ReturnAsyncObject _
            (New Asynctest.CFileTest, "DoBigJob", VbMethod, _
            MOD_TEST, FILE_1_MAX)
    End If
End Sub
Private Sub ICFileTest_CurrentWrite(ByVal LineNumber As Long, CancelJob As Boolean)
    pb1.Value = LineNumber
    CancelJob = mblnCancel
End Sub
Private Sub ICFileTest_Message(ByVal Message As String)
    RaiseEvent Change(Message & vbCrLf)
End Sub
Private Sub ICFileTest_SystemActive()
    RaiseEvent Change(Now & vbCrLf)
End Sub
Private Sub mAsyncFile_CurrentWrite(ByVal LineNumber As Long, CancelJob As Boolean)
    pb1.Value = LineNumber
    CancelJob = mblnCancel
End Sub
Private Sub mAsyncFile_Message(ByVal Message As String)
    RaiseEvent Change(Message & vbCrLf)
End Sub
Private Sub mAsyncFile_SystemActive()
    RaiseEvent Change(Now & vbCrLf)
End Sub

Private Sub ICallBackInterface_ErrorMessage(ByVal ErrorMessage As String, ErrorNumber As Long)
    RaiseEvent ErrorMessage(ErrorMessage, ErrorNumber)
End Sub
Private Sub ICallBackInterface_PreProcessFactoryObject(RequestedObject As Object)
    ' this event has the ability to HOLD up
    ' excecution in the AsyncServer. we can use
    ' this for showing a modal interface and/or
    ' populating further properties etc..
    With RequestedObject
        ' just setting this property to
        ' show that we can set our black-box
        ' object's properties after it's been
        ' created but before it get's sent async
        .Filename = mstrFile
        ' set the CFileTest interface (RequestedObject in this case)
        ' to our implmented interface (ICFileTest)
        Set .CallbackInterface = Me
    End With
End Sub
Private Sub ICallBackInterface_Finished(ByVal ProgramID As String, ByVal Status As Boolean)
    RaiseEvent Caption(CAPTION_TEXT & ProgramID & " (finished " & mstrFile & ") :" & Status)
    If Not mAsyncFile Is Nothing Then
        RaiseEvent Change(mAsyncFile.Result & vbCrLf)
    End If
    pb1.Value = 0
End Sub
Private Sub ICallBackInterface_Initialized(ByVal ProgramID As String, ByVal Status As Boolean)
    RaiseEvent Caption(CAPTION_TEXT & "running... " & mstrFile)
    With pb1
        .Min = 0
        .Max = FILE_1_MAX
        .Value = 1
    End With
    RaiseEvent Change(ProgramID & ":" & Status & vbCrLf)
End Sub
Private Sub mCAsyncServer_ErrorMessage(ByVal ErrorMessage As String, ErrorNumber As Long)
    RaiseEvent ErrorMessage(ErrorMessage, ErrorNumber)
End Sub
Private Sub mCAsyncServer_PreProcessFactoryObject(RequestedObject As Object)
    With RequestedObject
        ' just setting this property to
        ' show that we can set our black-box
        ' object's properties after it's been
        ' created but before it get's sent async
        .Filename = mstrFile
    End With
End Sub
Private Sub mCAsyncServer_Finished(ByVal ClassName As String, ByVal Status As Boolean)
    RaiseEvent Caption(CAPTION_TEXT & ClassName & " (finished " & mstrFile & ") :" & Status)
    If Not mAsyncFile Is Nothing Then
        RaiseEvent Change(mAsyncFile.Result & vbCrLf)
    End If
    pb1.Value = 0
End Sub
Private Sub mCAsyncServer_Initialized(ByVal ProgramID As String, ByVal Status As Boolean)
    RaiseEvent Caption(CAPTION_TEXT & "running... " & mstrFile)
    With pb1
        .Min = 0
        .Max = FILE_1_MAX
        .Value = 1
    End With
    RaiseEvent Change(ProgramID & ":" & Status & vbCrLf)
End Sub
Private Sub UserControl_Terminate()
    mblnCancel = True
End Sub
