VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Async File Save Test"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8205
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
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
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Asyncronous Thread Implements"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Log Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":030A
      Top             =   1080
      Width           =   7875
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Current Thread"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2220
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Asyncronous Thread Events"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6180
      TabIndex        =   0
      Top             =   3000
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   600
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CAPTION_TEXT = "Async File Save Test - "

Const FILE_1_MAX = 33400
Const FILE_2_MAX = 66810
Const MOD_TEST = 25

Const PATH_1 = "c:\temp\bigfile1.txt"
Const PATH_2 = "c:\temp\bigfile2.txt"

Private WithEvents mAsyncFile       As AsyncTest.CFileTest
Attribute mAsyncFile.VB_VarHelpID = -1
Private WithEvents mAsyncFile2      As AsyncTest.CFileTest
Attribute mAsyncFile2.VB_VarHelpID = -1
Private WithEvents mCAsyncServer    As AsyncServer.AsyncClass
Attribute mCAsyncServer.VB_VarHelpID = -1
Private mCAsyncServerImplements     As AsyncServer.AsyncClass

Implements ICallBackInterface
Private mblnCancel As Boolean

Private Sub cmdCancel_Click()
    mblnCancel = True
    Text1 = "Cancelled..." & vbCrLf
End Sub

Private Sub Command1_Click()
    ' the command1_click proc
    ' purely instanciates the objects
    ' and kicks off the required proc

    Set mCAsyncServer = New AsyncServer.AsyncClass
    ' it is important to set mAsyncFile to nothing
    ' in order to keep the ole object reference count
    ' in sync. this is by design and is true of many ole
    ' servers (Winword and Excel being good examples
    mblnCancel = False
    Set mAsyncFile = Nothing
    ' notice we HAVE to use the New declaration
    ' purely just to identify the GUID to the typlib
    ' in asyncserver. this means we don't get caught out
    ' with typo's now.
    Set mAsyncFile = mCAsyncServer.ReturnAsyncObject _
        (New AsyncTest.CFileTest, "DoBigJob", VbMethod, MOD_TEST, FILE_1_MAX)
End Sub

Private Sub Command2_Click()
    ' this test is run from the current exe thread
    cmdCancel.Enabled = False
    Caption = CAPTION_TEXT & "running..."
    With pb2
        .Min = 0
        .Max = FILE_2_MAX
        .Value = 1
    End With
    mblnCancel = False
    Set mAsyncFile2 = New CFileTest
    mAsyncFile2.FileName = PATH_2
    mAsyncFile2.DoBigJob MOD_TEST, FILE_2_MAX
    ' notice how the syncronous execution follows
    ' thro in a single procedure and we get our
    ' updated caption etc in the command2_click proc
    pb2.Value = 0
    Caption = CAPTION_TEXT & "mAsyncFile2 (finished 2) :" & Not mAsyncFile2 Is Nothing
    cmdCancel.Enabled = True
End Sub

Private Sub Command3_Click()
    Text1 = ""
End Sub

Private Sub Command4_Click()
    ' the command1_click proc
    ' purely instanciates the objects
    ' and kicks off the required proc
    Set mCAsyncServerImplements = New AsyncServer.AsyncClass
    ' it is important to set mAsyncFile to nothing
    ' in order to keep the ole object reference count
    ' in sync. this is by design and is true of many ole
    ' servers (Winword and Excel being good examples
    Set mAsyncFile2 = Nothing
    ' this is the only difference here
    ' however, look at how the rotine continues
    ' when the msgbox is showing
    mblnCancel = False
    ' set our callback interface to receive notification
    Set mCAsyncServerImplements.ClientInterface = Me
    ' notice we HAVE to use the New declaration
    ' purely just to identify the GUID to the typlib
    ' in asyncserver. this means we don't get caught out
    ' with typo's now.
    Set mAsyncFile2 = mCAsyncServerImplements.ReturnAsyncObject _
        (New AsyncTest.CFileTest, "DoBigJob", VbMethod, MOD_TEST, FILE_1_MAX)

End Sub

Private Sub ICallBackInterface_ErrorMessage(ByVal ErrorMessage As String, ErrorNumber As Long)
'
End Sub

Private Sub ICallBackInterface_PreProcessFactoryObject(RequestedObject As Object)
    RequestedObject.FileName = PATH_2
End Sub

Private Sub ICallBackInterface_Finished(ByVal ClassName As String, ByVal Status As Boolean)
    Caption = CAPTION_TEXT & ClassName & " (finished interface) :" & Status
    If Not mAsyncFile Is Nothing Then
        Text1 = Text1 & mAsyncFile.Result & vbCrLf
    End If
    pb2.Value = 0
End Sub

Private Sub ICallBackInterface_Initialized(ByVal ProgramID As String, ByVal Status As Boolean)
    Caption = CAPTION_TEXT & "running... interface"
    With pb2
        .Min = 0
        .Max = FILE_1_MAX
        .Value = 1
    End With
    Text1 = Text1 & ProgramID & ":" & Status & vbCrLf
End Sub

Private Sub mAsyncFile_CurrentWrite(ByVal LineNumber As Long, ByRef CancelJob As Boolean)
    pb1.Value = LineNumber
    CancelJob = mblnCancel
End Sub

Private Sub mAsyncFile_Message(ByVal Message As String)
    Text1 = Text1 & Message & vbCrLf
End Sub

Private Sub mAsyncFile_SystemActive()
    Text1 = Text1 & Now & vbCrLf
End Sub

Private Sub mAsyncFile2_CurrentWrite(ByVal LineNumber As Long, ByRef CancelJob As Boolean)
    pb2.Value = LineNumber
    CancelJob = mblnCancel
End Sub

Private Sub mAsyncFile2_Message(ByVal Message As String)
    Text1 = Text1 & Message & vbCrLf
End Sub

Private Sub mAsyncFile2_SystemActive()
    Text1 = Text1 & Now & vbCrLf
End Sub

Private Sub mCAsyncServer_ErrorMessage(ByVal ErrorMessage As String, ErrorNumber As Long)
    MsgBox ErrorMessage
End Sub

Private Sub mCAsyncServer_PreProcessFactoryObject(RequestedObject As Object)
    Dim CTest As AsyncTest.CFileTest
    Set CTest = New AsyncTest.CFileTest
    If TypeName(RequestedObject) = TypeName(CTest) Then
        With RequestedObject
            .FileName = PATH_1
        End With
    End If
End Sub

Private Sub mCAsyncServer_Initialized(ByVal ProgramID As String, ByVal Status As Boolean)
    ' this test is run from an asyncronous thread
    Caption = CAPTION_TEXT & "running...event"
    With pb1
        .Min = 0
        .Max = FILE_1_MAX
        .Value = 1
    End With
    Text1 = Text1 & ProgramID & ":" & Status & vbCrLf
End Sub
Private Sub mCAsyncServer_Finished(ByVal ClassName As String, ByVal Status As Boolean)
    Caption = CAPTION_TEXT & ClassName & " (finished event) :" & Status
    If Not mAsyncFile Is Nothing Then
        Text1 = Text1 & mAsyncFile.Result & vbCrLf
    End If
    pb1.Value = 0
End Sub
