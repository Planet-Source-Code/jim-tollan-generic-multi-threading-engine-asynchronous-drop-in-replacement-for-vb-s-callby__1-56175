VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Async File Save Test"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctlTest ctlTest1 
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   661
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
      Left            =   8160
      TabIndex        =   1
      Top             =   3780
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2115
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   8955
   End
   Begin Project1.ctlTest ctlTest2 
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   600
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   661
   End
   Begin Project1.ctlTest ctlTest3 
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   1080
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   661
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' Copyright:    All rights reserved.  © 2004
' Project:      Project1
' Module:       Form1
' Author:       james b tollan
' Date:         15-Sep-2004 15:38
' Revisions **************************************************************
' Date          Name                Description
' ************************************************************************

Option Explicit


Private Sub Command3_Click()
    Text1 = vbNullString
End Sub
Private Sub ctlTest1_Caption(ByVal NewCaption As String)
    Caption = NewCaption
End Sub
Private Sub ctlTest1_Change(ByVal NewText As String)
    Text1 = Text1 & NewText
End Sub
Private Sub ctlTest1_ErrorMessage(ByVal Message As String, ByVal ErrorNumber As Long)
    MsgBox Message
End Sub
Private Sub ctlTest2_Caption(ByVal NewCaption As String)
    Caption = NewCaption
End Sub
Private Sub ctlTest2_Change(ByVal NewText As String)
    Text1 = Text1 & NewText
End Sub
Private Sub ctlTest2_ErrorMessage(ByVal Message As String, ByVal ErrorNumber As Long)
    MsgBox Message
End Sub
Private Sub ctlTest3_Caption(ByVal NewCaption As String)
    Caption = NewCaption
End Sub
Private Sub ctlTest3_Change(ByVal NewText As String)
    Text1 = Text1 & NewText
End Sub
Private Sub ctlTest3_ErrorMessage(ByVal Message As String, ByVal ErrorNumber As Long)
    MsgBox Message
End Sub
Private Sub Form_Load()
    With ctlTest1
        .Filename = "c:\temp\bigfile1.txt"
        .UseCallBackInterface = True
    End With
    With ctlTest2
        .Filename = "c:\temp\bigfile2.txt"
        .UseCallBackInterface = True
    End With
    With ctlTest3
        .Filename = "c:\temp\bigfile3.txt"
        .UseExeThread = True
    End With
End Sub
