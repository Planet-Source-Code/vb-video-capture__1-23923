VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About vbVidCap"
   ClientHeight    =   3435
   ClientLeft      =   1665
   ClientTop       =   3420
   ClientWidth     =   4695
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   1950
      TabIndex        =   0
      Top             =   2970
      Width           =   960
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   3360
      Picture         =   "About.frx":000C
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   720
      Picture         =   "About.frx":044E
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "written entirely in Visual Basic"
      Height          =   225
      Index           =   2
      Left            =   75
      TabIndex        =   3
      Top             =   1245
      Width           =   4590
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "vbVidCap - a full featured video capture application"
      Height          =   225
      Index           =   1
      Left            =   75
      TabIndex        =   2
      Top             =   870
      Width           =   4590
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "Copyright (C) 1998-2000 by Sharon Elharar"
      Height          =   225
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   495
      Width           =   4590
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_Load()

End Sub

Private Sub lblAbout_Click(Index As Integer)
    If Index = 4 Then
        Call HyperJump(lblAbout(4).Caption)
    End If
End Sub

Private Function HyperJump(ByVal URL As String) As Long
   HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function
