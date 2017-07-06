VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   70
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.Image Image1 
         Height          =   240
         Left            =   240
         Picture         =   "Form2.frx":000C
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ViewWizard "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = "ViewWizard " & App.Major & "." & App.Minor & "." & App.Revision _
& vbCrLf & "作者:远方" & vbCrLf & "QQ:369823808" & vbCrLf & "Email:zzmzzff@163.com"
If AlwaysOnTop Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub

