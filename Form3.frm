VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Caption         =   "其它选项"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1815
      Begin VB.CheckBox Check5 
         Caption         =   "忽略隐藏窗口"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Check7 
         Caption         =   "启用隐藏模式"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check6 
         Caption         =   "自动捕获窗口"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   275
      Left            =   2520
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   275
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "基本选项"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      Begin VB.CheckBox Check4 
         Caption         =   "声音提示"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "显示隐藏对象"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "恢复无效对象"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1485
      End
      Begin VB.CheckBox Check1 
         Caption         =   "总在最前"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "绘图方式"
      Height          =   1695
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.CheckBox Check8 
         Caption         =   "在顶层绘图"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "聚焦矩形"
         Height          =   255
         Left            =   105
         TabIndex        =   9
         Top             =   960
         Width           =   1200
      End
      Begin VB.OptionButton Option3 
         Caption         =   "反色高亮"
         Height          =   255
         Left            =   105
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "边缘高亮"
         Height          =   255
         Left            =   105
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "不提示"
         Height          =   255
         Left            =   105
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
AlwaysOnTop = Check1.Value
EnableWin = Check2.Value
ShowHideWin = Check3.Value
OpenSound = Check4.Value
NotFindHideWin = Check5.Value
AutoFindWin = Check6.Value
OpenHideMode = Check7.Value
OnTopDraw = Check8.Value
If Option1.Value = True Then
        DrawType = 0
ElseIf Option2.Value = True Then
        DrawType = 1
ElseIf Option3.Value = True Then
        DrawType = 2
Else
        DrawType = 3
End If
SetWindowPos Form1.hWnd, IIf(AlwaysOnTop, -1, -2), 0, 0, 0, 0, 3
If Check2.Value = 1 Or Check3.Value = 1 Then
        Form1.Timer2.Enabled = True
Else
        Form1.Timer2.Enabled = False
End If
Form1.Timer1.Enabled = AutoFindWin
If AutoFindWin = False Then
        InvalidateRect 0, 0, True
        DrawFrame hSave, OnTopDraw, DrawType
        hSave = 0
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If AlwaysOnTop = True Then
        SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End If
Check1.Value = IIf(AlwaysOnTop, 1, 0)
Check2.Value = IIf(EnableWin, 1, 0)
Check3.Value = IIf(ShowHideWin, 1, 0)
Check4.Value = IIf(OpenSound, 1, 0)
Check5.Value = IIf(NotFindHideWin, 1, 0)
Check6.Value = IIf(AutoFindWin, 1, 0)
Check7.Value = IIf(OpenHideMode, 1, 0)
Check8.Value = IIf(OnTopDraw, 1, 0)
If DrawType = 0 Then
        Option1.Value = True
ElseIf DrawType = 1 Then
        Option2.Value = True
ElseIf DrawType = 2 Then
        Option3.Value = True
Else
        Option4.Value = True
End If
If Option1.Value = True Then Check8.Enabled = False
End Sub

Private Sub Option1_Click()
Check8.Enabled = False
End Sub

Private Sub Option2_Click()
Check8.Enabled = True
End Sub

Private Sub Option3_Click()
Check8.Enabled = True
End Sub

Private Sub Option4_Click()
Check8.Enabled = True
End Sub
