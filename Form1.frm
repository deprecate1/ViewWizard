VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ViewWizard"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   7635
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   5655
   ScaleWidth      =   7635
   Begin VB.CommandButton Command20 
      Caption         =   "关于"
      Height          =   330
      Left            =   1320
      TabIndex        =   70
      Top             =   5280
      Width           =   1080
   End
   Begin VB.CommandButton Command19 
      Caption         =   "选项..."
      Height          =   330
      Left            =   120
      TabIndex        =   69
      Top             =   5280
      Width           =   1080
   End
   Begin VB.CommandButton Command18 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3750
      TabIndex        =   68
      Top             =   4560
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   540
      Left            =   2760
      Picture         =   "Form1.frx":492A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   67
      ToolTipText     =   "移动窗口"
      Top             =   5040
      Width           =   540
   End
   Begin VB.TextBox Text24 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2760
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   66
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text23 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2760
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   65
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text22 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   62
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text21 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   61
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   58
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   57
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4980
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   54
      Top             =   2040
      Width           =   2505
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4980
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   53
      Top             =   1680
      Width           =   2505
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4980
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   50
      Top             =   1320
      Width           =   2505
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4980
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   43
      Top             =   960
      Width           =   2505
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   3360
      Picture         =   "Form1.frx":55F4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   40
      ToolTipText     =   "拖动进行查找"
      Top             =   5040
      Width           =   540
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1575
      Top             =   0
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4140
      TabIndex        =   33
      Top             =   2400
      Width           =   3435
   End
   Begin VB.CommandButton Command4 
      Caption         =   "修改窗口标题"
      Height          =   330
      Left            =   4140
      TabIndex        =   32
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "发送到组合框"
      Height          =   330
      Left            =   5820
      TabIndex        =   31
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4980
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   30
      Top             =   600
      Width           =   2505
   End
   Begin VB.Frame Frame2 
      Height          =   2325
      Left            =   4140
      TabIndex        =   23
      Top             =   3240
      Width           =   3435
      Begin VB.CommandButton Command17 
         Caption         =   "结束线程"
         Height          =   330
         Left            =   1200
         TabIndex        =   46
         Top             =   1890
         Width           =   960
      End
      Begin VB.CommandButton Command15 
         Caption         =   "强制关闭"
         Height          =   330
         Left            =   2280
         TabIndex        =   45
         Top             =   1470
         Width           =   960
      End
      Begin VB.CommandButton Command16 
         Caption         =   "退出线程"
         Height          =   330
         Left            =   105
         TabIndex        =   44
         Top             =   1890
         Width           =   960
      End
      Begin VB.CommandButton Command8 
         Caption         =   "销毁"
         Height          =   330
         Left            =   1200
         TabIndex        =   41
         Top             =   1470
         Width           =   960
      End
      Begin VB.CommandButton Command14 
         Caption         =   "结束进程"
         Height          =   330
         Left            =   2280
         TabIndex        =   39
         Top             =   1890
         Width           =   960
      End
      Begin VB.CommandButton Command13 
         Caption         =   "取消置顶"
         Height          =   330
         Left            =   1200
         TabIndex        =   38
         Top             =   1050
         Width           =   960
      End
      Begin VB.CommandButton Command12 
         Caption         =   "恢复"
         Height          =   330
         Left            =   2280
         TabIndex        =   37
         Top             =   1050
         Width           =   960
      End
      Begin VB.CommandButton Command11 
         Caption         =   "屏蔽"
         Height          =   330
         Left            =   2280
         TabIndex        =   36
         Top             =   630
         Width           =   960
      End
      Begin VB.CommandButton Command10 
         Caption         =   "隐藏"
         Height          =   330
         Left            =   1200
         TabIndex        =   35
         Top             =   630
         Width           =   960
      End
      Begin VB.CommandButton Command9 
         Caption         =   "显示"
         Height          =   330
         Left            =   105
         TabIndex        =   34
         Top             =   630
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "最小化"
         Height          =   330
         Left            =   105
         TabIndex        =   28
         Top             =   210
         Width           =   960
      End
      Begin VB.CommandButton Command2 
         Caption         =   "最大化"
         Height          =   330
         Left            =   1200
         TabIndex        =   27
         Top             =   210
         Width           =   960
      End
      Begin VB.CommandButton Command3 
         Caption         =   "正常"
         Height          =   330
         Left            =   2280
         TabIndex        =   26
         Top             =   210
         Width           =   960
      End
      Begin VB.CommandButton Command6 
         Caption         =   "关闭"
         Height          =   330
         Left            =   105
         TabIndex        =   25
         Top             =   1470
         Width           =   960
      End
      Begin VB.CommandButton Command7 
         Caption         =   "置顶"
         Height          =   330
         Left            =   105
         TabIndex        =   24
         Top             =   1050
         Width           =   960
      End
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6600
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   20
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      Top             =   233
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "窗口信息"
      Height          =   4095
      Left            =   50
      TabIndex        =   9
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text15 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   48
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   2145
         Width           =   2655
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   3600
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "窗口类值"
         Height          =   180
         Left            =   240
         TabIndex        =   47
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "窗口ID"
         Height          =   180
         Left            =   240
         TabIndex        =   18
         Top             =   2535
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "扩展样式"
         Height          =   180
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "窗口样式"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   1815
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "窗口句柄"
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "父窗标题"
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   3600
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "窗口类名"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "窗口标题"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "父窗句柄"
         Height          =   180
         Left            =   240
         MouseIcon       =   "Form1.frx":61B6
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "父窗类名"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   3255
         Width           =   720
      End
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "下"
      Height          =   180
      Left            =   2520
      TabIndex        =   64
      Top             =   4680
      Width           =   180
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "右"
      Height          =   180
      Left            =   2520
      TabIndex        =   63
      Top             =   4320
      Width           =   180
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "高"
      Height          =   180
      Left            =   1320
      TabIndex        =   60
      Top             =   4725
      Width           =   180
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "宽"
      Height          =   180
      Left            =   1320
      TabIndex        =   59
      Top             =   4365
      Width           =   180
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "上"
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   4725
      Width           =   180
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "左"
      Height          =   180
      Left            =   120
      TabIndex        =   55
      Top             =   4365
      Width           =   180
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "模块路径"
      Height          =   180
      Left            =   4140
      MouseIcon       =   "Form1.frx":6308
      MousePointer    =   99  'Custom
      TabIndex        =   52
      Top             =   2085
      Width           =   720
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "模块名称"
      Height          =   180
      Left            =   4140
      TabIndex        =   51
      Top             =   1725
      Width           =   720
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "实例句柄"
      Height          =   180
      Left            =   4140
      TabIndex        =   49
      Top             =   1365
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "进程路径"
      Height          =   180
      Left            =   4140
      MouseIcon       =   "Form1.frx":645A
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "进程名称"
      Height          =   180
      Left            =   4140
      TabIndex        =   29
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "进程ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6000
      TabIndex        =   22
      Top             =   278
      Width           =   525
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "线程ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   21
      Top             =   278
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xy As POINTAPI, lPos As POINTAPI, sPos As POINTAPI, sPos2 As POINTAPI
Dim WndText As String
Dim pWndText As String
Dim szClsName As String
Dim szpClsName As String
Dim lngClass As Long
Dim H As Long, pH As Long
Dim WndStl As String, pWndStl As String
Dim WndId As Long
Dim WndRect As RECT
Dim PID As Long
Dim hFore As Long
Dim hInstance As Long
Dim nLeft As Long, nTop As Long
Dim Enkey As Boolean
Dim TmpCur As Long
Dim hcursor As Long
Dim hDesktop As Long
Dim CurPId As Long
Dim lRect As RECT
Dim nRet As Long
Dim lngRet As Long
Dim strParentHwnd As String
Dim strParentClass As String
Dim strParentText As String

Private Sub Check2_Click()
If Check2.Value = 1 Then
   Timer1.Enabled = True
Else
   Timer1.Enabled = False
   InvalidateRect 0, 0, True
   DrawFrame hSave, OnTopDraw, DrawType
   hSave = 0
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
  Timer2.Enabled = True
Else
  Timer2.Enabled = False
End If
End Sub


Private Sub Command1_Click()
CloseWindow H
End Sub

Private Sub Command10_Click()
ShowWindow H, 0
End Sub

Private Sub Command11_Click()
EnableWindow H, False
End Sub

Private Sub Command12_Click()
EnableWindow H, True
End Sub

Private Sub Command13_Click()
SetWindowPos H, -2, 0, 0, 0, 0, 3
End Sub

Private Sub Command14_Click()
If Len(Text12) = 0 Then Exit Sub
Dim hProcess As Long
hProcess = OpenProcess(PROCESS_TERMINATE, False, Val(Text12.Text))
TerminateProcess hProcess, 0
CloseHandle hProcess
End Sub

Private Sub Command15_Click()
Dim frm As New Form1
SetParent H, frm.hWnd
Unload frm
End Sub

Private Sub Command16_Click()
'退出线程
If Len(Text11) = 0 Then Exit Sub
PostThreadMessage Val(Text11.Text), &H12, 0, 0
If IsWindow(H) Then PostMessage H, &H12, 0, 0
End Sub

Private Sub Command17_Click()
If Len(Text11) = 0 Then Exit Sub
Dim hThread As Long
hThread = OpenThread(THREAD_TERMINATE, False, Val(Text11.Text))
TerminateThread hThread, 0&
CloseHandle hThread
End Sub

Private Sub Command18_Click()
If Me.Width <= 4155 Then
    Me.Width = 7715
    Command18.Caption = "<"
    NormalSize = False
Else
    Me.Width = 4155
    Command18.Caption = ">"
    NormalSize = True
End If
End Sub

Private Sub Command19_Click()
Load Form3
Form3.Show vbModal
End Sub

Private Sub Command2_Click()
ShowWindow H, 3
End Sub

Private Sub Command20_Click()
Load Form2
Form2.Show vbModal
End Sub

Private Sub Command3_Click()
ShowWindow H, 1
End Sub

Private Sub Command4_Click()
SendMessage H, WM_SETTEXT, 255, ByVal Text7.Text
End Sub

Private Sub Command5_Click()
SendMessage H, CB_ADDSTRING, 255, ByVal Text7.Text
End Sub

Private Sub Command6_Click()
PostMessage H, &H10, 0&, 0
If IsWindow(H) Then
    PostMessage H, WM_SYSCOMMAND, SC_CLOSE, 0
End If
End Sub

Private Sub Command7_Click()
SetWindowPos H, -1, 0, 0, 0, 0, 3
End Sub


Private Sub Command8_Click()
PostMessage H, WM_DESTROY, 0&, 0
If IsWindow(H) Then
    PostMessage H, WM_NCDESTROY, 0&, 0
    PostMessage H, WM_NCDESTROY, 0&, 0
End If
End Sub

Private Sub Command9_Click()
ShowWindow H, 5
End Sub

Private Sub Form_Load()
App.TaskVisible = False
ConfigFile = App.Path & "\Settings.ini"
GetSettingsInfo
Me.Width = IIf(NormalSize, 4155, 7715)
Command18.Caption = IIf(NormalSize, ">", "<")
Me.Height = 6135
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Me.Caption = "ViewWizard " & App.Major & "." & App.Minor & "." & App.Revision
MakeFlatTextbox '设置平滑文本框
H = 0
Enkey = True
SetAdjustPrivileges  '设置为调试权限
hDesktop = GetDesktopWindow()
CurPId = GetCurrentProcessId
Call CreateSoundFile
If AlwaysOnTop = True Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub MakeFlatTextbox()
MakeFlat Text1.hWnd
MakeFlat Text2.hWnd
MakeFlat Text3.hWnd
MakeFlat Text4.hWnd
MakeFlat Text5.hWnd
MakeFlat Text6.hWnd
MakeFlat Text7.hWnd
MakeFlat Text8.hWnd
MakeFlat Text9.hWnd
MakeFlat Text10.hWnd
MakeFlat Text11.hWnd
MakeFlat Text12.hWnd
MakeFlat Text13.hWnd
MakeFlat Text14.hWnd
MakeFlat Text15.hWnd
MakeFlat Text16.hWnd
MakeFlat Text17.hWnd
MakeFlat Text18.hWnd
MakeFlat Text19.hWnd
MakeFlat Text20.hWnd
MakeFlat Text21.hWnd
MakeFlat Text22.hWnd
MakeFlat Text23.hWnd
MakeFlat Text24.hWnd
MakeFlat Picture1.hWnd
MakeFlat Picture2.hWnd
MakeFlat Command1.hWnd
MakeFlat Command2.hWnd
MakeFlat Command3.hWnd
MakeFlat Command4.hWnd
MakeFlat Command5.hWnd
MakeFlat Command6.hWnd
MakeFlat Command7.hWnd
MakeFlat Command8.hWnd
MakeFlat Command9.hWnd
MakeFlat Command10.hWnd
MakeFlat Command11.hWnd
MakeFlat Command12.hWnd
MakeFlat Command13.hWnd
MakeFlat Command14.hWnd
MakeFlat Command15.hWnd
MakeFlat Command16.hWnd
MakeFlat Command17.hWnd
MakeFlat Command18.hWnd
MakeFlat Command19.hWnd
MakeFlat Command20.hWnd
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveSettingsInfo
DeleteFile sndFile
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim TmpPath As String
TmpPath = Text14.Text
If Fe(TmpPath) = True Then
    WinExec "explorer.exe /select," & TmpPath, 1
End If
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim TmpPath As String
TmpPath = Text18.Text
If Fe(TmpPath) = True Then
    WinExec "explorer.exe /select," & TmpPath, 1
End If
End Sub

Private Sub Label4_DblClick()
If Text4.Text <> "" Then
        Text1.Text = Val(Mid$(Text4.Text, 2))
End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Text4.Text <> "" Then Text1.Text = Val(Mid$(Text4.Text, 2))
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button <> 1 Then Exit Sub
If AutoFindWin = True Then Exit Sub
ThisCur = GetCursor
TmpCur = CopyIcon(ThisCur)
hcursor = CopyIcon(Picture1.Picture)
r = SetSystemCursor(hcursor, OCR_NORMAL)
Picture1.MouseIcon = Picture1.Picture
Picture1.Picture = Nothing
H = 0
GetCursorPos sPos2
SetCursorPos sPos2.X + 1, sPos2.y
If OpenSound = True Then PlaySnd
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 And Timer1.Enabled = False Then
   If OpenHideMode = True Then Me.Hide
   Timer1.Enabled = True
End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button <> 1 Then Exit Sub
If AutoFindWin = True Then Exit Sub
If IsWindow(H) = 0 Then Exit Sub
AttachThread H, True
SetCapture Picture2.hWnd
SetForegroundWindow H
AttachThread H, True
SetCapture Picture2.hWnd
Picture2.MousePointer = 99
Picture2.MouseIcon = Picture2.Picture
Picture2.Picture = Nothing
GetCursorPos sPos
GetWindowRect H, lRect
SetCursorPos lRect.Left, lRect.Top
If OpenSound = True Then PlaySnd
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button <> 1 Then Exit Sub
If AutoFindWin = True Then Exit Sub
If IsWindow(H) = 0 Then Exit Sub
GetCursorPos lPos
nLeft = lPos.X: nTop = lPos.y
If GetParentW(H) <> 0 Then ScreenToClient GetParentW(H), lPos
GetWindowRect H, lRect
'Debug.Print lPos.X, lPos.y
MoveWindow H, lPos.X, lPos.y, lRect.Right - lRect.Left, lRect.Bottom - lRect.Top, True
Text19.Text = nLeft
Text20.Text = nTop
Text23.Text = Screen.Width / Screen.TwipsPerPixelX - nLeft - Val(Text21)
Text24.Text = Screen.Height / Screen.TwipsPerPixelY - nTop - Val(Text22)
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button <> 1 Then Exit Sub
If AutoFindWin = True Then Exit Sub
If IsWindow(H) = 0 Then Exit Sub
SetForegroundWindow Me.hWnd
AttachThread H, False
Call ReleaseCapture
AttachThread H, False
Call ReleaseCapture
Picture2.MousePointer = 0
Picture2.Picture = Picture2.MouseIcon
SetCursorPos sPos.X, sPos.y
SetForegroundWindow Me.hWnd
If OpenSound = True Then PlaySnd
End Sub

Private Sub Text1_Change()
'句柄最大值&H7fffffff
If Val(Text1.Text) > &H7FFFFFFF Then Text1.Text = CStr(2147483647)
H = Val(Text1.Text)
Call GetWindowThreadProcessId(H, PID)
If IsWindow(H) = 0 Then
  Text2.Text = ""
  Text3.Text = ""
  Text4.Text = ""
  Text5.Text = ""
  Text6.Text = ""
  Text8.Text = ""
  Text9.Text = ""
  Text10.Text = ""
  Text11.Text = ""
  Text12.Text = ""
  Text13.Text = ""
  Text14.Text = ""
  Text15.Text = ""
  Text16.Text = ""
  Text17.Text = ""
  Text18.Text = ""
  Text19.Text = ""
  Text20.Text = ""
  Text21.Text = ""
  Text22.Text = ""
  Text23.Text = ""
  Text24.Text = ""
  If Text1.Text <> "" Then
      Text1.ToolTipText = "句柄无效"
  Else
      Text1.ToolTipText = ""
  End If
  Exit Sub  '句柄无效退出
End If
Text1.ToolTipText = ""
WndText = String$(255, vbNullChar)
pWndText = String$(255, vbNullChar)
szClsName = String$(255, vbNullChar)
szpClsName = String$(255, vbNullChar)
strParentHwnd = ""
strParentClass = ""
strParentText = ""
WndStl = ""
pWndStl = ""
If H = hDesktop Then
    WndText = "[Desktop]" + Chr$(0)
Else
    nRet = SendMessage(H, EM_GETPASSWORDCHAR, 0&, 0)
    If nRet <> 42 Then
        lngRet = GetWindowText(H, WndText, Len(WndText))
        If lngRet = 0 Then
            SendMessage H, WM_GETTEXT, Len(WndText), ByVal WndText
        End If
    Else
        PostMessage H, EM_SETPASSWORDCHAR, 0&, 0&
        Sleep 10
        lngRet = GetWindowText(H, WndText, Len(WndText))
        If lngRet = 0 Then
            Call SendMessage(H, WM_GETTEXT, Len(WndText), ByVal WndText)
        End If
        PostMessage H, EM_SETPASSWORDCHAR, nRet, 0&
    End If
End If
GetClassName H, szClsName, 255
Text2.Text = CheckStr(szClsName)
Text3.Text = CheckStr(WndText)
pH = GetAncestor(H, GA_PARENT)
If pH = hDesktop Then
    strParentText = "[Desktop]" & Chr$(0)
Else
    SendMessage pH, WM_GETTEXT, 255, ByVal pWndText
    strParentText = "[" & CheckStr(pWndText) & "]"
End If
If pH <> 0 Then
        GetClassName pH, szpClsName, 255
        strParentHwnd = "[" & CStr(pH) & "]"
        strParentClass = "[" & CheckStr(szpClsName) & "]"
        pH = GetAncestor(pH, GA_PARENT)
        Do While pH <> 0
                strParentHwnd = strParentHwnd & "[" & pH & "]"
                GetClassName pH, szpClsName, 255
                strParentClass = strParentClass & "[" & CheckStr(szpClsName) & "]"
                If pH = hDesktop Then
                    strParentText = strParentText & "[Desktop]" & Chr$(0)
                Else
                    SendMessage pH, WM_GETTEXT, 255, ByVal pWndText
                    strParentText = strParentText & "[" & CheckStr(pWndText) & "]"
                End If
                pH = GetAncestor(pH, GA_PARENT)
        Loop
        Text4.Text = strParentHwnd
        Text5.Text = strParentClass
        Text6.Text = strParentText
Else
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
End If
lngClass = GetClassLong(H, GCW_ATOM)
Text15.Text = lngClass
WndStl = GetWindowStyle(H)
Text8.Text = WndStl
pWndStl = GetWindowExStyle(H)
Text9.Text = pWndStl
WndId = GetWindowLong(H, GWL_ID)
If WndId = 0 Then
   Text10.Text = ""
Else
   Text10.Text = Trim(Str$(WndId))
End If
Text11.Text = GetWindowThreadProcessId(H, PID)
Text12.Text = PID
Text13.Text = GetProcessNameByPID(PID)
Text14.Text = GetProcessPath(PID)
hInstance = GetWinhInstance(H)
Text16.Text = "0x" & Hex(hInstance)
Text17.Text = GetModuleName(H)
Text18.Text = GetModulePath(H)
Call GetWndSizePos
If H <> hSave Then
    DrawFrame hSave, OnTopDraw, DrawType
    DrawFrame H, OnTopDraw, DrawType
    hSave = H
End If
End Sub

Private Sub Text1_GotFocus()
    If IsWindow(Val(Text1.Text)) = 0 And Text1.Text <> "" Then
        Text1.ToolTipText = "句柄无效"
    Else
        Text1.ToolTipText = ""
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Enkey = True Then Enkey = False: Exit Sub
If KeyAscii > 57 Or KeyAscii < 48 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text14_Change()
Label13.ToolTipText = IIf(Text14.Text <> "", Text14.Text, "")
End Sub

Private Sub Text18_Change()
Label17.ToolTipText = IIf(Text18.Text <> "", Text18.Text, "")
End Sub

Private Sub Text4_Change()
If Text4.Text = "0" Then
Text4.Text = ""
End If
End Sub
Private Sub jubing()
Dim jb As Long
Dim hpc As String * 255
Dim hpa As Long
Dim hpt As String * 255
If Text4.Text <> "" Then
 jb = CLng(Text4.Text)
 hpa = GetParent(jb)
 If hpa <> 0 Then
   Text4.Text = Text4.Text + "|" + Trim$(CStr(hpa))
   GetClassName hpa, hpc, 255
   Text5.Text = Text5.Text + "|" + Trim(hpc)
   SendMessage hpa, WM_GETTEXT, 255, ByVal hpt
   Text6.Text = Text6.Text + "|" + Trim(hpt)
 End If
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If InStr("XCV", UCase$(Chr$(KeyCode))) And Shift = 2 Then
  Enkey = True
Else
  Enkey = False
End If
End Sub

Private Sub Timer1_Timer()
If AutoFindWin = False Then
   If LButtonDown = False Then
         Timer1.Enabled = False
         If OpenHideMode = True Then
            Me.Show
            If Me.WindowState <> 0 Then ShowWindow Me.hWnd, 9
            SetForegroundWindow Me.hWnd
         End If
         Picture1.Picture = Picture1.MouseIcon
         SetSystemCursor TmpCur, OCR_NORMAL
         DestroyCursor hcursor
         H = Val(Text1)
         If OpenSound = True Then PlaySnd
         InvalidateRect ByVal 0&, ByVal 0&, True
         hSave = 0
         Exit Sub
   End If
End If
GetCursorPos xy
a = xy.X
b = xy.y
H = WindowFromPoint(a, b)
GetWindowThreadProcessId H, PID
If PID = CurPId Then
   H = Val(Text1.Text)
   Exit Sub
End If

    '捕获屏蔽、隐藏以及透明窗口
    TmpHwnd = H
    'Debug.Print H
    Call ScreenToClient(H, xy)
    H = ChildWindowFromPoint(H, xy.X, xy.y)
    'Debug.Print H
    If H = 0 Then H = TmpHwnd
    If NotFindHideWin = True Then
        If IsWindowVisible(H) = 0 Then H = TmpHwnd
    End If
    '检查窗口
    H = CheckHwnd(H)
    
Text1.Text = Trim$(Str$(H))
Text1_Change
End Sub

'显示和激活隐藏或无效对象
Private Sub Timer2_Timer()
On Error Resume Next
Dim i As Long
If EnableWin = True Then
        wCount = 0
        EnumChildWindows GetForegroundWindow(), AddressOf EnumCallback, 0
        If wCount = 0 Then Exit Sub
        For i = 1 To UBound(HwndChild)
            EnableWindow HwndChild(i), True
        Next i
End If
If ShowHideWin = True Then
        wCount = 0
        EnumChildWindows GetForegroundWindow(), AddressOf EnumCallback, 0
        If wCount = 0 Then Exit Sub
        For i = 1 To UBound(HwndChild)
                ShowWindow HwndChild(i), 5
        Next i
End If
End Sub

'坐标
Sub GetWndSizePos()
GetWindowRect H, WndRect
Text19.Text = WndRect.Left
Text20.Text = WndRect.Top
Text21.Text = WndRect.Right - WndRect.Left
Text22.Text = WndRect.Bottom - WndRect.Top
Text23.Text = Screen.Width / Screen.TwipsPerPixelX - WndRect.Right
Text24.Text = Screen.Height / Screen.TwipsPerPixelY - WndRect.Bottom
End Sub

'连接窗口线程
Sub AttachThread(ByVal hWin As Long, ByVal bAttach As Boolean)
Dim ThreadID As Long
ThreadID = GetWindowThreadProcessId(hWin, 0&)
AttachThreadInput App.ThreadID, ThreadID, IIf(bAttach, 1, 0)
End Sub

