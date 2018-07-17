VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Zero Temperature Compensation Application 一课传感器调零助手"
   ClientHeight    =   8010
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8667.538
   ScaleMode       =   0  'User
   ScaleWidth      =   13727.32
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timerTimeout 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   4680
      Top             =   0
   End
   Begin VB.Timer timerAuto 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4200
      Top             =   0
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   1200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer timerUart 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer timerOvenTemp 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   3720
      Top             =   0
   End
   Begin VB.Timer timerCHTest 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3240
      Top             =   0
   End
   Begin VB.Timer timerTest 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2760
      Top             =   0
   End
   Begin VB.Timer timerTemp 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2280
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "系数设置"
      Height          =   375
      Left            =   120
      TabIndex        =   125
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "终止测试"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton btnPause 
      Caption         =   "暂停测试"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "开始测试"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   7320
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7740
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18283
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2013-1-30"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "13:47"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   9600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   10080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "全局设定"
      TabPicture(0)   =   "frmMain.frx":0336
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame21"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "通道设置 （一、二）"
      TabPicture(1)   =   "frmMain.frx":0352
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "tb1680ComInput"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cbbBTest"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "tb1680ComReading"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame8"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame9"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame10"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "通道设置 （三、四）"
      TabPicture(2)   =   "frmMain.frx":036E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame16"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame15"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame14"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame13"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame12"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame11"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "通道设置 （五）、通道检测"
      TabPicture(3)   =   "frmMain.frx":038A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame20(13)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame19"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame18"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame17"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame21 
         Caption         =   "端口设置"
         Height          =   855
         Left            =   120
         TabIndex        =   599
         Top             =   5760
         Width           =   3855
         Begin VB.CommandButton btnSetCOM 
            Caption         =   "修改"
            Height          =   255
            Left            =   120
            TabIndex        =   603
            Top             =   510
            Width           =   1455
         End
         Begin VB.CommandButton btnCOMApply 
            Caption         =   "应用"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2280
            TabIndex        =   602
            Top             =   510
            Width           =   1455
         End
         Begin VB.TextBox tbTempCOM 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   601
            Text            =   "0"
            Top             =   160
            Width           =   375
         End
         Begin VB.TextBox tb1680COM 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            TabIndex        =   600
            Text            =   "0"
            Top             =   160
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "烘箱：  COM"
            Height          =   255
            Left            =   120
            TabIndex        =   605
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "TI1680：  COM"
            Height          =   255
            Left            =   2160
            TabIndex        =   604
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "通道连通性检测"
         Height          =   6255
         Index           =   13
         Left            =   -68160
         TabIndex        =   472
         Top             =   360
         Width           =   6615
         Begin VB.TextBox tbAuto 
            Height          =   285
            Left            =   5570
            TabIndex        =   609
            Text            =   "5"
            Top             =   1290
            Width           =   375
         End
         Begin VB.CheckBox cbAuto 
            Caption         =   "Check13"
            Height          =   255
            Left            =   4970
            TabIndex        =   608
            Top             =   1300
            Width           =   255
         End
         Begin VB.TextBox tbCH 
            BeginProperty Font 
               Name            =   "Elephant"
               Size            =   45
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   598
            Text            =   "CH"
            Top             =   240
            Width           =   4455
         End
         Begin VB.TextBox tbCHTest 
            BeginProperty Font 
               Name            =   "Elephant"
               Size            =   45
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   597
            Text            =   "0.00000"
            Top             =   1440
            Width           =   4455
         End
         Begin VB.CommandButton btnCHTestDown 
            Caption         =   "↓"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   26.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5640
            TabIndex        =   596
            Top             =   1680
            Width           =   855
         End
         Begin VB.CommandButton btnCHTestUp 
            Caption         =   "↑"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   26.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4680
            TabIndex        =   595
            Top             =   1680
            Width           =   855
         End
         Begin VB.CommandButton btnStopCHTest 
            Caption         =   "停止"
            Height          =   855
            Left            =   5640
            TabIndex        =   594
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton btnStartCHTest 
            Caption         =   "开始"
            Height          =   855
            Left            =   4680
            TabIndex        =   593
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   59
            Left            =   5360
            TabIndex        =   592
            Top             =   5940
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   58
            Left            =   5360
            TabIndex        =   591
            Top             =   5640
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   57
            Left            =   5360
            TabIndex        =   590
            Top             =   5340
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   56
            Left            =   5360
            TabIndex        =   589
            Top             =   5040
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   55
            Left            =   5360
            TabIndex        =   588
            Top             =   4740
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   54
            Left            =   5360
            TabIndex        =   587
            Top             =   4440
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   53
            Left            =   5360
            TabIndex        =   586
            Top             =   4140
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   52
            Left            =   5360
            TabIndex        =   585
            Top             =   3840
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   51
            Left            =   5360
            TabIndex        =   584
            Top             =   3540
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   50
            Left            =   5360
            TabIndex        =   583
            Top             =   3240
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   49
            Left            =   5360
            TabIndex        =   582
            Top             =   2940
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   48
            Left            =   5360
            TabIndex        =   581
            Top             =   2640
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   47
            Left            =   4110
            TabIndex        =   580
            Top             =   5940
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   46
            Left            =   4110
            TabIndex        =   579
            Top             =   5640
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   45
            Left            =   4110
            TabIndex        =   578
            Top             =   5340
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   44
            Left            =   4110
            TabIndex        =   577
            Top             =   5040
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   43
            Left            =   4110
            TabIndex        =   576
            Top             =   4740
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   42
            Left            =   4110
            TabIndex        =   575
            Top             =   4440
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   41
            Left            =   4110
            TabIndex        =   574
            Top             =   4140
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   40
            Left            =   4110
            TabIndex        =   573
            Top             =   3840
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   39
            Left            =   4110
            TabIndex        =   572
            Top             =   3540
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   38
            Left            =   4110
            TabIndex        =   571
            Top             =   3240
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   37
            Left            =   4110
            TabIndex        =   570
            Top             =   2940
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   36
            Left            =   4110
            TabIndex        =   569
            Top             =   2640
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   35
            Left            =   2860
            TabIndex        =   568
            Top             =   5940
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   34
            Left            =   2860
            TabIndex        =   567
            Top             =   5640
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   33
            Left            =   2860
            TabIndex        =   566
            Top             =   5340
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   32
            Left            =   2860
            TabIndex        =   565
            Top             =   5040
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   31
            Left            =   2860
            TabIndex        =   564
            Top             =   4740
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   30
            Left            =   2860
            TabIndex        =   563
            Top             =   4440
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   29
            Left            =   2860
            TabIndex        =   562
            Top             =   4140
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   28
            Left            =   2860
            TabIndex        =   561
            Top             =   3840
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   27
            Left            =   2860
            TabIndex        =   560
            Top             =   3540
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   26
            Left            =   2860
            TabIndex        =   559
            Top             =   3240
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   25
            Left            =   2860
            TabIndex        =   558
            Top             =   2940
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   24
            Left            =   2860
            TabIndex        =   557
            Top             =   2640
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   23
            Left            =   1610
            TabIndex        =   556
            Top             =   5940
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   22
            Left            =   1610
            TabIndex        =   555
            Top             =   5640
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   21
            Left            =   1610
            TabIndex        =   554
            Top             =   5340
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   20
            Left            =   1610
            TabIndex        =   553
            Top             =   5040
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   19
            Left            =   1610
            TabIndex        =   552
            Top             =   4740
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   18
            Left            =   1610
            TabIndex        =   551
            Top             =   4440
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   17
            Left            =   1610
            TabIndex        =   550
            Top             =   4140
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   16
            Left            =   1610
            TabIndex        =   549
            Top             =   3840
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   15
            Left            =   1610
            TabIndex        =   548
            Top             =   3540
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   14
            Left            =   1610
            TabIndex        =   547
            Top             =   3240
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   13
            Left            =   1610
            TabIndex        =   546
            Top             =   2940
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   12
            Left            =   1610
            TabIndex        =   545
            Top             =   2640
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   11
            Left            =   360
            TabIndex        =   544
            Top             =   5940
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   10
            Left            =   360
            TabIndex        =   543
            Top             =   5640
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   542
            Top             =   5340
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   541
            Top             =   5040
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   540
            Top             =   4740
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   539
            Top             =   4440
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   538
            Top             =   4140
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   537
            Top             =   3840
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   536
            Top             =   3540
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   535
            Top             =   3240
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   534
            Top             =   2940
            Width           =   255
         End
         Begin VB.OptionButton rbCHTest 
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   473
            Top             =   2640
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.Label Label19 
            Caption         =   "秒"
            Height          =   255
            Left            =   6000
            TabIndex        =   610
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "自动"
            Height          =   255
            Left            =   5210
            TabIndex        =   607
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH60"
            Height          =   255
            Index           =   59
            Left            =   5715
            TabIndex        =   533
            Top             =   5950
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH59"
            Height          =   255
            Index           =   58
            Left            =   5715
            TabIndex        =   532
            Top             =   5650
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH58"
            Height          =   255
            Index           =   57
            Left            =   5715
            TabIndex        =   531
            Top             =   5350
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH57"
            Height          =   255
            Index           =   56
            Left            =   5715
            TabIndex        =   530
            Top             =   5050
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH56"
            Height          =   255
            Index           =   55
            Left            =   5715
            TabIndex        =   529
            Top             =   4750
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH55"
            Height          =   255
            Index           =   54
            Left            =   5715
            TabIndex        =   528
            Top             =   4450
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH54"
            Height          =   255
            Index           =   53
            Left            =   5715
            TabIndex        =   527
            Top             =   4150
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH53"
            Height          =   255
            Index           =   52
            Left            =   5715
            TabIndex        =   526
            Top             =   3850
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH52"
            Height          =   255
            Index           =   51
            Left            =   5715
            TabIndex        =   525
            Top             =   3550
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH51"
            Height          =   255
            Index           =   50
            Left            =   5715
            TabIndex        =   524
            Top             =   3250
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH50"
            Height          =   255
            Index           =   49
            Left            =   5715
            TabIndex        =   523
            Top             =   2950
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH48"
            Height          =   255
            Index           =   47
            Left            =   4470
            TabIndex        =   522
            Top             =   5950
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH47"
            Height          =   255
            Index           =   46
            Left            =   4470
            TabIndex        =   521
            Top             =   5650
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH46"
            Height          =   255
            Index           =   45
            Left            =   4470
            TabIndex        =   520
            Top             =   5350
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH45"
            Height          =   255
            Index           =   44
            Left            =   4470
            TabIndex        =   519
            Top             =   5050
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH44"
            Height          =   255
            Index           =   43
            Left            =   4470
            TabIndex        =   518
            Top             =   4750
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH43"
            Height          =   255
            Index           =   42
            Left            =   4470
            TabIndex        =   517
            Top             =   4450
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH42"
            Height          =   255
            Index           =   41
            Left            =   4470
            TabIndex        =   516
            Top             =   4150
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH41"
            Height          =   255
            Index           =   40
            Left            =   4470
            TabIndex        =   515
            Top             =   3850
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH40"
            Height          =   255
            Index           =   39
            Left            =   4470
            TabIndex        =   514
            Top             =   3550
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH39"
            Height          =   255
            Index           =   38
            Left            =   4470
            TabIndex        =   513
            Top             =   3250
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH38"
            Height          =   255
            Index           =   37
            Left            =   4470
            TabIndex        =   512
            Top             =   2950
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH36"
            Height          =   255
            Index           =   35
            Left            =   3225
            TabIndex        =   511
            Top             =   5950
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH35"
            Height          =   255
            Index           =   34
            Left            =   3225
            TabIndex        =   510
            Top             =   5650
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH34"
            Height          =   255
            Index           =   33
            Left            =   3225
            TabIndex        =   509
            Top             =   5350
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH33"
            Height          =   255
            Index           =   32
            Left            =   3225
            TabIndex        =   508
            Top             =   5050
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH32"
            Height          =   255
            Index           =   31
            Left            =   3225
            TabIndex        =   507
            Top             =   4750
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH31"
            Height          =   255
            Index           =   30
            Left            =   3225
            TabIndex        =   506
            Top             =   4450
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH30"
            Height          =   255
            Index           =   29
            Left            =   3225
            TabIndex        =   505
            Top             =   4150
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH29"
            Height          =   255
            Index           =   28
            Left            =   3225
            TabIndex        =   504
            Top             =   3850
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH28"
            Height          =   255
            Index           =   27
            Left            =   3225
            TabIndex        =   503
            Top             =   3550
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH27"
            Height          =   255
            Index           =   26
            Left            =   3225
            TabIndex        =   502
            Top             =   3250
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH26"
            Height          =   255
            Index           =   25
            Left            =   3225
            TabIndex        =   501
            Top             =   2950
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH24"
            Height          =   255
            Index           =   23
            Left            =   1965
            TabIndex        =   500
            Top             =   5950
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH23"
            Height          =   255
            Index           =   22
            Left            =   1965
            TabIndex        =   499
            Top             =   5650
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH22"
            Height          =   255
            Index           =   21
            Left            =   1965
            TabIndex        =   498
            Top             =   5350
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH21"
            Height          =   255
            Index           =   20
            Left            =   1965
            TabIndex        =   497
            Top             =   5050
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH20"
            Height          =   255
            Index           =   19
            Left            =   1965
            TabIndex        =   496
            Top             =   4750
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH19"
            Height          =   255
            Index           =   18
            Left            =   1965
            TabIndex        =   495
            Top             =   4450
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH18"
            Height          =   255
            Index           =   17
            Left            =   1965
            TabIndex        =   494
            Top             =   4150
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH17"
            Height          =   255
            Index           =   16
            Left            =   1965
            TabIndex        =   493
            Top             =   3850
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH16"
            Height          =   255
            Index           =   15
            Left            =   1965
            TabIndex        =   492
            Top             =   3550
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH15"
            Height          =   255
            Index           =   14
            Left            =   1965
            TabIndex        =   491
            Top             =   3250
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH14"
            Height          =   255
            Index           =   13
            Left            =   1965
            TabIndex        =   490
            Top             =   2950
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH49"
            Height          =   255
            Index           =   48
            Left            =   5715
            TabIndex        =   489
            Top             =   2650
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH37"
            Height          =   255
            Index           =   36
            Left            =   4470
            TabIndex        =   488
            Top             =   2650
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH25"
            Height          =   255
            Index           =   24
            Left            =   3225
            TabIndex        =   487
            Top             =   2650
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH13"
            Height          =   255
            Index           =   12
            Left            =   1965
            TabIndex        =   486
            Top             =   2650
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH11"
            Height          =   255
            Index           =   10
            Left            =   720
            TabIndex        =   485
            Top             =   5650
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH10"
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   484
            Top             =   5350
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH9"
            Height          =   255
            Index           =   8
            Left            =   720
            TabIndex        =   483
            Top             =   5050
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH8"
            Height          =   255
            Index           =   7
            Left            =   720
            TabIndex        =   482
            Top             =   4750
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH7"
            Height          =   255
            Index           =   6
            Left            =   720
            TabIndex        =   481
            Top             =   4450
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH6"
            Height          =   255
            Index           =   5
            Left            =   720
            TabIndex        =   480
            Top             =   4150
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH5"
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   479
            Top             =   3850
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH4"
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   478
            Top             =   3550
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH3"
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   477
            Top             =   3250
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH2"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   476
            Top             =   2950
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH12"
            Height          =   255
            Index           =   11
            Left            =   720
            TabIndex        =   475
            Top             =   5950
            Width           =   615
         End
         Begin VB.Label lblCHTest 
            Caption         =   "CH1"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   474
            Top             =   2650
            Width           =   615
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "通道设置（三）"
         Height          =   5655
         Left            =   -74880
         TabIndex        =   432
         Top             =   440
         Width           =   3735
         Begin VB.CheckBox cbCHSec5 
            Caption         =   "Check13"
            Height          =   255
            Left            =   120
            TabIndex        =   458
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox cbB5CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   457
            Top             =   840
            Width           =   255
         End
         Begin VB.CheckBox cbB5CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   456
            Top             =   1240
            Width           =   255
         End
         Begin VB.CheckBox cbB6CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   455
            Top             =   1640
            Width           =   255
         End
         Begin VB.CheckBox cbB6CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   454
            Top             =   2040
            Width           =   255
         End
         Begin VB.CheckBox cbB6CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   453
            Top             =   2440
            Width           =   255
         End
         Begin VB.CheckBox cbB6CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   452
            Top             =   2840
            Width           =   255
         End
         Begin VB.CheckBox cbB6CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   451
            Top             =   3240
            Width           =   255
         End
         Begin VB.CheckBox cbB6CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   450
            Top             =   3640
            Width           =   255
         End
         Begin VB.CheckBox cbB6CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   449
            Top             =   4040
            Width           =   255
         End
         Begin VB.CheckBox cbB6CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   448
            Top             =   4440
            Width           =   255
         End
         Begin VB.ComboBox cbbCHSec5 
            Height          =   315
            ItemData        =   "frmMain.frx":03A6
            Left            =   1080
            List            =   "frmMain.frx":03A8
            TabIndex        =   447
            Top             =   300
            Width           =   2535
         End
         Begin VB.ComboBox cbbB5CH 
            Height          =   315
            Index           =   8
            ItemData        =   "frmMain.frx":03AA
            Left            =   1080
            List            =   "frmMain.frx":03AC
            TabIndex        =   446
            Top             =   800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB5CH 
            Height          =   315
            Index           =   9
            Left            =   1080
            TabIndex        =   445
            Top             =   1200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB6CH 
            Height          =   315
            Index           =   0
            Left            =   1080
            TabIndex        =   444
            Top             =   1600
            Width           =   2535
         End
         Begin VB.ComboBox cbbB6CH 
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   443
            Top             =   2000
            Width           =   2535
         End
         Begin VB.ComboBox cbbB6CH 
            Height          =   315
            Index           =   2
            Left            =   1080
            TabIndex        =   442
            Top             =   2400
            Width           =   2535
         End
         Begin VB.ComboBox cbbB6CH 
            Height          =   315
            Index           =   3
            Left            =   1080
            TabIndex        =   441
            Top             =   2800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB6CH 
            Height          =   315
            Index           =   4
            Left            =   1080
            TabIndex        =   440
            Top             =   3200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB6CH 
            Height          =   315
            Index           =   5
            Left            =   1080
            TabIndex        =   439
            Top             =   3600
            Width           =   2535
         End
         Begin VB.ComboBox cbbB6CH 
            Height          =   315
            Index           =   6
            Left            =   1080
            TabIndex        =   438
            Top             =   4000
            Width           =   2535
         End
         Begin VB.ComboBox cbbB6CH 
            Height          =   315
            Index           =   7
            Left            =   1080
            TabIndex        =   437
            Top             =   4400
            Width           =   2535
         End
         Begin VB.CheckBox cbB6CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   436
            Top             =   4840
            Width           =   255
         End
         Begin VB.CheckBox cbB6CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   435
            Top             =   5240
            Width           =   255
         End
         Begin VB.ComboBox cbbB6CH 
            Height          =   315
            Index           =   8
            Left            =   1080
            TabIndex        =   434
            Top             =   4800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB6CH 
            Height          =   315
            Index           =   9
            Left            =   1080
            TabIndex        =   433
            Top             =   5200
            Width           =   2535
         End
         Begin VB.Label lblCHSec5 
            Caption         =   "全/反选"
            Height          =   255
            Left            =   375
            TabIndex        =   471
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblCH 
            Caption         =   "CH49"
            Height          =   255
            Index           =   48
            Left            =   360
            TabIndex        =   470
            Top             =   870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH50"
            Height          =   255
            Index           =   49
            Left            =   360
            TabIndex        =   469
            Top             =   1270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH51"
            Height          =   255
            Index           =   50
            Left            =   360
            TabIndex        =   468
            Top             =   1670
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH52"
            Height          =   255
            Index           =   51
            Left            =   360
            TabIndex        =   467
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH53"
            Height          =   255
            Index           =   52
            Left            =   360
            TabIndex        =   466
            Top             =   2470
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH54"
            Height          =   255
            Index           =   53
            Left            =   360
            TabIndex        =   465
            Top             =   2870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH55"
            Height          =   255
            Index           =   54
            Left            =   360
            TabIndex        =   464
            Top             =   3270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH56"
            Height          =   255
            Index           =   55
            Left            =   360
            TabIndex        =   463
            Top             =   3670
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH57"
            Height          =   255
            Index           =   56
            Left            =   360
            TabIndex        =   462
            Top             =   4070
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH58"
            Height          =   255
            Index           =   57
            Left            =   360
            TabIndex        =   461
            Top             =   4470
            Width           =   735
         End
         Begin VB.Line Line10 
            X1              =   0
            X2              =   3720
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Label lblCH 
            Caption         =   "CH59"
            Height          =   255
            Index           =   58
            Left            =   360
            TabIndex        =   460
            Top             =   4870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH60"
            Height          =   255
            Index           =   59
            Left            =   360
            TabIndex        =   459
            Top             =   5270
            Width           =   735
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "传感器读数"
         Height          =   5655
         Left            =   -71160
         TabIndex        =   405
         Top             =   440
         Width           =   1935
         Begin VB.TextBox tbB5CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   429
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB5CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   1080
            TabIndex        =   428
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB5CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   427
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB5CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   1080
            TabIndex        =   426
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB6CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   425
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB6CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   424
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB6CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   423
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB6CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   422
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB6CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   421
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB6CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   420
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB6CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   419
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB6CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   418
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB6CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   417
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB6CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   416
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB6CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   415
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB6CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   414
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB6CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   413
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB6CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   412
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB6CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   411
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB6CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   410
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB6CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   409
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB6CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   408
            Top             =   5200
            Width           =   735
         End
         Begin VB.TextBox tbB6CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   1080
            TabIndex        =   407
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB6CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   1080
            TabIndex        =   406
            Top             =   5200
            Width           =   735
         End
         Begin VB.Line Line9 
            X1              =   960
            X2              =   960
            Y1              =   240
            Y2              =   5640
         End
         Begin VB.Label Label88 
            Caption         =   "低温"
            Height          =   255
            Left            =   320
            TabIndex        =   431
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label75 
            Caption         =   "高温"
            Height          =   255
            Left            =   1270
            TabIndex        =   430
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "结果 L"
         Height          =   5655
         Left            =   -69240
         TabIndex        =   392
         Top             =   440
         Width           =   975
         Begin VB.TextBox tbB5CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   404
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB5CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   403
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB6CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   402
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB6CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   401
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB6CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   400
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB6CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   399
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB6CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   398
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB6CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   397
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB6CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   396
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB6CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   395
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB6CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   394
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB6CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   393
            Top             =   5200
            Width           =   735
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "通道设置（三）"
         Height          =   5655
         Left            =   -68160
         TabIndex        =   352
         Top             =   440
         Width           =   3735
         Begin VB.ComboBox cbbB5CH 
            Height          =   315
            Index           =   7
            Left            =   1080
            TabIndex        =   378
            Top             =   5200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB5CH 
            Height          =   315
            Index           =   6
            Left            =   1080
            TabIndex        =   377
            Top             =   4800
            Width           =   2535
         End
         Begin VB.CheckBox cbB5CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   376
            Top             =   5240
            Width           =   255
         End
         Begin VB.CheckBox cbB5CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   375
            Top             =   4840
            Width           =   255
         End
         Begin VB.ComboBox cbbB5CH 
            Height          =   315
            Index           =   5
            Left            =   1080
            TabIndex        =   374
            Top             =   4400
            Width           =   2535
         End
         Begin VB.ComboBox cbbB5CH 
            Height          =   315
            Index           =   4
            Left            =   1080
            TabIndex        =   373
            Top             =   4000
            Width           =   2535
         End
         Begin VB.ComboBox cbbB5CH 
            Height          =   315
            Index           =   3
            Left            =   1080
            TabIndex        =   372
            Top             =   3600
            Width           =   2535
         End
         Begin VB.ComboBox cbbB5CH 
            Height          =   315
            Index           =   2
            Left            =   1080
            TabIndex        =   371
            Top             =   3200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB5CH 
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   370
            Top             =   2800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB5CH 
            Height          =   315
            Index           =   0
            Left            =   1080
            TabIndex        =   369
            Top             =   2400
            Width           =   2535
         End
         Begin VB.ComboBox cbbB4CH 
            Height          =   315
            Index           =   9
            Left            =   1080
            TabIndex        =   368
            Top             =   2000
            Width           =   2535
         End
         Begin VB.ComboBox cbbB4CH 
            Height          =   315
            Index           =   8
            Left            =   1080
            TabIndex        =   367
            Top             =   1600
            Width           =   2535
         End
         Begin VB.ComboBox cbbB4CH 
            Height          =   315
            Index           =   7
            Left            =   1080
            TabIndex        =   366
            Top             =   1200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB4CH 
            Height          =   315
            Index           =   6
            Left            =   1080
            TabIndex        =   365
            Top             =   800
            Width           =   2535
         End
         Begin VB.ComboBox cbbCHSec4 
            Height          =   315
            ItemData        =   "frmMain.frx":03AE
            Left            =   1080
            List            =   "frmMain.frx":03B0
            TabIndex        =   364
            Top             =   300
            Width           =   2535
         End
         Begin VB.CheckBox cbB5CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   363
            Top             =   4440
            Width           =   255
         End
         Begin VB.CheckBox cbB5CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   362
            Top             =   4040
            Width           =   255
         End
         Begin VB.CheckBox cbB5CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   361
            Top             =   3640
            Width           =   255
         End
         Begin VB.CheckBox cbB5CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   360
            Top             =   3240
            Width           =   255
         End
         Begin VB.CheckBox cbB5CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   359
            Top             =   2840
            Width           =   255
         End
         Begin VB.CheckBox cbB5CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   358
            Top             =   2440
            Width           =   255
         End
         Begin VB.CheckBox cbB4CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   357
            Top             =   2040
            Width           =   255
         End
         Begin VB.CheckBox cbB4CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   356
            Top             =   1640
            Width           =   255
         End
         Begin VB.CheckBox cbB4CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   355
            Top             =   1240
            Width           =   255
         End
         Begin VB.CheckBox cbB4CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   354
            Top             =   840
            Width           =   255
         End
         Begin VB.CheckBox cbCHSec4 
            Caption         =   "Check13"
            Height          =   255
            Left            =   120
            TabIndex        =   353
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lblCH 
            Caption         =   "CH48"
            Height          =   255
            Index           =   47
            Left            =   360
            TabIndex        =   391
            Top             =   5270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH47"
            Height          =   255
            Index           =   46
            Left            =   360
            TabIndex        =   390
            Top             =   4870
            Width           =   735
         End
         Begin VB.Line Line8 
            X1              =   0
            X2              =   3720
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Label lblCH 
            Caption         =   "CH46"
            Height          =   255
            Index           =   45
            Left            =   360
            TabIndex        =   389
            Top             =   4470
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH45"
            Height          =   255
            Index           =   44
            Left            =   360
            TabIndex        =   388
            Top             =   4070
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH44"
            Height          =   255
            Index           =   43
            Left            =   360
            TabIndex        =   387
            Top             =   3670
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH43"
            Height          =   255
            Index           =   42
            Left            =   360
            TabIndex        =   386
            Top             =   3270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH42"
            Height          =   255
            Index           =   41
            Left            =   360
            TabIndex        =   385
            Top             =   2870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH41"
            Height          =   255
            Index           =   40
            Left            =   360
            TabIndex        =   384
            Top             =   2470
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH40"
            Height          =   255
            Index           =   39
            Left            =   360
            TabIndex        =   383
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH39"
            Height          =   255
            Index           =   38
            Left            =   360
            TabIndex        =   382
            Top             =   1670
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH38"
            Height          =   255
            Index           =   37
            Left            =   360
            TabIndex        =   381
            Top             =   1270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH37"
            Height          =   255
            Index           =   36
            Left            =   360
            TabIndex        =   380
            Top             =   870
            Width           =   735
         End
         Begin VB.Label lblCHSec4 
            Caption         =   "全/反选"
            Height          =   255
            Left            =   375
            TabIndex        =   379
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "传感器读数"
         Height          =   5655
         Left            =   -64440
         TabIndex        =   325
         Top             =   440
         Width           =   1935
         Begin VB.TextBox tbB5CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   349
            Top             =   5200
            Width           =   735
         End
         Begin VB.TextBox tbB5CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   348
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB5CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   347
            Top             =   5200
            Width           =   735
         End
         Begin VB.TextBox tbB5CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   346
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB5CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   345
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB5CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   344
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB5CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   343
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB5CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   342
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB5CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   341
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB5CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   340
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB5CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   339
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB5CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   338
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB5CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   337
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB5CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   336
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB5CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   335
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB5CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   334
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB4CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   1080
            TabIndex        =   333
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB4CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   332
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB4CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   1080
            TabIndex        =   331
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB4CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   330
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB4CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   329
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB4CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   328
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB4CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   327
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB4CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   326
            Top             =   800
            Width           =   735
         End
         Begin VB.Label Label74 
            Caption         =   "高温"
            Height          =   255
            Left            =   1270
            TabIndex        =   351
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label73 
            Caption         =   "低温"
            Height          =   255
            Left            =   320
            TabIndex        =   350
            Top             =   360
            Width           =   495
         End
         Begin VB.Line Line7 
            X1              =   960
            X2              =   960
            Y1              =   240
            Y2              =   5640
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "结果 L"
         Height          =   5655
         Left            =   -62520
         TabIndex        =   312
         Top             =   440
         Width           =   975
         Begin VB.TextBox tbB5CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   324
            Top             =   5200
            Width           =   735
         End
         Begin VB.TextBox tbB5CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   323
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB5CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   322
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB5CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   321
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB5CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   320
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB5CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   319
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB5CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   318
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB5CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   317
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB4CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   316
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB4CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   315
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB4CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   314
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB4CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   313
            Top             =   800
            Width           =   735
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "通道设置（三）"
         Height          =   5655
         Left            =   -74880
         TabIndex        =   272
         Top             =   440
         Width           =   3735
         Begin VB.CheckBox cbCHSec3 
            Caption         =   "Check13"
            Height          =   255
            Left            =   120
            TabIndex        =   298
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox cbB3CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   297
            Top             =   840
            Width           =   255
         End
         Begin VB.CheckBox cbB3CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   296
            Top             =   1240
            Width           =   255
         End
         Begin VB.CheckBox cbB3CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   295
            Top             =   1640
            Width           =   255
         End
         Begin VB.CheckBox cbB3CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   294
            Top             =   2040
            Width           =   255
         End
         Begin VB.CheckBox cbB3CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   293
            Top             =   2440
            Width           =   255
         End
         Begin VB.CheckBox cbB3CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   292
            Top             =   2840
            Width           =   255
         End
         Begin VB.CheckBox cbB4CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   291
            Top             =   3240
            Width           =   255
         End
         Begin VB.CheckBox cbB4CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   290
            Top             =   3640
            Width           =   255
         End
         Begin VB.CheckBox cbB4CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   289
            Top             =   4040
            Width           =   255
         End
         Begin VB.CheckBox cbB4CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   288
            Top             =   4440
            Width           =   255
         End
         Begin VB.ComboBox cbbCHSec3 
            Height          =   315
            ItemData        =   "frmMain.frx":03B2
            Left            =   1080
            List            =   "frmMain.frx":03B4
            TabIndex        =   287
            Top             =   300
            Width           =   2535
         End
         Begin VB.ComboBox cbbB3CH 
            Height          =   315
            Index           =   4
            Left            =   1080
            TabIndex        =   286
            Top             =   800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB3CH 
            Height          =   315
            Index           =   5
            Left            =   1080
            TabIndex        =   285
            Top             =   1200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB3CH 
            Height          =   315
            Index           =   6
            Left            =   1080
            TabIndex        =   284
            Top             =   1600
            Width           =   2535
         End
         Begin VB.ComboBox cbbB3CH 
            Height          =   315
            Index           =   7
            Left            =   1080
            TabIndex        =   283
            Top             =   2000
            Width           =   2535
         End
         Begin VB.ComboBox cbbB3CH 
            Height          =   315
            Index           =   8
            Left            =   1080
            TabIndex        =   282
            Top             =   2400
            Width           =   2535
         End
         Begin VB.ComboBox cbbB3CH 
            Height          =   315
            Index           =   9
            Left            =   1080
            TabIndex        =   281
            Top             =   2800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB4CH 
            Height          =   315
            Index           =   0
            Left            =   1080
            TabIndex        =   280
            Top             =   3200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB4CH 
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   279
            Top             =   3600
            Width           =   2535
         End
         Begin VB.ComboBox cbbB4CH 
            Height          =   315
            Index           =   2
            Left            =   1080
            TabIndex        =   278
            Top             =   4000
            Width           =   2535
         End
         Begin VB.ComboBox cbbB4CH 
            Height          =   315
            Index           =   3
            Left            =   1080
            TabIndex        =   277
            Top             =   4400
            Width           =   2535
         End
         Begin VB.CheckBox cbB4CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   276
            Top             =   4840
            Width           =   255
         End
         Begin VB.CheckBox cbB4CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   275
            Top             =   5240
            Width           =   255
         End
         Begin VB.ComboBox cbbB4CH 
            Height          =   315
            Index           =   4
            Left            =   1080
            TabIndex        =   274
            Top             =   4800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB4CH 
            Height          =   315
            Index           =   5
            Left            =   1080
            TabIndex        =   273
            Top             =   5200
            Width           =   2535
         End
         Begin VB.Label lblCHSec3 
            Caption         =   "全/反选"
            Height          =   255
            Left            =   375
            TabIndex        =   311
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblCH 
            Caption         =   "CH25"
            Height          =   255
            Index           =   24
            Left            =   360
            TabIndex        =   310
            Top             =   870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH26"
            Height          =   255
            Index           =   25
            Left            =   360
            TabIndex        =   309
            Top             =   1270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH27"
            Height          =   255
            Index           =   26
            Left            =   360
            TabIndex        =   308
            Top             =   1670
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH28"
            Height          =   255
            Index           =   27
            Left            =   360
            TabIndex        =   307
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH29"
            Height          =   255
            Index           =   28
            Left            =   360
            TabIndex        =   306
            Top             =   2470
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH30"
            Height          =   255
            Index           =   29
            Left            =   360
            TabIndex        =   305
            Top             =   2870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH31"
            Height          =   255
            Index           =   30
            Left            =   360
            TabIndex        =   304
            Top             =   3270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH32"
            Height          =   255
            Index           =   31
            Left            =   360
            TabIndex        =   303
            Top             =   3670
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH33"
            Height          =   255
            Index           =   32
            Left            =   360
            TabIndex        =   302
            Top             =   4070
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH34"
            Height          =   255
            Index           =   33
            Left            =   360
            TabIndex        =   301
            Top             =   4470
            Width           =   735
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   3720
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Label lblCH 
            Caption         =   "CH35"
            Height          =   255
            Index           =   34
            Left            =   360
            TabIndex        =   300
            Top             =   4870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH36"
            Height          =   255
            Index           =   35
            Left            =   360
            TabIndex        =   299
            Top             =   5270
            Width           =   735
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "传感器读数"
         Height          =   5655
         Left            =   -71160
         TabIndex        =   245
         Top             =   440
         Width           =   1935
         Begin VB.TextBox tbB3CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   269
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB3CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   268
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB3CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   267
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB3CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   266
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB3CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   265
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB3CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   264
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB3CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   263
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB3CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   262
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB3CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   261
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB3CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   1080
            TabIndex        =   260
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB3CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   259
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB3CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   1080
            TabIndex        =   258
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB4CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   257
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB4CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   256
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB4CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   255
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB4CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   254
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB4CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   253
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB4CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   252
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB4CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   251
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB4CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   250
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB4CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   249
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB4CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   248
            Top             =   5200
            Width           =   735
         End
         Begin VB.TextBox tbB4CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   247
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB4CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   246
            Top             =   5200
            Width           =   735
         End
         Begin VB.Line Line5 
            X1              =   960
            X2              =   960
            Y1              =   240
            Y2              =   5640
         End
         Begin VB.Label Label60 
            Caption         =   "低温"
            Height          =   255
            Left            =   320
            TabIndex        =   271
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label30 
            Caption         =   "高温"
            Height          =   255
            Left            =   1270
            TabIndex        =   270
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "结果 L"
         Height          =   5655
         Left            =   -69240
         TabIndex        =   232
         Top             =   440
         Width           =   975
         Begin VB.TextBox tbB3CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   244
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB3CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   243
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB3CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   242
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB3CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   241
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB3CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   240
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB3CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   239
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB4CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   238
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB4CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   237
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB4CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   236
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB4CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   235
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB4CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   234
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB4CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   233
            Top             =   5200
            Width           =   735
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "通道设置（二）"
         Height          =   5655
         Left            =   -68160
         TabIndex        =   202
         Top             =   440
         Width           =   3735
         Begin VB.CheckBox cbCHSec2 
            Caption         =   "Check13"
            Height          =   255
            Left            =   120
            TabIndex        =   218
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox cbB2CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   217
            Top             =   840
            Width           =   255
         End
         Begin VB.CheckBox cbB2CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   216
            Top             =   1240
            Width           =   255
         End
         Begin VB.CheckBox cbB2CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   215
            Top             =   1640
            Width           =   255
         End
         Begin VB.CheckBox cbB2CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   214
            Top             =   2040
            Width           =   255
         End
         Begin VB.CheckBox cbB2CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   213
            Top             =   2440
            Width           =   255
         End
         Begin VB.CheckBox cbB2CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   212
            Top             =   2840
            Width           =   255
         End
         Begin VB.CheckBox cbB2CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   211
            Top             =   3240
            Width           =   255
         End
         Begin VB.CheckBox cbB2CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   210
            Top             =   3640
            Width           =   255
         End
         Begin VB.CheckBox cbB3CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   209
            Top             =   4040
            Width           =   255
         End
         Begin VB.CheckBox cbB3CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   208
            Top             =   4440
            Width           =   255
         End
         Begin VB.ComboBox cbbCHSec2 
            Height          =   315
            ItemData        =   "frmMain.frx":03B6
            Left            =   1080
            List            =   "frmMain.frx":03B8
            TabIndex        =   207
            Top             =   300
            Width           =   2535
         End
         Begin VB.ComboBox cbbB2CH 
            Height          =   315
            Index           =   2
            Left            =   1080
            TabIndex        =   206
            Top             =   800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB2CH 
            Height          =   315
            Index           =   3
            Left            =   1080
            TabIndex        =   205
            Top             =   1200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB2CH 
            Height          =   315
            Index           =   4
            Left            =   1080
            TabIndex        =   204
            Top             =   1600
            Width           =   2535
         End
         Begin VB.ComboBox cbbB2CH 
            Height          =   315
            Index           =   5
            Left            =   1080
            TabIndex        =   203
            Top             =   2000
            Width           =   2535
         End
         Begin VB.ComboBox cbbB2CH 
            Height          =   315
            Index           =   6
            Left            =   1080
            TabIndex        =   201
            Top             =   2400
            Width           =   2535
         End
         Begin VB.ComboBox cbbB2CH 
            Height          =   315
            Index           =   7
            Left            =   1080
            TabIndex        =   200
            Top             =   2800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB2CH 
            Height          =   315
            Index           =   8
            Left            =   1080
            TabIndex        =   199
            Top             =   3200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB2CH 
            Height          =   315
            Index           =   9
            Left            =   1080
            TabIndex        =   198
            Top             =   3600
            Width           =   2535
         End
         Begin VB.ComboBox cbbB3CH 
            Height          =   315
            Index           =   0
            Left            =   1080
            TabIndex        =   197
            Top             =   4000
            Width           =   2535
         End
         Begin VB.ComboBox cbbB3CH 
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   196
            Top             =   4400
            Width           =   2535
         End
         Begin VB.CheckBox cbB3CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   195
            Top             =   4840
            Width           =   255
         End
         Begin VB.CheckBox cbB3CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   194
            Top             =   5240
            Width           =   255
         End
         Begin VB.ComboBox cbbB3CH 
            Height          =   315
            Index           =   2
            Left            =   1080
            TabIndex        =   193
            Top             =   4800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB3CH 
            Height          =   315
            Index           =   3
            Left            =   1080
            TabIndex        =   192
            Top             =   5200
            Width           =   2535
         End
         Begin VB.Label lblCHSec2 
            Caption         =   "全/反选"
            Height          =   255
            Left            =   375
            TabIndex        =   231
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblCH 
            Caption         =   "CH20"
            Height          =   255
            Index           =   19
            Left            =   360
            TabIndex        =   230
            Top             =   3670
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH19"
            Height          =   255
            Index           =   18
            Left            =   360
            TabIndex        =   229
            Top             =   3270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH18"
            Height          =   255
            Index           =   17
            Left            =   360
            TabIndex        =   228
            Top             =   2870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH17"
            Height          =   255
            Index           =   16
            Left            =   360
            TabIndex        =   227
            Top             =   2470
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH16"
            Height          =   255
            Index           =   15
            Left            =   360
            TabIndex        =   226
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH15"
            Height          =   255
            Index           =   14
            Left            =   360
            TabIndex        =   225
            Top             =   1670
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH14"
            Height          =   255
            Index           =   13
            Left            =   360
            TabIndex        =   224
            Top             =   1270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH13"
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   223
            Top             =   870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH21"
            Height          =   255
            Index           =   20
            Left            =   360
            TabIndex        =   222
            Top             =   4070
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH22"
            Height          =   255
            Index           =   21
            Left            =   360
            TabIndex        =   221
            Top             =   4470
            Width           =   735
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   3720
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Label lblCH 
            Caption         =   "CH23"
            Height          =   255
            Index           =   22
            Left            =   360
            TabIndex        =   220
            Top             =   4870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH24"
            Height          =   255
            Index           =   23
            Left            =   360
            TabIndex        =   219
            Top             =   5270
            Width           =   735
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "传感器读数"
         Height          =   5655
         Left            =   -64440
         TabIndex        =   166
         Top             =   440
         Width           =   1935
         Begin VB.TextBox tbB2CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   189
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB2CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   188
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB2CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   187
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB2CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   186
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB2CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   185
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB2CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   184
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB2CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   183
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB2CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   182
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB2CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   181
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB2CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   180
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB2CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   179
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB2CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   178
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB2CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   177
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB2CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   1080
            TabIndex        =   176
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB2CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   175
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB2CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   1080
            TabIndex        =   174
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB3CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   173
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB3CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   172
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB3CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   171
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB3CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   170
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB3CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   169
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB3CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   168
            Top             =   5200
            Width           =   735
         End
         Begin VB.TextBox tbB3CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   167
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB3CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   165
            Top             =   5200
            Width           =   735
         End
         Begin VB.Line Line3 
            X1              =   960
            X2              =   960
            Y1              =   240
            Y2              =   5640
         End
         Begin VB.Label Label47 
            Caption         =   "低温"
            Height          =   255
            Left            =   320
            TabIndex        =   191
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label46 
            Caption         =   "高温"
            Height          =   255
            Left            =   1270
            TabIndex        =   190
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "结果 L"
         Height          =   5655
         Left            =   -62520
         TabIndex        =   152
         Top             =   440
         Width           =   975
         Begin VB.TextBox tbB2CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   164
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB2CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   163
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB2CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   162
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB2CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   161
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB2CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   160
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB2CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   159
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB2CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   158
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB2CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   157
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB3CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   156
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB3CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   155
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB3CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   154
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB3CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   153
            Top             =   5200
            Width           =   735
         End
      End
      Begin VB.TextBox tb1680ComReading 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67080
         TabIndex        =   132
         Top             =   6220
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "测试连接"
         Height          =   255
         Left            =   -66240
         TabIndex        =   131
         Top             =   6220
         Width           =   975
      End
      Begin VB.ComboBox cbbBTest 
         Height          =   315
         ItemData        =   "frmMain.frx":03BA
         Left            =   -70920
         List            =   "frmMain.frx":03BC
         TabIndex        =   129
         Top             =   6200
         Width           =   615
      End
      Begin VB.TextBox tb1680ComInput 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70200
         TabIndex        =   128
         Top             =   6220
         Width           =   3015
      End
      Begin VB.Frame Frame7 
         Caption         =   "结果 L"
         Height          =   5655
         Left            =   -69240
         TabIndex        =   92
         Top             =   440
         Width           =   975
         Begin VB.TextBox tbB2CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   151
            Top             =   5200
            Width           =   735
         End
         Begin VB.TextBox tbB2CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   150
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB1CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   121
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB1CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   120
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB1CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   119
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB1CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   118
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB1CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   117
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB1CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   116
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB1CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   115
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB1CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   114
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB1CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   113
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB1CHL 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   94
            Top             =   800
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "传感器读数"
         Height          =   5655
         Left            =   -71160
         TabIndex        =   88
         Top             =   440
         Width           =   1935
         Begin VB.TextBox tbB2CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   149
            Top             =   5200
            Width           =   735
         End
         Begin VB.TextBox tbB2CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   148
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB2CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   147
            Top             =   5200
            Width           =   735
         End
         Begin VB.TextBox tbB2CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   146
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox tbB1CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   1080
            TabIndex        =   112
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB1CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   111
            Top             =   4400
            Width           =   735
         End
         Begin VB.TextBox tbB1CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   1080
            TabIndex        =   110
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB1CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   109
            Top             =   4000
            Width           =   735
         End
         Begin VB.TextBox tbB1CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   108
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB1CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   107
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox tbB1CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   106
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB1CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   105
            Top             =   3200
            Width           =   735
         End
         Begin VB.TextBox tbB1CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   104
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB1CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   103
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox tbB1CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   102
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB1CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   101
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox tbB1CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   100
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB1CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   99
            Top             =   2000
            Width           =   735
         End
         Begin VB.TextBox tbB1CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   98
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB1CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   97
            Top             =   1600
            Width           =   735
         End
         Begin VB.TextBox tbB1CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   96
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB1CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   95
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox tbB1CHHT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   93
            Top             =   800
            Width           =   735
         End
         Begin VB.TextBox tbB1CHLT 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   91
            Top             =   800
            Width           =   735
         End
         Begin VB.Label Label42 
            Caption         =   "高温"
            Height          =   255
            Left            =   1270
            TabIndex        =   90
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label41 
            Caption         =   "低温"
            Height          =   255
            Left            =   320
            TabIndex        =   89
            Top             =   360
            Width           =   495
         End
         Begin VB.Line Line2 
            X1              =   960
            X2              =   960
            Y1              =   240
            Y2              =   5640
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "通道设置（一）"
         Height          =   5655
         Left            =   -74880
         TabIndex        =   54
         Top             =   440
         Width           =   3735
         Begin VB.ComboBox cbbB2CH 
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   145
            Top             =   5200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB2CH 
            Height          =   315
            Index           =   0
            Left            =   1080
            TabIndex        =   144
            Top             =   4800
            Width           =   2535
         End
         Begin VB.CheckBox cbB2CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   141
            Top             =   5240
            Width           =   255
         End
         Begin VB.CheckBox cbB2CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   140
            Top             =   4840
            Width           =   255
         End
         Begin VB.ComboBox cbbB1CH 
            Height          =   315
            Index           =   9
            Left            =   1080
            TabIndex        =   87
            Top             =   4400
            Width           =   2535
         End
         Begin VB.ComboBox cbbB1CH 
            Height          =   315
            Index           =   8
            Left            =   1080
            TabIndex        =   86
            Top             =   4000
            Width           =   2535
         End
         Begin VB.ComboBox cbbB1CH 
            Height          =   315
            Index           =   7
            Left            =   1080
            TabIndex        =   85
            Top             =   3600
            Width           =   2535
         End
         Begin VB.ComboBox cbbB1CH 
            Height          =   315
            Index           =   6
            Left            =   1080
            TabIndex        =   84
            Top             =   3200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB1CH 
            Height          =   315
            Index           =   5
            Left            =   1080
            TabIndex        =   83
            Top             =   2800
            Width           =   2535
         End
         Begin VB.ComboBox cbbB1CH 
            Height          =   315
            Index           =   4
            Left            =   1080
            TabIndex        =   82
            Top             =   2400
            Width           =   2535
         End
         Begin VB.ComboBox cbbB1CH 
            Height          =   315
            Index           =   3
            Left            =   1080
            TabIndex        =   81
            Top             =   2000
            Width           =   2535
         End
         Begin VB.ComboBox cbbB1CH 
            Height          =   315
            Index           =   2
            Left            =   1080
            TabIndex        =   80
            Top             =   1600
            Width           =   2535
         End
         Begin VB.ComboBox cbbB1CH 
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   79
            Top             =   1200
            Width           =   2535
         End
         Begin VB.ComboBox cbbB1CH 
            Height          =   315
            Index           =   0
            Left            =   1080
            TabIndex        =   78
            Top             =   800
            Width           =   2535
         End
         Begin VB.ComboBox cbbCHSec1 
            Height          =   315
            ItemData        =   "frmMain.frx":03BE
            Left            =   1080
            List            =   "frmMain.frx":03C0
            TabIndex        =   77
            Top             =   300
            Width           =   2535
         End
         Begin VB.CheckBox cbB1CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   75
            Top             =   4440
            Width           =   255
         End
         Begin VB.CheckBox cbB1CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   73
            Top             =   4040
            Width           =   255
         End
         Begin VB.CheckBox cbB1CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   71
            Top             =   3640
            Width           =   255
         End
         Begin VB.CheckBox cbB1CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   69
            Top             =   3240
            Width           =   255
         End
         Begin VB.CheckBox cbB1CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   67
            Top             =   2840
            Width           =   255
         End
         Begin VB.CheckBox cbB1CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   65
            Top             =   2440
            Width           =   255
         End
         Begin VB.CheckBox cbB1CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   63
            Top             =   2040
            Width           =   255
         End
         Begin VB.CheckBox cbB1CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   61
            Top             =   1640
            Width           =   255
         End
         Begin VB.CheckBox cbB1CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Top             =   1240
            Width           =   255
         End
         Begin VB.CheckBox cbB1CH 
            Caption         =   "Check13"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   255
         End
         Begin VB.CheckBox cbCHSec1 
            Caption         =   "Check13"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lblCH 
            Caption         =   "CH12"
            Height          =   255
            Index           =   11
            Left            =   360
            TabIndex        =   143
            Top             =   5270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH11"
            Height          =   255
            Index           =   10
            Left            =   360
            TabIndex        =   142
            Top             =   4870
            Width           =   735
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   3720
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Label lblCH 
            Caption         =   "CH10"
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   76
            Top             =   4470
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH9"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   74
            Top             =   4070
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH8"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   72
            Top             =   3670
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH7"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   70
            Top             =   3270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH6"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   68
            Top             =   2870
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH5"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   66
            Top             =   2470
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH4"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   64
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH3"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   62
            Top             =   1670
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH2"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   60
            Top             =   1270
            Width           =   735
         End
         Begin VB.Label lblCH 
            Caption         =   "CH1"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   58
            Top             =   870
            Width           =   735
         End
         Begin VB.Label lblCHSec1 
            Caption         =   "全/反选"
            Height          =   255
            Left            =   375
            TabIndex        =   56
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "温度设置"
         Height          =   2055
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   3855
         Begin VB.CommandButton btnTempApply 
            Caption         =   "应用"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2280
            TabIndex        =   122
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox tbHighTemp 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   53
            Text            =   "0"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox tbLowTemp 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   46
            Text            =   "0"
            Top             =   210
            Width           =   1335
         End
         Begin VB.TextBox tbLTTime 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   45
            Text            =   "0"
            Top             =   570
            Width           =   1335
         End
         Begin VB.TextBox tbHTTime 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   44
            Text            =   "0"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton btnSetTemp 
            Caption         =   "修改温度设置"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label44 
            Caption         =   "分钟"
            Height          =   255
            Left            =   3360
            TabIndex        =   124
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label43 
            Caption         =   "℃"
            Height          =   255
            Left            =   3360
            TabIndex        =   123
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "最低温度："
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "最低温度保持时间："
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "℃"
            Height          =   255
            Left            =   3360
            TabIndex        =   50
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "分钟"
            Height          =   255
            Left            =   3360
            TabIndex        =   49
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label5 
            Caption         =   "最高温度："
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "最高温度保持时间："
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "状态监控"
         Height          =   1335
         Left            =   120
         TabIndex        =   38
         Top             =   2400
         Width           =   3855
         Begin VB.TextBox tbTempComInput 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   127
            Top             =   600
            Width           =   3615
         End
         Begin VB.CommandButton Command1 
            Caption         =   "测试连接"
            Height          =   255
            Left            =   2280
            TabIndex        =   126
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox tbCurrTemp 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   39
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "烘箱当前温度："
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "℃"
            Height          =   255
            Left            =   3360
            TabIndex        =   40
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "测试步骤"
         Height          =   6255
         Left            =   4080
         TabIndex        =   15
         Top             =   360
         Width           =   9375
         Begin VB.TextBox tbHTReach 
            Height          =   285
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   612
            Top             =   3760
            Width           =   1335
         End
         Begin VB.TextBox tbLTReach 
            Height          =   285
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   611
            Top             =   2320
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "重新检测"
            Height          =   375
            Left            =   8280
            TabIndex        =   139
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox cbP1 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox cbP2 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   255
         End
         Begin VB.CheckBox cbP3 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1920
            Width           =   255
         End
         Begin VB.CheckBox cbP4 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   2400
            Width           =   255
         End
         Begin VB.CheckBox cbP5 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   2880
            Width           =   255
         End
         Begin VB.CheckBox cbP6 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   3360
            Width           =   255
         End
         Begin VB.CheckBox cbP7 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   3840
            Width           =   255
         End
         Begin VB.CheckBox cbP8 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   4320
            Width           =   255
         End
         Begin VB.CheckBox cbP9 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   4800
            Width           =   255
         End
         Begin VB.CheckBox cbP10 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   5280
            Width           =   255
         End
         Begin VB.CheckBox cbP11 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   5760
            Width           =   255
         End
         Begin VB.Label lblB6 
            Caption         =   "主板6"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7080
            TabIndex        =   138
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblB5 
            Caption         =   "主板5"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5760
            TabIndex        =   137
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblB4 
            Caption         =   "主板4"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4440
            TabIndex        =   136
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblB3 
            Caption         =   "主板3"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   135
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblB2 
            Caption         =   "主板2"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   134
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblB1 
            Caption         =   "主板1"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   133
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "程序自检温度、时间和型号、系数等设定"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   37
            Top             =   360
            Width           =   8775
         End
         Begin VB.Label Label14 
            Caption         =   "程序自检主板和连接状态"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   36
            Top             =   840
            Width           =   8775
         End
         Begin VB.Label Label15 
            Caption         =   "程序自检各通道的设置状况"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   35
            Top             =   1800
            Width           =   8775
         End
         Begin VB.Label Label16 
            Caption         =   "等待烘箱到达最低温度设置.........."
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   34
            Top             =   2280
            Width           =   8775
         End
         Begin VB.Label Label17 
            Caption         =   "保持最低温度相应时间"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   33
            Top             =   2760
            Width           =   8775
         End
         Begin VB.Label Label18 
            Caption         =   "读取最低温度时的传感器读数"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   32
            Top             =   3240
            Width           =   8775
         End
         Begin VB.Label Label20 
            Caption         =   "等待烘箱到达最高温度设置.........."
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   31
            Top             =   3720
            Width           =   8775
         End
         Begin VB.Label Label21 
            Caption         =   "保持最高温度相应时间"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   30
            Top             =   4200
            Width           =   8775
         End
         Begin VB.Label Label22 
            Caption         =   "读取最高温度时的传感器读数"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   29
            Top             =   4680
            Width           =   8775
         End
         Begin VB.Label Label23 
            Caption         =   "依据公式计算L值"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   28
            Top             =   5160
            Width           =   8775
         End
         Begin VB.Label Label24 
            Caption         =   "本次测试结束，请保存结果并打印"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   20.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   27
            Top             =   5640
            Width           =   8775
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "测试时间"
         Height          =   2055
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   3855
         Begin VB.TextBox tbStartTime 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox tbEndTime 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox tbTestTime 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label25 
            Caption         =   "本次检测开始时间："
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "本次检测结束时间："
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label27 
            Caption         =   "本次检测已经过时间："
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label28 
            Caption         =   "分钟"
            Height          =   255
            Left            =   1320
            TabIndex        =   11
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "已暂停！"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2040
            TabIndex        =   10
            Top             =   1680
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Label Label11 
         Caption         =   "CH"
         Height          =   255
         Left            =   -71160
         TabIndex        =   130
         Top             =   6265
         Width           =   255
      End
   End
   Begin MSCommLib.MSComm comTemp 
      Left            =   8400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm com1680 
      Left            =   9000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblSaveSuccess 
      Height          =   255
      Left            =   8760
      TabIndex        =   606
      Top             =   7440
      Width           =   4695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' 自其他两个表单接受
Public hasTempAuth As Boolean
Public hasCOMAuth As Boolean

Private timerTempTickCount As Integer
Private timerTestTickCount As Integer
Private timerInterval As Double

Private modelName() As String
Private modelResist() As Integer
Private modelCoeff() As Double

Private autoSave As Boolean
Private autoSaveFN As String

Private B1Installed As Boolean
Private B2Installed As Boolean
Private B3Installed As Boolean
Private B4Installed As Boolean
Private B5Installed As Boolean
Private B6Installed As Boolean

Private chTest As Integer

Private TestDone As Boolean

Private Paused As Boolean

Private timeout As Boolean
Private timeoutTime As Integer

Private Sub btnCHTestDown_Click()
    If chTest = 60 Then
        chTest = 1
    Else
        chTest = chTest + 1
    End If
End Sub

Private Sub btnCHTestUp_Click()
    If chTest = 1 Then
        chTest = 60
    Else
        chTest = chTest - 1
    End If
End Sub

Private Sub btnPause_Click()
    If Paused = False Then
        btnPause.Caption = "继续测试"
        Paused = True
        timerTemp.Enabled = False
    Else
        btnPause.Caption = "暂停测试"
        Paused = False
        timerTemp.Enabled = True
    End If
End Sub

Private Sub btnSetCOM_Click()
    Dim comform As New frmCOMPass
    comform.Show
    Do While comform.closed = False
        DoEvents
    Loop
    If hasCOMAuth = True Then
        tbTempCOM.Enabled = True
        tb1680COM.Enabled = True
        btnCOMApply.Enabled = True
        btnSetCOM.Enabled = False
    End If
End Sub

Private Sub btnSetTemp_Click()
    Dim passform As New frmTempPass
    passform.Show
    Do While passform.closed = False
        DoEvents
    Loop
    If hasTempAuth = True Then
        tbLowTemp.Enabled = True
        tbLTTime.Enabled = True
        tbHighTemp.Enabled = True
        tbHTTime.Enabled = True
        btnTempApply.Enabled = True
        btnSetTemp.Enabled = False
    End If
End Sub

Private Sub btnStart_Click()
    cbP3.Value = 0
    cbP4.Value = 0
    cbP5.Value = 0
    cbP6.Value = 0
    cbP7.Value = 0
    cbP8.Value = 0
    cbP9.Value = 0
    cbP10.Value = 0
    cbP11.Value = 0
    
    tbLTReach.Text = ""
    tbHTReach.Text = ""

    If tbLowTemp.Enabled = True Then
        MsgBox "请首先点击应用键以确认温度设置！"
    ElseIf tbTempCOM.Enabled = True Then
        MsgBox "请首先点击应用键以确认端口设置！"
    ElseIf ValidateCHSettings = False Then
        MsgBox "请修复上述错误后再点击开始测试"
    Else
        cbP3.Value = 1
        
        If MsgBox("是否需要于测试结束时自动保存？", vbYesNo) = vbYes Then
            autoSave = True
            cdlg.CancelError = True
            cdlg.DialogTitle = "保存测试数据"
            cdlg.Filter = "Word file (*.doc)|*.doc"
            cdlg.FileName = Replace$(DateValue(Now) & " " & TimeValue(Now), ":", "")
            cdlg.ShowSave
        End If
        
        tbStartTime.Text = DateValue(Now) & " " & TimeValue(Now)
        btnStart.Enabled = False
        btnStop.Enabled = True
        timerOvenTemp.Enabled = True
        timerTest.Enabled = True
        If SetTempUart(CInt(tbLowTemp)) Then
            Do Until CInt(tbCurrTemp.Text) = CInt(tbLowTemp.Text)
                DoEvents
            Loop
            cbP4.Value = 1
            tbLTReach.Text = TimeValue(Now)
            
            btnPause.Enabled = True
            CountTime CInt(tbLTTime.Text)
            cbP5.Value = 1
            btnPause.Enabled = False
            
            ReadSensorData "LT"
            cbP6.Value = 1
            
            If SetTempUart(CInt(tbHighTemp)) Then
                Do Until CInt(tbCurrTemp.Text) = CInt(tbHighTemp.Text)
                    DoEvents
                Loop
                cbP7.Value = 1
                tbHTReach.Text = TimeValue(Now)
                
                btnPause.Enabled = True
                CountTime CInt(tbHTTime.Text)
                cbP8.Value = 1
                btnPause.Enabled = False
                
                ReadSensorData "HT"
                cbP9.Value = 1
                
                ComputeL
                cbP10.Value = 1
            End If
        End If
        tbEndTime = DateValue(Now) & " " & TimeValue(Now)
        If SetTempUart(CInt(tbLowTemp)) Then
            tbTempComInput.Text = "本次测试结束，回到最低温度设置！"
            cbP11.Value = 1
        End If
        btnStart.Enabled = True
        btnPause.Enabled = False
        btnStop.Enabled = False
        timerOvenTemp.Enabled = False
        timerTest.Enabled = False
        timerTestTickCount = 0
        
        ' indicate at least one test is done, the result can be saved
        TestDone = True
        ' auto save the result if the user wants to
        If autoSave = True Then
            SaveResultWord cdlg.FileName
            SaveResultSQL
            autoSave = False
        End If
    End If
End Sub

Private Function ValidateCHSettings() As Boolean
    Dim errMsg As String
    errMsg = "通道设置的错误：" & vbCrLf & vbCrLf
    If B1Installed = True Then
        For i = 0 To 9
            If cbB1CH(i).Value = 1 And StrComp(cbbB1CH(i).Text, "") = 0 Then
                errMsg = errMsg & "    通道" & (i + 1) & "开启但未选择传感器型号！" & vbCrLf
            End If
        Next i
    End If
    
    If StrComp(errMsg, "通道设置的错误：" & vbCrLf & vbCrLf) = 0 Then
        ValidateCHSettings = True
    Else
        MsgBox errMsg
        ValidateCHSettings = False
    End If
End Function

Private Sub btnStartCHTest_Click()
    btnCHTestUp.Enabled = True
    btnCHTestDown.Enabled = True
    For i = 0 To 59
        If rbCHTest(i).Value = True Then
            chTest = i + 1
        End If
    Next i
    timerCHTest.Enabled = True
    If cbAuto.Value = 1 Then
        timerAuto.Interval = CInt(tbAuto.Text) * 1000
        timerAuto.Enabled = True
    End If
End Sub

Private Sub btnStop_Click()
    Dim r As Integer
    r = MsgBox("测试未完成，终止测试会导致程序退出！" & vbCrLf & vbCrLf & "是否确定退出程序？", vbYesNo + vbExclamation)
    If r = vbYes Then
        Unload Me
    End If
End Sub

Private Sub btnStopCHTest_Click()
    btnCHTestUp.Enabled = False
    btnCHTestDown.Enabled = False
    timerCHTest.Enabled = False
    timerAuto.Enabled = False
End Sub

Private Sub btnTempApply_Click()
    If IsNumeric(tbLowTemp.Text) And IsNumeric(tbLTTime.Text) And IsNumeric(tbHighTemp.Text) And IsNumeric(tbHTTime.Text) Then
        If CInt(tbLowTemp.Text) >= CInt(tbHighTemp.Text) Then
            MsgBox "最低温度不能大于或等于最高温度！"
        ElseIf CInt(tbLTTime.Text) <= 0 Or CInt(tbHTTime.Text) <= 0 Then
            MsgBox "持续时间必须大于等于0！"
        ElseIf CInt(tbLowTemp.Text) >= 200 Or CInt(tbHighTemp.Text) > 200 Or CInt(tbLTTime.Text) >= 600 _
                Or CInt(tbHTTime.Text) >= 600 Then
            MsgBox "输入的温度或时间值超出可用范围！"
        Else
            tbLowTemp.Text = CInt(tbLowTemp.Text)
            tbLTTime.Text = CInt(tbLTTime.Text)
            tbHighTemp.Text = CInt(tbHighTemp.Text)
            tbHTTime.Text = CInt(tbHTTime.Text)
        
            SaveTempComConfig "config.cfg"
            
            hasTempAuth = False
            tbLowTemp.Enabled = False
            tbLTTime.Enabled = False
            tbHighTemp.Enabled = False
            tbHTTime.Enabled = False
            btnTempApply.Enabled = False
            btnSetTemp.Enabled = True
        End If
    Else
        MsgBox "请输入正确的数字！"
    End If
End Sub

Private Sub btnCOMApply_Click()
    If IsNumeric(tbTempCOM.Text) And IsNumeric(tb1680COM.Text) Then
        If CInt(tbTempCOM.Text) = CInt(tb1680COM.Text) Then
            MsgBox "两个端口不能相同！"
        Else
            SaveTempComConfig "config.cfg"
            
            hasCOMAuth = False
            btnSetCOM.Enabled = True
            tbTempCOM.Enabled = False
            tb1680COM.Enabled = False
            btnCOMApply.Enabled = False
        End If
    Else
        MsgBox "请输入对应的端口数字！"
    End If
End Sub

Private Sub cbbCHSec1_Click()
    For i = 0 To 9
        cbbB1CH(i).ListIndex = cbbCHSec1.ListIndex
    Next i
    For i = 0 To 1
        cbbB2CH(i).ListIndex = cbbCHSec1.ListIndex
    Next i
End Sub

Private Sub cbbCHSec2_Click()
    For i = 2 To 9
        cbbB2CH(i).ListIndex = cbbCHSec2.ListIndex
    Next i
    For i = 0 To 3
        cbbB3CH(i).ListIndex = cbbCHSec2.ListIndex
    Next i
End Sub

Private Sub cbbCHSec3_Click()
    For i = 4 To 9
        cbbB3CH(i).ListIndex = cbbCHSec3.ListIndex
    Next i
    For i = 0 To 5
        cbbB4CH(i).ListIndex = cbbCHSec3.ListIndex
    Next i
End Sub

Private Sub cbbCHSec4_Click()
    For i = 6 To 9
        cbbB4CH(i).ListIndex = cbbCHSec4.ListIndex
    Next i
    For i = 0 To 7
        cbbB5CH(i).ListIndex = cbbCHSec4.ListIndex
    Next i
End Sub

Private Sub cbbCHSec5_Click()
    For i = 8 To 9
        cbbB5CH(i).ListIndex = cbbCHSec5.ListIndex
    Next i
    For i = 0 To 9
        cbbB6CH(i).ListIndex = cbbCHSec5.ListIndex
    Next i
End Sub

Private Sub cbCHSec1_Click()
    If cbCHSec1.Value = 1 Then
        For i = 0 To 9
            cbB1CH(i).Value = 1
        Next i
        For i = 0 To 1
            cbB2CH(i).Value = 1
        Next i
    Else
        For i = 0 To 9
            cbB1CH(i).Value = 0
        Next i
        For i = 0 To 1
            cbB2CH(i).Value = 0
        Next i
    End If
End Sub

Private Sub cbCHSec2_Click()
    If cbCHSec2.Value = 1 Then
        For i = 2 To 9
            cbB2CH(i).Value = 1
        Next i
        For i = 0 To 3
            cbB3CH(i).Value = 1
        Next i
    Else
        For i = 2 To 9
            cbB2CH(i).Value = 0
        Next i
        For i = 0 To 3
            cbB3CH(i).Value = 0
        Next i
    End If
End Sub

Private Sub cbCHSec3_Click()
    If cbCHSec3.Value = 1 Then
        For i = 4 To 9
            cbB3CH(i).Value = 1
        Next i
        For i = 0 To 5
            cbB4CH(i).Value = 1
        Next i
    Else
        For i = 4 To 9
            cbB3CH(i).Value = 0
        Next i
        For i = 0 To 5
            cbB4CH(i).Value = 0
        Next i
    End If
End Sub

Private Sub cbCHSec4_Click()
    If cbCHSec4.Value = 1 Then
        For i = 6 To 9
            cbB4CH(i).Value = 1
        Next i
        For i = 0 To 7
            cbB5CH(i).Value = 1
        Next i
    Else
        For i = 6 To 9
            cbB4CH(i).Value = 0
        Next i
        For i = 0 To 7
            cbB5CH(i).Value = 0
        Next i
    End If
End Sub

Private Sub cbCHSec5_Click()
    If cbCHSec5.Value = 1 Then
        For i = 8 To 9
            cbB5CH(i).Value = 1
        Next i
        For i = 0 To 9
            cbB6CH(i).Value = 1
        Next i
    Else
        For i = 8 To 9
            cbB5CH(i).Value = 0
        Next i
        For i = 0 To 9
            cbB6CH(i).Value = 0
        Next i
    End If
End Sub

Private Sub Command1_Click()
    Dim getTempStr As String
    getTempStr = "01RSD,01,0001"
    SendTempUart getTempStr
End Sub

Private Sub Command2_Click()
    Dim get1680Str As String
    If cbbBTest.ListIndex < 10 Then
        get1680Str = "0" & cbbBTest.Text
    Else
        get1680Str = cbbBTest.Text
    End If
    tb1680ComInput.Text = Send1680Uart(get1680Str)
    tb1680ComReading.Text = Extract1680Data(tb1680ComInput.Text)
End Sub

Private Sub Command4_Click()
    TestBoardConnection
    MsgBox "重新检测完成！"
End Sub

Private Sub Form_Load()
    autoSave = False
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    GetTempComConfig "config.cfg"
    
    ReDim Preserve modelName(0 To 0) As String
    ReDim Preserve modelResist(0 To 0) As Integer
    ReDim Preserve modelCoeff(0 To 0) As Double
    GetModelConfig App.Path & "\传感器系数设定.xls"
    FillComboBox
    
    cbP1.Value = 1
    
    comTemp.CommPort = tbTempCOM.Text
    comTemp.Settings = "9600,n,8,1"
    comTemp.Handshaking = comNone
    comTemp.InputLen = 1
    
    com1680.CommPort = tb1680COM.Text
    com1680.Settings = "9600,n,8,1"
    com1680.Handshaking = comNone
    com1680.InputLen = 20
    
    TestBoardConnection
    
    cbP2.Value = 1
    
    hasTempAuth = False
    hasCOMAuth = False
    timerInterval = 60000
    timerTemp.Interval = timerInterval
    timerTest.Interval = timerInterval
    
    ' oven temp reading is updated very 10 seconds
    timerOvenTemp.Interval = timerInterval / 6
    
    ' indicate there is no test done yet
    TestDone = False
    
    Paused = False
End Sub

Private Sub GetTempComConfig(fileStr As String)
    Dim f As Long
    f = FreeFile()
    Dim tempConfigStr As String
    Dim comConfigStr As String
    tempConfigStr = String(80, " ")
    comConfigStr = String(20, " ")
    
    Open "config.cfg" For Binary As #f
    Get #f, , tempConfigStr
    tempConfigStr = decryptStr(tempConfigStr)
    Get #f, 80, comConfigStr
    comConfigStr = decryptStr(comConfigStr)
    Close #f
    
    Dim tempStr() As String
    tempStr = Split(tempConfigStr, "|")
    tbLowTemp.Text = tempStr(0)
    tbLTTime.Text = tempStr(1)
    tbHighTemp.Text = tempStr(2)
    tbHTTime.Text = tempStr(3)
    
    Dim comStr() As String
    comStr = Split(comConfigStr, "|")
    tbTempCOM.Text = comStr(0)
    tb1680COM.Text = comStr(1)
End Sub

Private Sub GetModelConfig(fileStr As String)
    Dim f As New Excel.Application
    Dim book As Excel.Workbook
    Dim sheet As Excel.Worksheet
    
    Set book = f.Workbooks.Open(fileStr, , , , "123456")
    Set sheet = book.Sheets.Item(1)
    
    Dim i As Integer
    i = 0
    Do
        modelName(i) = sheet.Cells(i + 1, 1)
        modelResist(i) = sheet.Cells(i + 1, 2)
        modelCoeff(i) = sheet.Cells(i + 1, 3)
        If Not StrComp(sheet.Cells(i + 2, 1), "") = 0 Then
            ReDim Preserve modelName(0 To UBound(modelName) + 1) As String
            ReDim Preserve modelResist(0 To UBound(modelResist) + 1) As Integer
            ReDim Preserve modelCoeff(0 To UBound(modelCoeff) + 1) As Double
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    
    f.ActiveWorkbook.Close False, fileStr
    f.Quit
End Sub

Private Sub FillComboBox()
    For i = 0 To 59
        cbbBTest.AddItem i + 1, i
    Next i
    For i = 0 To UBound(modelName)
        cbbCHSec1.AddItem modelName(i), i
        cbbCHSec2.AddItem modelName(i), i
        cbbCHSec3.AddItem modelName(i), i
        cbbCHSec4.AddItem modelName(i), i
        cbbCHSec5.AddItem modelName(i), i
        For j = 0 To 9
            cbbB1CH(j).AddItem modelName(i), i
            cbbB1CH(j).ItemData(cbbB1CH(j).NewIndex) = modelCoeff(i)
            cbbB2CH(j).AddItem modelName(i), i
            cbbB2CH(j).ItemData(cbbB1CH(j).NewIndex) = modelCoeff(i)
            cbbB3CH(j).AddItem modelName(i), i
            cbbB3CH(j).ItemData(cbbB1CH(j).NewIndex) = modelCoeff(i)
            cbbB4CH(j).AddItem modelName(i), i
            cbbB4CH(j).ItemData(cbbB1CH(j).NewIndex) = modelCoeff(i)
            cbbB5CH(j).AddItem modelName(i), i
            cbbB5CH(j).ItemData(cbbB1CH(j).NewIndex) = modelCoeff(i)
            cbbB6CH(j).AddItem modelName(i), i
            cbbB6CH(j).ItemData(cbbB1CH(j).NewIndex) = modelCoeff(i)
        Next j
    Next i
End Sub

Private Sub TestBoardConnection()
    If StrComp(Mid$(Send1680Uart("01"), 10, 1), "S") = 0 Then
        lblB1.ForeColor = vbGreen
        B1Installed = True
    Else
        lblB1.ForeColor = vbRed
        B1Installed = False
    End If
    
    If StrComp(Mid$(Send1680Uart("11"), 10, 1), "S") = 0 Then
        lblB2.ForeColor = vbGreen
        B2Installed = True
    Else
        lblB2.ForeColor = vbRed
        B2Installed = False
    End If
    
    If StrComp(Mid$(Send1680Uart("21"), 10, 1), "S") = 0 Then
        lblB3.ForeColor = vbGreen
        B3Installed = True
    Else
        lblB3.ForeColor = vbRed
        B3Installed = False
    End If
    
    If StrComp(Mid$(Send1680Uart("31"), 10, 1), "S") = 0 Then
        lblB4.ForeColor = vbGreen
        B4Installed = True
    Else
        lblB4.ForeColor = vbRed
        B4Installed = False
    End If
    
    If StrComp(Mid$(Send1680Uart("41"), 10, 1), "S") = 0 Then
        lblB5.ForeColor = vbGreen
        B5Installed = True
    Else
        lblB5.ForeColor = vbRed
        B5Installed = False
    End If
    
    If StrComp(Mid$(Send1680Uart("51"), 10, 1), "S") = 0 Then
        lblB6.ForeColor = vbGreen
        B6Installed = True
    Else
        lblB6.ForeColor = vbRed
        B6Installed = False
    End If
    
    If B1Installed = False Then
        For i = 0 To 9
            lblCH(i).ForeColor = vbRed
            cbB1CH(i).Enabled = False
        Next i
    Else
        For i = 0 To 9
            lblCH(i).ForeColor = vbBlack
            cbB1CH(i).Enabled = True
        Next i
    End If
    If B2Installed = False Then
        For i = 0 To 9
            lblCH(i + 10).ForeColor = vbRed
            cbB2CH(i).Enabled = False
        Next i
    Else
        For i = 0 To 9
            lblCH(i + 10).ForeColor = vbBlack
            cbB2CH(i).Enabled = True
        Next i
    End If
    If B3Installed = False Then
        For i = 0 To 9
            lblCH(i + 20).ForeColor = vbRed
            cbB3CH(i).Enabled = False
        Next i
    Else
        For i = 0 To 9
            lblCH(i + 20).ForeColor = vbBlack
            cbB3CH(i).Enabled = True
        Next i
    End If
    If B4Installed = False Then
        For i = 0 To 9
            lblCH(i + 30).ForeColor = vbRed
            cbB4CH(i).Enabled = False
        Next i
    Else
        For i = 0 To 9
            lblCH(i + 30).ForeColor = vbBlack
            cbB4CH(i).Enabled = True
        Next i
    End If
    If B5Installed = False Then
        For i = 0 To 9
            lblCH(i + 40).ForeColor = vbRed
            cbB5CH(i).Enabled = False
        Next i
    Else
        For i = 0 To 9
            lblCH(i + 40).ForeColor = vbBlack
            cbB5CH(i).Enabled = True
        Next i
    End If
    If B6Installed = False Then
        For i = 0 To 9
            lblCH(i + 50).ForeColor = vbRed
            cbB6CH(i).Enabled = False
        Next i
    Else
        For i = 0 To 9
            lblCH(i + 50).ForeColor = vbBlack
            cbB6CH(i).Enabled = True
        Next i
    End If
End Sub

Private Sub SaveTempComConfig(fileStr As String)
    Dim f As Long
    f = FreeFile()
    
    'empty the file
    Open "config.cfg" For Output As #f
    Close #f
    
    Open "config.cfg" For Binary As #f
    Dim tempStr As String
    tempStr = tbLowTemp.Text & "|" & tbLTTime.Text & "|" & tbHighTemp.Text & "|" & tbHTTime.Text
    tempStr = encryptStr(tempStr)
    Put #f, , tempStr
    Dim comStr As String
    comStr = tbTempCOM.Text & "|" & tb1680COM.Text
    comStr = encryptStr(comStr)
    Put #f, 80, comStr
    Close #f
End Sub

Private Function encryptStr(str As String) As String
    Dim eStr As String
    For i = 1 To Len(str)
        eStr = eStr & ((Asc(Mid(str, i, 1)) + 100) / 5)
        If Not i = Len(str) Then
            eStr = eStr + "|"
        End If
    Next i
    encryptStr = eStr
End Function

Private Function decryptStr(str As String) As String
    Dim splitEStr() As String
    splitEStr = Split(str, "|")
    
    Dim dStr As String
    For i = 0 To UBound(splitEStr())
        dStr = dStr + Chr(Val(splitEStr(i)) * 5 - 100)
    Next i
    decryptStr = dStr
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub SendTempUart(outStr As String)
    Dim tempOutStr As String
    tempOutStr = Chr$(2)
    
    Dim checksum As Integer
    checksum = 0
    For i = 1 To Len(outStr)
        tempOutStr = tempOutStr + Mid$(outStr, i, 1)
        checksum = checksum + Asc(Mid$(outStr, i, 1))
    Next i
    
    Dim checksumStr As String
    checksumStr = Mid$(Hex$(checksum), Len(Hex$(checksum)) - 1, 2)
    tempOutStr = tempOutStr & checksumStr
    tempOutStr = tempOutStr & vbCrLf
    timeout = False
    timeoutTime = 0
    
SendTemp:
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    comTemp.PortOpen = True
    If comTemp.PortOpen Then
        comTemp.Output = tempOutStr
        
        Dim tempInStr As String
        timerTimeout.Enabled = True
        Do
            If timeoutTime = 3 Then
                MsgBox "未发现可用的串口温度数据传输。请检查串口连接！"
                GoTo EndFuncTemp
            End If
            If timeout = True Then
                GoTo TimeoutERRTemp
            End If
            DoEvents
            tempInStr = tempInStr & comTemp.Input
            countRead = countRead + 1
        Loop Until InStr(tempInStr, vbCrLf)
        If VerifyTempChecksum(tempInStr) Then
            tbTempComInput.Text = tempInStr
            tbCurrTemp.Text = ExtractTempData(tempInStr)
        Else
            MsgBox "温度数据校检错误。请检查串口连接！"
        End If
        GoTo EndFuncTemp
    End If
TimeoutERRTemp:
    timeout = False
    timeoutTime = timeoutTime + 1
    comTemp.PortOpen = False
    GoTo SendTemp
EndFuncTemp:
    timeout = False
    timeoutTime = 0
    timerTimeout.Enabled = False
    comTemp.PortOpen = False
End Sub

Private Function Send1680Uart(outStr As String) As String
    Dim TIOutStr As String
    TIOutStr = Chr$(2) & outStr & vbCrLf
    timeout = False
    timeoutTime = 0
    
    'add a time interval for re-opening the port
    'previously error of re-opening already opened port may occur
    
    Sleep 300
    
Send1680:
    com1680.PortOpen = True
    If com1680.PortOpen Then
        com1680.Output = TIOutStr
        
        Dim TIInStr As String
        timerTimeout.Enabled = True
        Do
            If timeoutTime = 3 Then
                MsgBox "未发现可用的1680数据传输。请检查串口连接！"
                Send1680Uart = "ERR"
                GoTo EndFunc1680
            End If
            If timeout = True Then
                GoTo TimeoutERR1680
            End If
            DoEvents
            TIInStr = TIInStr & com1680.Input
        Loop Until InStr(TIInStr, vbCrLf)
        Send1680Uart = TIInStr
        GoTo EndFunc1680
    End If
TimeoutERR1680:
    timeout = False
    timeoutTime = timeoutTime + 1
    com1680.PortOpen = False
    GoTo Send1680
EndFunc1680:
    timeout = False
    timeoutTime = 0
    timerTimeout.Enabled = False
    com1680.PortOpen = False
End Function

Private Function SetTempUart(temp As Integer) As Boolean
    Dim tempOutStr As String
    tempOutStr = Chr$(2) & "01WSD,01,0201,"
    
    Dim tempStr As String
    tempStr = Hex$(temp * 10)
    'add zero in front when necessary
    Dim numZero As Integer
    numZero = 4 - Len(tempStr)
    For i = 1 To numZero
        tempStr = "0" & tempStr
    Next i
    tempOutStr = tempOutStr & tempStr
    
    Dim checksum As Integer
    checksum = 0
    For i = 2 To Len(tempOutStr)
        checksum = checksum + Asc(Mid$(tempOutStr, i, 1))
    Next i
    
    Dim checksumStr As String
    checksumStr = Mid$(Hex$(checksum), Len(Hex$(checksum)) - 1, 2)
    tempOutStr = tempOutStr & checksumStr
    tempOutStr = tempOutStr & vbCrLf
    
    comTemp.PortOpen = True
    If comTemp.PortOpen Then
        comTemp.Output = tempOutStr
        Sleep 1000
        
        Dim tempInStr As String
        Dim countRead As Integer
        countRead = 0
        Do
            DoEvents
            tempInStr = tempInStr & comTemp.Input
            countRead = countRead + 1
        Loop Until InStr(tempInStr, vbCrLf) Or countRead = 10000
        If countRead = 10000 Then
            MsgBox "未发现可用的串口温度数据传输。请检查串口连接！"
            SetTempUart = False
        Else
            If VerifyTempChecksum(tempInStr) And InStr(tempInStr, "01WSD,OK15") Then
                tbTempComInput.Text = "温度设定为" & temp & "℃成功！"
                SetTempUart = True
            ElseIf Not VerifyTempChecksum(tempInStr) Then
                MsgBox "温度数据校检错误。请检查串口连接！"
                SetTempUart = False
            Else
                MsgBox "温度设定失败！"
                SetTempUart = False
            End If
        End If
        comTemp.PortOpen = False
    End If
End Function

Private Function VerifyTempChecksum(tempInStr As String) As Boolean
    Dim dataStr As String
    Dim checksumStr As String
    dataStr = Mid$(tempInStr, 2, Len(tempInStr) - 5)
    checksumStr = Mid$(tempInStr, Len(tempInStr) - 3, 2)
    
    Dim checksumCalcInt As Integer
    checksumCalcInt = 0
    For i = 1 To Len(tempInStr) - 5
        checksumCalcInt = checksumCalcInt + Asc(Mid$(dataStr, i, 1))
    Next i
    checksumCalcStr = Mid$(Hex$(checksumCalcInt), Len(Hex$(checksumCalcInt)) - 1, 2)
    
    If StrComp(checksumStr, checksumCalcStr) = 0 Then
        VerifyTempChecksum = True
    Else
        VerifyTempChecksum = False
    End If
End Function

Private Function ExtractTempData(tempInStr As String) As String
    Dim dataStr As String
    dataStr = Mid$(tempInStr, 2, Len(tempInStr) - 5)
    If StrComp(Left$(dataStr, 9), "01RSD,OK,") = 0 Then
        ExtractTempData = "" & ((CInt("&H" & Mid$(dataStr, 10, 4))) / 10)
    Else
        ExtractTempData = "ERR"
    End If
End Function

Private Function Extract1680Data(TIInStr As String) As Double
    Dim dataStr As String
    dataStr = Mid$(TIInStr, 2, 7)
    If StrComp(Mid$(TIInStr, 9, 1), "A") = 0 Then
        Extract1680Data = CDbl(dataStr)
    ElseIf StrComp(Mid$(TIInStr, 9, 1), "B") = 0 Then
        Extract1680Data = CDbl(dataStr) / 10
    ElseIf StrComp(Mid$(TIInStr, 9, 1), "C") = 0 Then
        Extract1680Data = CDbl(dataStr) / 100
    ElseIf StrComp(Mid$(TIInStr, 9, 1), "D") = 0 Then
        Extract1680Data = CDbl(dataStr) / 1000
    ElseIf StrComp(Mid$(TIInStr, 9, 1), "E") = 0 Then
        Extract1680Data = CDbl(dataStr) / 10000
    Else
        Extract1680Data = 0
    End If
End Function

Private Sub CountTime(intTime As Integer)
    timerTempTickCount = 0
    timerTemp.Enabled = True
    Do
        DoEvents
    Loop Until timerTempTickCount >= intTime
    timerTemp.Enabled = False
End Sub

Private Sub ReadSensorData(LowOrHigh As String)
    If B1Installed = True And StrComp(LowOrHigh, "LT") = 0 Then
        For i = 0 To 9
            If cbB1CH(i).Value = 1 Then
                If i < 9 Then
                    tbB1CHLT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("0" & (i + 1))), 4, vbTrue, vbFalse, vbFalse)
                Else
                    tbB1CHLT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 1))), 4, vbTrue, vbFalse, vbFalse)
                End If
            End If
        Next i
    End If
    
    If B1Installed = True And StrComp(LowOrHigh, "HT") = 0 Then
        For i = 0 To 9
            If cbB1CH(i).Value = 1 Then
                If i < 9 Then
                    tbB1CHHT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("0" & (i + 1))), 4, vbTrue, vbFalse, vbFalse)
                Else
                    tbB1CHHT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 1))), 4, vbTrue, vbFalse, vbFalse)
                End If
            End If
        Next i
    End If
    
    If B2Installed = True And StrComp(LowOrHigh, "LT") = 0 Then
        For i = 0 To 9
            If cbB2CH(i).Value = 1 Then
                tbB2CHLT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 11))), 4, vbTrue, vbFalse, vbFalse)
            End If
        Next i
    End If
    
    If B2Installed = True And StrComp(LowOrHigh, "HT") = 0 Then
        For i = 0 To 9
            If cbB2CH(i).Value = 1 Then
                    tbB2CHHT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 11))), 4, vbTrue, vbFalse, vbFalse)
            End If
        Next i
    End If
    
    If B3Installed = True And StrComp(LowOrHigh, "LT") = 0 Then
        For i = 0 To 9
            If cbB3CH(i).Value = 1 Then
                tbB3CHLT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 21))), 4, vbTrue, vbFalse, vbFalse)
            End If
        Next i
    End If
    
    If B3Installed = True And StrComp(LowOrHigh, "HT") = 0 Then
        For i = 0 To 9
            If cbB3CH(i).Value = 1 Then
                    tbB3CHHT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 21))), 4, vbTrue, vbFalse, vbFalse)
            End If
        Next i
    End If
    
    If B4Installed = True And StrComp(LowOrHigh, "LT") = 0 Then
        For i = 0 To 9
            If cbB4CH(i).Value = 1 Then
                tbB4CHLT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 31))), 4, vbTrue, vbFalse, vbFalse)
            End If
        Next i
    End If
    
    If B4Installed = True And StrComp(LowOrHigh, "HT") = 0 Then
        For i = 0 To 9
            If cbB4CH(i).Value = 1 Then
                    tbB4CHHT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 31))), 4, vbTrue, vbFalse, vbFalse)
            End If
        Next i
    End If
    
    If B5Installed = True And StrComp(LowOrHigh, "LT") = 0 Then
        For i = 0 To 9
            If cbB5CH(i).Value = 1 Then
                tbB5CHLT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 41))), 4, vbTrue, vbFalse, vbFalse)
            End If
        Next i
    End If
    
    If B5Installed = True And StrComp(LowOrHigh, "HT") = 0 Then
        For i = 0 To 9
            If cbB5CH(i).Value = 1 Then
                    tbB5CHHT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 41))), 4, vbTrue, vbFalse, vbFalse)
            End If
        Next i
    End If
    
    If B6Installed = True And StrComp(LowOrHigh, "LT") = 0 Then
        For i = 0 To 9
            If cbB6CH(i).Value = 1 Then
                tbB6CHLT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 51))), 4, vbTrue, vbFalse, vbFalse)
            End If
        Next i
    End If
    
    If B6Installed = True And StrComp(LowOrHigh, "HT") = 0 Then
        For i = 0 To 9
            If cbB6CH(i).Value = 1 Then
                    tbB6CHHT(i).Text = FormatNumber(Extract1680Data(Send1680Uart("" & (i + 51))), 4, vbTrue, vbFalse, vbFalse)
            End If
        Next i
    End If
End Sub

Private Sub ComputeL()
    If B1Installed = True Then
        For i = 0 To 9
            If cbB1CH(i).Value = 1 Then
                tbB1CHL(i).Text = CInt(CDbl((Val(tbB1CHHT(i)) - Val(tbB1CHLT(i))) * _
                    cbbB1CH(i).ItemData(cbbB1CH(i).ListIndex)))
            End If
        Next i
    End If
    If B2Installed = True Then
        For i = 0 To 9
            If cbB2CH(i).Value = 1 Then
                tbB2CHL(i).Text = CInt(CDbl((Val(tbB2CHHT(i)) - Val(tbB2CHLT(i))) * _
                    cbbB2CH(i).ItemData(cbbB2CH(i).ListIndex)))
            End If
        Next i
    End If
    If B3Installed = True Then
        For i = 0 To 9
            If cbB3CH(i).Value = 1 Then
                tbB3CHL(i).Text = CInt(CDbl((Val(tbB3CHHT(i)) - Val(tbB3CHLT(i))) * _
                    cbbB3CH(i).ItemData(cbbB3CH(i).ListIndex)))
            End If
        Next i
    End If
    If B4Installed = True Then
        For i = 0 To 9
            If cbB4CH(i).Value = 1 Then
                tbB4CHL(i).Text = CInt(CDbl((Val(tbB4CHHT(i)) - Val(tbB4CHLT(i))) * _
                    cbbB4CH(i).ItemData(cbbB4CH(i).ListIndex)))
            End If
        Next i
    End If
    If B5Installed = True Then
        For i = 0 To 9
            If cbB5CH(i).Value = 1 Then
                tbB5CHL(i).Text = CInt(CDbl((Val(tbB5CHHT(i)) - Val(tbB5CHLT(i))) * _
                    cbbB5CH(i).ItemData(cbbB5CH(i).ListIndex)))
            End If
        Next i
    End If
    If B6Installed = True Then
        For i = 0 To 9
            If cbB6CH(i).Value = 1 Then
                tbB6CHL(i).Text = CInt(CDbl((Val(tbB6CHHT(i)) - Val(tbB6CHLT(i))) * _
                    cbbB6CH(i).ItemData(cbbB6CH(i).ListIndex)))
            End If
        Next i
    End If
End Sub

Private Sub lblCHSec1_Click()
    If cbCHSec1.Value = 0 Then
        cbCHSec1.Value = 1
    Else
        cbCHSec1.Value = 0
    End If
    cbCHSec1_Click
End Sub

Private Sub lblCHSec2_Click()
    If cbCHSec2.Value = 0 Then
        cbCHSec2.Value = 1
    Else
        cbCHSec2.Value = 0
    End If
    cbCHSec2_Click
End Sub

Private Sub lblCHSec3_Click()
    If cbCHSec3.Value = 0 Then
        cbCHSec3.Value = 1
    Else
        cbCHSec3.Value = 0
    End If
    cbCHSec3_Click
End Sub

Private Sub lblCHSec4_Click()
    If cbCHSec4.Value = 0 Then
        cbCHSec4.Value = 1
    Else
        cbCHSec4.Value = 0
    End If
    cbCHSec4_Click
End Sub

Private Sub lblCHSec5_Click()
    If cbCHSec5.Value = 0 Then
        cbCHSec5.Value = 1
    Else
        cbCHSec5.Value = 0
    End If
    cbCHSec5_Click
End Sub

Private Sub rbCHTest_Click(Index As Integer)
    chTest = Index + 1
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            'ToDo: Add 'New' button code.
            MsgBox "Add 'New' button code."
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFilePrint_Click()
    cdlg.CancelError = True
    cdlg.DialogTitle = "打印测试数据"
    cdlg.Filter = "Word file (*.doc)|*.doc"
    cdlg.FileName = Replace$(tbEndTime.Text, ":", "")
    cdlg.ShowOpen
    
    PrintResult cdlg.FileName
End Sub

Private Sub PrintResult(file As String)
    Dim wordApp As Word.Application
    Dim wordDoc As Word.Document
    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Open(file)
    With wordApp
        .ActiveDocument.PrintOut
        .ActiveDocument.Close
    End With
    wordDoc.Close False
    wordApp.Quit False
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

Private Sub mnuFileSave_Click()
    If TestDone = False Then
        MsgBox "未做测试，无数据可以保存！"
    Else
        cdlg.CancelError = True
        cdlg.DialogTitle = "保存测试数据"
        cdlg.Filter = "Word file (*.doc)|*.doc"
        cdlg.FileName = Replace$(tbStartTime.Text, ":", "")
        cdlg.ShowSave
        
        SaveResultWord cdlg.FileName
        SaveResultSQL
    End If
End Sub

Private Sub SaveResultWord(file As String)
    Dim wordApp As Word.Application
    Dim wordDoc As Word.Document
    Dim wordRng As Word.Range
    
    Dim wordPara1 As Word.Paragraph
    Dim wordPara2 As Word.Paragraph
    Dim wordParaSep As Word.Paragraph
    Dim wordTable As Word.Table
    
    Set wordApp = CreateObject("Word.Application")
    
    With wordApp
        .WindowState = wdWindowStateMaximize
        Set wordDoc = .Documents.Add
        
        With wordDoc
            With .PageSetup
                .TopMargin = CentimetersToPoints(1.5)
                .BottomMargin = CentimetersToPoints(1.5)
                .LeftMargin = CentimetersToPoints(1.5)
                .RightMargin = CentimetersToPoints(1.5)
            End With
            
            Dim j As Integer
            If cbB5CH(8).Value = 1 Then
                j = 5
            ElseIf cbB4CH(6).Value = 1 Then
                j = 4
            ElseIf cbB3CH(4).Value = 1 Then
                j = 3
            ElseIf cbB2CH(2).Value = 1 Then
                j = 2
            ElseIf cbB1CH(0).Value = 1 Then
                j = 1
            End If
            
            For i = 1 To j
                If i = 1 Then
                    Set wordPara1 = .Content.Paragraphs.Add
                Else
                    Set wordPara1 = .Content.Paragraphs.Add(.Bookmarks("\endofdoc").Range)
                End If
                With wordPara1.Range
                    .Text = "ZERO   TEMPERATURE   COMPENSATION   REPORT"
                    .Font.Name = "Times New Roman"
                    .Font.Size = 12
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .InsertParagraphAfter
                End With
            
                Set wordPara2 = .Content.Paragraphs.Add(.Bookmarks("\endofdoc").Range)
                With wordPara2.Range
                    .Text = "SN:A" & tbEndTime.Text & "   INSPECTOR:               PAGE:" & i & "      TOTAL:" & j & vbCrLf
                    .Font.Name = "Times New Roman"
                    .Font.Size = 12
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .InsertParagraphAfter
                End With
            
                Set wordTable = .Tables.Add(.Bookmarks("\endofdoc").Range, 13, 6)
                With wordTable
                    .Borders.Enable = True
                    .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Range.Font.Name = "Times New Roman"
                    .Range.Font.Size = 12
                    .Range.Font.Bold = True
                    .Range.Rows.Height = CentimetersToPoints(0.8)
                    .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                
                    .Cell(1, 1).Range.Text = "MODEL"
                    .Columns(1).Width = CentimetersToPoints(4)
                    .Cell(1, 2).Range.Text = "NO"
                    .Columns(2).Width = CentimetersToPoints(1.5)
                    .Cell(1, 3).Range.Text = "27℃"
                    .Columns(3).Width = CentimetersToPoints(3.5)
                    .Cell(1, 4).Range.Text = "77℃"
                    .Columns(4).Width = CentimetersToPoints(3.5)
                    .Cell(1, 5).Range.Text = "L(mm)"
                    .Columns(5).Width = CentimetersToPoints(3.5)
                    .Cell(1, 6).Range.Text = "CH"
                    .Columns(6).Width = CentimetersToPoints(2)
                    
                    If i = 1 Then
                        For a = 0 To 9
                            If cbB1CH(a).Value = 1 And cbB1CH(a).Enabled = True Then
                                .Cell(a + 2, 1).Range.Text = cbbB1CH(a).Text
                                .Cell(a + 2, 3).Range.Text = tbB1CHLT(a).Text
                                .Cell(a + 2, 4).Range.Text = tbB1CHHT(a).Text
                                .Cell(a + 2, 5).Range.Text = tbB1CHL(a).Text
                                .Cell(a + 2, 6).Range.Text = (a + 1)
                            End If
                        Next a
                        For a = 0 To 1
                            If cbB2CH(a).Value = 1 And cbB2CH(a).Enabled = True Then
                                .Cell(a + 12, 1).Range.Text = cbbB2CH(a).Text
                                .Cell(a + 12, 3).Range.Text = tbB2CHLT(a).Text
                                .Cell(a + 12, 4).Range.Text = tbB2CHHT(a).Text
                                .Cell(a + 12, 5).Range.Text = tbB2CHL(a).Text
                                .Cell(a + 12, 6).Range.Text = (a + 11)
                            End If
                        Next a
                    End If
                    
                    If i = 2 Then
                        For a = 2 To 9
                            If cbB2CH(a).Value = 1 And cbB2CH(a).Enabled = True Then
                                .Cell(a, 1).Range.Text = cbbB2CH(a).Text
                                .Cell(a, 3).Range.Text = tbB2CHLT(a).Text
                                .Cell(a, 4).Range.Text = tbB2CHHT(a).Text
                                .Cell(a, 5).Range.Text = tbB2CHL(a).Text
                                .Cell(a, 6).Range.Text = (a + 11)
                            End If
                        Next a
                        For a = 0 To 3
                            If cbB3CH(a).Value = 1 And cbB3CH(a).Enabled = True Then
                                .Cell(a + 10, 1).Range.Text = cbbB3CH(a).Text
                                .Cell(a + 10, 3).Range.Text = tbB3CHLT(a).Text
                                .Cell(a + 10, 4).Range.Text = tbB3CHHT(a).Text
                                .Cell(a + 10, 5).Range.Text = tbB3CHL(a).Text
                                .Cell(a + 10, 6).Range.Text = (a + 21)
                            End If
                        Next a
                    End If
                    
                    If i = 3 Then
                        For a = 4 To 9
                            If cbB3CH(a).Value = 1 And cbB3CH(a).Enabled = True Then
                                .Cell(a - 2, 1).Range.Text = cbbB3CH(a).Text
                                .Cell(a - 2, 3).Range.Text = tbB3CHLT(a).Text
                                .Cell(a - 2, 4).Range.Text = tbB3CHHT(a).Text
                                .Cell(a - 2, 5).Range.Text = tbB3CHL(a).Text
                                .Cell(a - 2, 6).Range.Text = (a + 21)
                            End If
                        Next a
                        For a = 0 To 5
                            If cbB4CH(a).Value = 1 And cbB4CH(a).Enabled = True Then
                                .Cell(a + 8, 1).Range.Text = cbbB4CH(a).Text
                                .Cell(a + 8, 3).Range.Text = tbB4CHLT(a).Text
                                .Cell(a + 8, 4).Range.Text = tbB4CHHT(a).Text
                                .Cell(a + 8, 5).Range.Text = tbB4CHL(a).Text
                                .Cell(a + 8, 6).Range.Text = (a + 31)
                            End If
                        Next a
                    End If
                    
                    If i = 4 Then
                        For a = 6 To 9
                            If cbB4CH(a).Value = 1 And cbB4CH(a).Enabled = True Then
                                .Cell(a - 4, 1).Range.Text = cbbB4CH(a).Text
                                .Cell(a - 4, 3).Range.Text = tbB4CHLT(a).Text
                                .Cell(a - 4, 4).Range.Text = tbB4CHHT(a).Text
                                .Cell(a - 4, 5).Range.Text = tbB4CHL(a).Text
                                .Cell(a - 4, 6).Range.Text = (a + 31)
                            End If
                        Next a
                        For a = 0 To 7
                            If cbB5CH(a).Value = 1 And cbB5CH(a).Enabled = True Then
                                .Cell(a + 6, 1).Range.Text = cbbB5CH(a).Text
                                .Cell(a + 6, 3).Range.Text = tbB5CHLT(a).Text
                                .Cell(a + 6, 4).Range.Text = tbB5CHHT(a).Text
                                .Cell(a + 6, 5).Range.Text = tbB5CHL(a).Text
                                .Cell(a + 6, 6).Range.Text = (a + 41)
                            End If
                        Next a
                    End If
                    
                    If i = 5 Then
                        For a = 8 To 9
                            If cbB5CH(a).Value = 1 And cbB5CH(a).Enabled = True Then
                                .Cell(a - 6, 1).Range.Text = cbbB5CH(a).Text
                                .Cell(a - 6, 3).Range.Text = tbB5CHLT(a).Text
                                .Cell(a - 6, 4).Range.Text = tbB5CHHT(a).Text
                                .Cell(a - 6, 5).Range.Text = tbB5CHL(a).Text
                                .Cell(a - 6, 6).Range.Text = (a + 41)
                            End If
                        Next a
                        For a = 0 To 9
                            If cbB6CH(a).Value = 1 And cbB6CH(a).Enabled = True Then
                                .Cell(a + 4, 1).Range.Text = cbbB6CH(a).Text
                                .Cell(a + 4, 3).Range.Text = tbB6CHLT(a).Text
                                .Cell(a + 4, 4).Range.Text = tbB6CHHT(a).Text
                                .Cell(a + 4, 5).Range.Text = tbB6CHL(a).Text
                                .Cell(a + 4, 6).Range.Text = (a + 51)
                            End If
                        Next a
                    End If
                End With
                
                If i = 1 Or i = 3 Then
                    Set wordParaSep = .Content.Paragraphs.Add(.Bookmarks("\endofdoc").Range)
                    With wordParaSep.Range
                        .Text = vbCrLf & vbCrLf
                        .Font.Name = "Times New Roman"
                        .Font.Size = 12
                        .Font.Bold = True
                        .ParagraphFormat.Alignment = wdAlignParagraphCenter
                        .InsertParagraphAfter
                    End With
                End If
            Next i
        .SaveAs file
        End With
    End With
    
    wordDoc.Close False
    wordApp.Quit False
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    If autoSave = False Then
        MsgBox "成功保存至Word文档：" & file & "。"
    End If
End Sub

Private Sub SaveResultSQL()
    Dim conn As New ADODB.Connection
    Dim cmd As ADODB.Command
    
    conn.Open "Provider=sqloledb; Data Source=.\SQLEXPRESS; Initial Catalog=ZTC; Integrated Security=SSPI"
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = conn
        .Prepared = False
        For i = 0 To 9
            If cbB1CH(i).Value = 1 And cbB1CH(a).Enabled = True Then
                .CommandText = "INSERT INTO tests VALUES ('" & tbEndTime.Text & "'," & (i + 1) & ",'" & cbbB1CH(i).Text & "',Null," & tbB1CHLT(i).Text & "," & tbB1CHHT(i).Text & "," & tbB1CHL(i).Text & ")"
                .Execute
            End If
        Next i
        For i = 0 To 9
            If cbB2CH(i).Value = 1 And cbB2CH(a).Enabled = True Then
                .CommandText = "INSERT INTO tests VALUES ('" & tbEndTime.Text & "'," & (i + 11) & ",'" & cbbB2CH(i).Text & "',Null," & tbB2CHLT(i).Text & "," & tbB2CHHT(i).Text & "," & tbB2CHL(i).Text & ")"
                .Execute
            End If
        Next i
        For i = 0 To 9
            If cbB3CH(i).Value = 1 And cbB3CH(a).Enabled = True Then
                .CommandText = "INSERT INTO tests VALUES ('" & tbEndTime.Text & "'," & (i + 21) & ",'" & cbbB3CH(i).Text & "',Null," & tbB3CHLT(i).Text & "," & tbB3CHHT(i).Text & "," & tbB3CHL(i).Text & ")"
                .Execute
            End If
        Next i
        For i = 0 To 9
            If cbB4CH(i).Value = 1 And cbB4CH(a).Enabled = True Then
                .CommandText = "INSERT INTO tests VALUES ('" & tbEndTime.Text & "'," & (i + 31) & ",'" & cbbB4CH(i).Text & "',Null," & tbB4CHLT(i).Text & "," & tbB4CHHT(i).Text & "," & tbB4CHL(i).Text & ")"
                .Execute
            End If
        Next i
        For i = 0 To 9
            If cbB5CH(i).Value = 1 And cbB5CH(a).Enabled = True Then
                .CommandText = "INSERT INTO tests VALUES ('" & tbEndTime.Text & "'," & (i + 41) & ",'" & cbbB5CH(i).Text & "',Null," & tbB5CHLT(i).Text & "," & tbB5CHHT(i).Text & "," & tbB5CHL(i).Text & ")"
                .Execute
            End If
        Next i
        For i = 0 To 9
            If cbB6CH(i).Value = 1 And cbB6CH(a).Enabled = True Then
                .CommandText = "INSERT INTO tests VALUES ('" & tbEndTime.Text & "'," & (i + 51) & ",'" & cbbB6CH(i).Text & "',Null," & tbB6CHLT(i).Text & "," & tbB6CHHT(i).Text & "," & tbB6CHL(i).Text & ")"
                .Execute
            End If
        Next i
    End With
    conn.Close
    
    If autoSave = False Then
        MsgBox "成功保存至SQL Server 2008 Express数据库。"
    Else
        MsgBox "成功保存至Word文档和SQL Server 2008 Express数据库。"
    End If
End Sub

Private Sub mnuFileNew_Click()
    'ToDo: Add 'mnuFileNew_Click' code.
    MsgBox "Add 'mnuFileNew_Click' code."
End Sub

Private Sub timerAuto_Timer()
    btnCHTestDown_Click
End Sub

Private Sub timerCHTest_Timer()
    On Error Resume Next
    If chTest < 10 Then
        tbCH.Text = "CH0" & chTest
        tbCHTest.Text = Extract1680Data(Send1680Uart("0" & chTest))
    Else
        tbCH.Text = "CH" & chTest
        tbCHTest.Text = Extract1680Data(Send1680Uart("" & chTest))
    End If
End Sub

Private Sub timerOvenTemp_Timer()
    On Error Resume Next
    Command1_Click
End Sub

Private Sub timerTemp_Timer()
    timerTempTickCount = timerTempTickCount + 1
End Sub

Private Sub timerTest_Timer()
    timerTestTickCount = timerTestTickCount + 1
    tbTestTime.Text = timerTestTickCount
End Sub

Private Sub timerTimeout_Timer()
    timeout = True
End Sub
