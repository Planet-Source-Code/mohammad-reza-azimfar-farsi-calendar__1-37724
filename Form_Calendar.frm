VERSION 5.00
Begin VB.Form Form_Calender 
   Caption         =   " ﬁÊÌ„ Ê ”——”Ìœ"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command_Report 
      Caption         =   "ç«Å "
      Height          =   435
      Left            =   4680
      TabIndex        =   57
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   56
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text_Notes 
      Alignment       =   1  'Right Justify
      Height          =   2415
      Left            =   120
      MaxLength       =   64000
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   53
      Top             =   3240
      Width           =   9135
   End
   Begin VB.CommandButton Command_Ok 
      Caption         =   "–ŒÌ—Â «ÿ·«⁄« "
      Enabled         =   0   'False
      Height          =   435
      Left            =   7800
      TabIndex        =   54
      Top             =   5880
      Width           =   1485
   End
   Begin VB.CommandButton Command_Cancel 
      Caption         =   "»«“ê‘ "
      Height          =   435
      Left            =   6240
      TabIndex        =   55
      Top             =   5880
      Width           =   1485
   End
   Begin VB.TextBox Text_Today 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text_Year 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label_Today 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "«„—Ê“"
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   38
      Left            =   3360
      TabIndex        =   52
      Tag             =   "”Â ‘‰»Â"
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   37
      Left            =   4440
      TabIndex        =   51
      Tag             =   "œÊ‘‰»Â"
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   36
      Left            =   5520
      TabIndex        =   50
      Tag             =   "Ìﬂ‘‰»Â"
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   35
      Left            =   6600
      TabIndex        =   49
      Tag             =   "‘‰»Â"
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   34
      Left            =   120
      TabIndex        =   48
      Tag             =   "Ã„⁄Â"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   33
      Left            =   1200
      TabIndex        =   47
      Tag             =   "Å‰Ã‘‰»Â"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   32
      Left            =   2280
      TabIndex        =   46
      Tag             =   "çÂ«—‘‰»Â"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   31
      Left            =   3360
      TabIndex        =   45
      Tag             =   "”Â ‘‰»Â"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   30
      Left            =   4440
      TabIndex        =   44
      Tag             =   "œÊ‘‰»Â"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   29
      Left            =   5520
      TabIndex        =   43
      Tag             =   "Ìﬂ‘‰»Â"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   28
      Left            =   6600
      TabIndex        =   42
      Tag             =   "‘‰»Â"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   27
      Left            =   120
      TabIndex        =   41
      Tag             =   "Ã„⁄Â"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   26
      Left            =   1200
      TabIndex        =   40
      Tag             =   "Å‰Ã‘‰»Â"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   25
      Left            =   2280
      TabIndex        =   39
      Tag             =   "çÂ«—‘‰»Â"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   24
      Left            =   3360
      TabIndex        =   38
      Tag             =   "”Â ‘‰»Â"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   23
      Left            =   4440
      TabIndex        =   37
      Tag             =   "œÊ‘‰»Â"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   22
      Left            =   5520
      TabIndex        =   36
      Tag             =   "Ìﬂ‘‰»Â"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   21
      Left            =   6600
      TabIndex        =   35
      Tag             =   "‘‰»Â"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   20
      Left            =   120
      TabIndex        =   34
      Tag             =   "Ã„⁄Â"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   19
      Left            =   1200
      TabIndex        =   33
      Tag             =   "Å‰Ã‘‰»Â"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   18
      Left            =   2280
      TabIndex        =   32
      Tag             =   "çÂ«—‘‰»Â"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   17
      Left            =   3360
      TabIndex        =   31
      Tag             =   "”Â ‘‰»Â"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   16
      Left            =   4440
      TabIndex        =   30
      Tag             =   "œÊ‘‰»Â"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   15
      Left            =   5520
      TabIndex        =   29
      Tag             =   "Ìﬂ‘‰»Â"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   6600
      TabIndex        =   28
      Tag             =   "‘‰»Â"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   13
      Left            =   120
      TabIndex        =   27
      Tag             =   "Ã„⁄Â"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   1200
      TabIndex        =   26
      Tag             =   "Å‰Ã‘‰»Â"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   2280
      TabIndex        =   25
      Tag             =   "çÂ«—‘‰»Â"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   3360
      TabIndex        =   24
      Tag             =   "”Â ‘‰»Â"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   4440
      TabIndex        =   23
      Tag             =   "œÊ‘‰»Â"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   5520
      TabIndex        =   22
      Tag             =   "Ìﬂ‘‰»Â"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   6600
      TabIndex        =   21
      Tag             =   "‘‰»Â"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Tag             =   "Ã„⁄Â"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   1200
      TabIndex        =   19
      Tag             =   "Å‰Ã‘‰»Â"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   2280
      TabIndex        =   18
      Tag             =   "çÂ«—‘‰»Â"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   3360
      TabIndex        =   17
      Tag             =   "”Â ‘‰»Â"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   4440
      TabIndex        =   16
      Tag             =   "œÊ‘‰»Â"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   5520
      TabIndex        =   15
      Tag             =   "Ìﬂ‘‰»Â"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   6600
      TabIndex        =   14
      Tag             =   "‘‰»Â"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Ã„⁄Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Å‰Ã‘‰»Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   1200
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "çÂ«—‘‰»Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "”Â ‘‰»Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "œÊ‘‰»Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Ìﬂ‘‰»Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "‘‰»Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Form_Calender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim J As Integer
Dim K As Integer
'

Private Sub Set_Date()
Dim C_D As String
C_D = En_Date(Text_Year.Text & "/" & Format(List1.ListIndex + 1, "00") & "/01")
For J = 0 To Label2.Count - 1
    If Mid(Fa_Date(C_D), 6, 2) = Format(List1.ListIndex + 1, "00") Then
        If Fa_Day(C_D) = Label2(J).Tag Then
            Label2(J).Caption = CInt(Mid(Fa_Date(C_D), 9, 2))
            C_D = CDate(C_D) + 1
            Label2(J).Visible = True
        Else
            Label2(J).Visible = False
            Label2(J).Caption = "0"
        End If
    Else
        Label2(J).Visible = False
        Label2(J).Caption = "0"
    End If
Next
For K = 0 To Label2.Count - 1
    Label2(K).BorderStyle = 0
    Label2(K).BackColor = 12632100
    Text_Notes.Text = ""
    Text1.Text = ""
Next
End Sub

Private Sub Get_Data()
If Text1.Text = "" Then
    Text_Notes.Text = ""
    Exit Sub
End If
Dim Rs As New ADODB.Recordset
Rs.Open "Select * From Calendar Where Date Like '" & Text1.Text & "' ", Cn, adOpenStatic, adLockOptimistic
If Rs.RecordCount > 0 Then
    Text_Notes.Text = IIf(IsNull(Rs!Notes), "", Rs!Notes)
Else
    Text_Notes.Text = ""
End If
Rs.Close
End Sub

Private Sub Update_Data()
If Text1.Text = "" Then
    Exit Sub
End If
Dim Rs As New ADODB.Recordset
Rs.Open "Select * From Calendar Where Date like '" & Text1.Text & "' ", Cn, adOpenStatic, adLockOptimistic
If Rs.RecordCount > 0 Then
    Rs!Notes = IIf(Text_Notes.Text = "", Null, Text_Notes.Text)
    Rs.Update
Else
    Rs.AddNew
    Rs!Date = Text1.Text
    Rs!Notes = IIf(Text_Notes.Text = "", Null, Text_Notes.Text)
    Rs.Update
End If
Rs.Close
End Sub

Private Sub Command_Cancel_Click()
Unload Me
End Sub

Private Sub Command_Ok_Click()
Call Update_Data
End Sub

Private Sub Command_Report_Click()
Call Set_Report
Dim RPT As String
Dim F As File
Set F = Fs.GetFile(App.Path & "\Report.htm")
RPT = S_HTML
RPT = RPT & S_Body
RPT = RPT & S_Title
RPT = RPT & List1.Text & " „«Â " & Text_Year.Text
RPT = RPT & E_Title
'
RPT = RPT & S_Table
RPT = RPT & "<tr>"
'
RPT = RPT & S_Header_TD & " «—ÌŒ" & E_Header_TD
RPT = RPT & S_Header_TD & "„Ê÷Ê⁄" & E_Header_TD
'
RPT = RPT & "</tr>" & vbCrLf
Dim Rs As New ADODB.Recordset
Rs.Open "Select * From Calendar Where Date like '" & Text_Year.Text & "/" & Format(List1.ListIndex + 1, "00") & "/" & "%' And Notes is Not Null Order By Date", Cn, adOpenStatic, adLockOptimistic
If Rs.RecordCount > 0 Then
    Do While Not Rs.EOF
        RPT = RPT & "<tr>" & _
        S_Row_TD1 & Fa_Day(En_Date(Rs!Date)) & "</br>" & Chr(253) & Rs!Date & "<br>" & Chr(253) & En_Date(Rs!Date) & E_Row_TD & _
        S_Row_TD2 & Encode(IIf(IsNull(Rs!Notes), "", Rs!Notes)) & E_Row_TD & _
        "</tr>"
        Rs.MoveNext
    Loop
End If
Rs.Close
RPT = RPT & "</Table></body></thml>"
F.OpenAsTextStream(ForWriting).Write RPT
ShellExecute Me.hwnd, vbNullString, App.Path & "\Report.htm", vbNullString, App.Path & "\", 1
End Sub

Private Sub Command1_Click()
Text_Year.Text = Text_Year.Text + 1
Call Set_Date
End Sub

Private Sub Command2_Click()
Text_Year.Text = Text_Year.Text - 1
Call Set_Date
End Sub

Private Sub Form_Load()
Call Open_Cn
Text_Year = Left(Fa_Date(Date), 4)
Text_Today.Text = Fa_Day(Date) & " " & Fa_Date(Date)
List1.AddItem "›—Ê—œÌ‰"
List1.AddItem "«—œÌ»Â‘ "
List1.AddItem "Œ—œ«œ"
List1.AddItem " Ì—"
List1.AddItem "„—œ«œ"
List1.AddItem "‘Â—ÌÊ—"
List1.AddItem "„Â—"
List1.AddItem "¬»«‰"
List1.AddItem "¬–—"
List1.AddItem "œÌ"
List1.AddItem "»Â„‰"
List1.AddItem "«”›‰œ"
List1.ListIndex = CLng(Mid(Fa_Date(Date), 6, 2)) - 1
For K = 0 To Label2.Count - 1
    If Format(Label2(K).Caption, "00") = Mid(Fa_Date(Date), 9) Then
        Call Label2_Click(CInt(K))
    End If
Next
End Sub

Private Sub Label2_Click(Index As Integer)
For K = 0 To Label2.Count - 1
    Label2(K).BorderStyle = 0
    Label2(K).BackColor = 12632100
Next
Label2(Index).BorderStyle = 1
Label2(Index).BackColor = RGB(255, 255, 255)
Text1.Text = Text_Year.Text & "/" & Format(List1.ListIndex + 1, "00") & "/" & Format(Label2(Index).Caption, "00")
Call Get_Data
End Sub

Private Sub List1_Click()
Call Set_Date
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then
    Command_Ok.Enabled = False
    Text2.Text = ""
Else
    Command_Ok.Enabled = True
    Text2.Text = En_Date(Text1.Text)
End If
End Sub
