VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmADOSQLGenerate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Statement and ADO Code Generator"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   HelpContextID   =   10
   Icon            =   "frmADOSQLGenerate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9000
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdGenCode 
      Caption         =   "&Generate Code"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdShowCode 
      Caption         =   "Show Code &Window "
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdShowGrid 
      Caption         =   "&Show Grid"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdMakeSQLStmt 
      Caption         =   "&Edit SQL"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpenDB 
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Instructions"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   0
      Width           =   1095
   End
   Begin VB.CheckBox chkNoWhere 
      Caption         =   "SQL - ""Where"" Clause =""No"""
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      ToolTipText     =   "Invoke SQL ""Where"" Clause"
      Top             =   600
      WhatsThisHelpID =   10
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8520
      Top             =   480
   End
   Begin VB.TextBox txtCodeWindow 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2235
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   5280
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   8775
   End
   Begin VB.Frame fraGenCode 
      Caption         =   "Generate Code"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   6600
      TabIndex        =   22
      ToolTipText     =   "Pick the database provider you want to generate code with.  Note, in order to open an Access 2000 database, select Jet 4.0."
      Top             =   3960
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   2295
      Begin VB.OptionButton optCodeGen 
         Caption         =   "No - Show Grid"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   720
         WhatsThisHelpID =   10
         Width           =   1695
      End
      Begin VB.OptionButton optCodeGen 
         Caption         =   "Yes-Open Code Window"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   25
         ToolTipText     =   "Show Code Window"
         Top             =   240
         WhatsThisHelpID =   10
         Width           =   1815
      End
   End
   Begin VB.Frame fraGridOptions 
      Caption         =   "Grid Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   4613
      TabIndex        =   19
      ToolTipText     =   "Select a Grid Option"
      Top             =   3960
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   1335
      Begin VB.OptionButton optGridOptions 
         Caption         =   "No Grid"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "No Grid Code"
         Top             =   960
         Value           =   -1  'True
         WhatsThisHelpID =   10
         Width           =   1095
      End
      Begin VB.OptionButton optGridOptions 
         Caption         =   "DataGrid"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Code for Data Grid"
         Top             =   240
         WhatsThisHelpID =   10
         Width           =   1095
      End
      Begin VB.OptionButton optGridOptions 
         Caption         =   "FlexGrid"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Code For FlexGrid"
         Top             =   600
         WhatsThisHelpID =   10
         Width           =   1095
      End
   End
   Begin VB.Frame fraConnOptions 
      Caption         =   "cn/rs Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Select a Grid Option"
      Top             =   3960
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   2295
      Begin VB.OptionButton CnRs1 
         Caption         =   "cn/rs - Only"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Select Connection & Recordset Title Options"
         Top             =   240
         WhatsThisHelpID =   10
         Width           =   1695
      End
      Begin VB.OptionButton CnRs2 
         Caption         =   "connTable Name/"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Select Connection & Recordset Title Options"
         Top             =   720
         WhatsThisHelpID =   10
         Width           =   1935
      End
      Begin VB.Label lblrsTableName 
         Caption         =   "rsTableName"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   960
         WhatsThisHelpID =   10
         Width           =   1335
      End
   End
   Begin VB.Frame fraProvider 
      Caption         =   "Provider"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   2760
      TabIndex        =   12
      ToolTipText     =   "Pick the database provider you want to generate code with.  Note, in order to open an Access 2000 database, select Jet 4.0."
      Top             =   3960
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   1335
      Begin VB.OptionButton optProvider 
         Caption         =   "Jet 3.51"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         WhatsThisHelpID =   10
         Width           =   1095
      End
      Begin VB.OptionButton optProvider 
         Caption         =   "Jet 4.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Value           =   -1  'True
         WhatsThisHelpID =   10
         Width           =   975
      End
   End
   Begin VB.TextBox txtGetDB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   405
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "frmADOSQLGenerate.frx":0442
      ToolTipText     =   "Select a Data Base"
      Top             =   0
      WhatsThisHelpID =   10
      Width           =   4455
   End
   Begin VB.CheckBox chkWhere 
      Caption         =   "SQL - ""Where"" Clause = ""Yes"""
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "Invoke SQL ""Where"" Clause"
      Top             =   840
      WhatsThisHelpID =   10
      Width           =   3015
   End
   Begin VB.ListBox lstSQLMath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1260
      ItemData        =   "frmADOSQLGenerate.frx":045E
      Left            =   7080
      List            =   "frmADOSQLGenerate.frx":0460
      TabIndex        =   5
      ToolTipText     =   "Select SQL Math Operator"
      Top             =   1080
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid dbGrid 
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5953
      _Version        =   393216
      ForeColor       =   12583104
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   1215
      Left            =   4320
      TabIndex        =   3
      ToolTipText     =   "Click Node to Compare Values"
      Top             =   1080
      WhatsThisHelpID =   10
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2143
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2760
      WhatsThisHelpID =   10
      Width           =   8775
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   3600
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1230
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Select Field(s)"
      Top             =   1080
      WhatsThisHelpID =   10
      Width           =   1935
   End
   Begin VB.ListBox lstTables 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Select Table"
      Top             =   1080
      WhatsThisHelpID =   10
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOSQLGenerate.frx":0462
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOSQLGenerate.frx":0576
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOSQLGenerate.frx":068A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFieldsOpening 
      Caption         =   "DB - Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   30
      Top             =   840
      WhatsThisHelpID =   10
      Width           =   1215
   End
   Begin VB.Label lblOpeningTables 
      Caption         =   "DB - Tables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   840
      WhatsThisHelpID =   10
      Width           =   1215
   End
   Begin VB.Label lblClipCode 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Code is on the Clipboard"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   28
      Top             =   2400
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   3015
   End
   Begin VB.Label lblSQLOper 
      Caption         =   "SQL - Operator(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7320
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   1575
   End
   Begin VB.Label lblFieldsCount 
      Caption         =   "lblFieldsCount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   1935
   End
   Begin VB.Label lblRecordCount 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "lblRecordCount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   1320
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   1320
   End
   Begin VB.Label lblTableCount 
      Caption         =   "lblTableCount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   1935
   End
   Begin VB.Label lblSQLStatement 
      Caption         =   "SQL Statement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      WhatsThisHelpID =   10
      Width           =   1695
   End
End
Attribute VB_Name = "frmADOSQLGenerate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'*                     Project: SQL Statement and ADO Code Generator        *
'****************************************************************************
' Modules:        ContextIDs - Help File Module
'                 modtxtEffect - see About Form
'                 modWebEmail - send email
                  
' Description:   Sets up a basic ADO Database Code Template
'            :   Instructions for use:  It would be best to create a new folder,
'                then start a new VB Project, Name your Form and Project and
'                make a Reference to the ActiveX Data Objects Library. At your
'                option, add either a MSDataGrid or MSFlexGrid to the Form.
'                (Note: DataGrids are updatable but FlexGrids are not).
'
'                (Note): Selecting the MSDataGrid Option causes code for Update
'                and Delete Command Buttons to be generated so if you intend to
'                use these, you must add them to your Form.
'
'                Finally save Project in the new folder. Next, run this program in it's
'                "Exe" mode.  The code generated will be automatically placed
'                on the Clipboard.  Finally, just paste it into your app.
'
' ==========================================================================
' ====           Full Credit to Jerry Barnes for the idea, the          ====
' ====           original connection code and much of the other code    ====
' ====           found in his ADO Toolbox at Planetwide Software.       ====
' ==========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-JUL-2000  John P. Cunningham - johnpc@ids.net
'              Module created
' ***************************************************************************

' Make sure to open Project-References and select
' Microsoft ActiveX Data Object Library 2.0 or higher.

Option Explicit
 
Dim rsRecordset                     As ADODB.Recordset
Dim connConnection                  As ADODB.Connection
Dim mstrDatabasePath                As String
Dim mstrConnectionString            As String
Dim mstrProvider                    As String
Dim mstrTableName                   As String
Dim mstrSQL                         As String
Dim recCount                        As Integer
Dim mintRelative                    As Integer
Dim RetVal                          As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''
Dim DbFile                          As String
Dim Wildcard                        As Boolean
Dim mstrCheckForDatabase            As String
Dim mstrDataSource                  As String

Dim mstrRecordSet                   As String
Dim mstrDatabaseName                As String

Dim mstrConnectionObject            As String
Dim mstrRecordSetObject             As String
Dim mstrFieldName                   As String
Dim tmpDBstring As String
Dim mstrDataGrid                    As Boolean
Dim mstrFlexGrid                    As Boolean
Dim GenFlexGridCode                 As Boolean
Dim GenDBGridCode                   As Boolean
Const mstrAccessProvider351         As String = "Provider= Microsoft.Jet.OLEDB.3.51;"
Const mstrAccessProvider40          As String = "Provider=Microsoft.Jet.OLEDB.4.0;"
'Dim Clearit                         As Boolean
Const tvwFirst                      As Integer = 0     'tvwFirst
Const tvwLast                       As Integer = 1     'tvwLast
Const tvwNext                       As Integer = 2     'tvwNext
Const tvwPrevious                   As Integer = 3     'tvwPrevious
Const tvwChild                      As Integer = 4     'tvwChild
Private Sub WriteCode()
        
    '***********************************************************
    '               Declarations Code
    '***********************************************************
    
    txtCodeWindow = Space(27) & "'**** Form Level Declarations ****" & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "'*************************************************************************************************" _
        & vbCrLf
    txtCodeWindow = txtCodeWindow & "'  Be sure to add a Reference to Ms ActiveX Data Objects 2.x Library to Project" _
        & vbCrLf
      txtCodeWindow = txtCodeWindow & "'*************************************************************************************************" _
            & vbCrLf & vbCrLf
            
    If optGridOptions(0) Then
        
         txtCodeWindow = txtCodeWindow & "'     ********** Data Grid Option has been selected ********* " & vbCrLf
         txtCodeWindow = txtCodeWindow & "'     Add a Microsoft DataGrid Control to Project. " & vbCrLf
         txtCodeWindow = txtCodeWindow & "'     Add cmdUpdate and cmdDelete Command Buttons to Form" & vbCrLf
         txtCodeWindow = txtCodeWindow & "'     (This will allow Record(s) Updates and Delete(s) via the DataGrid," & vbCrLf
         txtCodeWindow = txtCodeWindow & "'     to the appropriate Access DB Table)" & vbCrLf
         
         txtCodeWindow = txtCodeWindow & "'********************************************************************************************" _
            & vbCrLf & vbCrLf
   End If
    
    txtCodeWindow = txtCodeWindow & "Dim DbFile As String" _
        & Space(30) & "'Name of DataBase" & vbCrLf
    txtCodeWindow = txtCodeWindow & "Dim " & mstrConnectionObject _
        & " as ADODB.Connection" & Space(15) & " 'Connect to the ADO Data Type" & vbCrLf
        
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   If CnRs1 = True Then
        txtCodeWindow = txtCodeWindow & "Dim rs as ADODB.Recordset" _
            & Space(19) & "'Record Source Name" & vbCrLf
        
    Else
    
        txtCodeWindow = txtCodeWindow & "Dim " & mstrRecordSetObject & " as ADODB.Recordset" _
            & Space(18) & "'Record Source Name" & vbCrLf
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    txtCodeWindow = txtCodeWindow & "Dim SQLstmt as String" _
        & Space(28) & "'SQL Statement String(s)" & vbCrLf & vbCrLf
     
    
    '***********************************************************
    '                 Form Load Code
    '***********************************************************
    txtCodeWindow = txtCodeWindow & "Private Sub Form_Load()" & vbCrLf
'*** Code added by HelpWriter ***
    SetAppHelp Me.hWnd
'***********************************

    txtCodeWindow = txtCodeWindow & "     Open_" & mstrConnectionObject & vbCrLf
    
    '*****************************************************************
    
    'Generate DataGrid or FlexGrid or No Grid String
    If mstrDataGrid Then
    
        If CnRs1 = True Then
             txtCodeWindow = txtCodeWindow & "     Set DataGrid1.DataSource = rs" & vbCrLf
        Else
        
            txtCodeWindow = txtCodeWindow & "     Set DataGrid1.DataSource = " _
            & mstrRecordSetObject & vbCrLf
        
        End If
        
        GenDBGridCode = True
        
        ElseIf mstrFlexGrid Then
            txtCodeWindow = txtCodeWindow & "     Call ShowFlexGrid" & vbCrLf
            GenFlexGridCode = True
   ' Else
        'No Grid Selected
    End If
        
    txtCodeWindow = txtCodeWindow & "End Sub" & vbCrLf & vbCrLf
    
    '***********************************************************
    '                 Open Subroutine Code
    '***********************************************************
    
    txtCodeWindow = txtCodeWindow & "Private Sub Open_" _
        & mstrConnectionObject & " ()" & vbCrLf
    txtCodeWindow = txtCodeWindow & "'     Set the Database Applicable Path" & vbCrLf
    txtCodeWindow = txtCodeWindow & "      DbFile = App.Path " & Chr(38) _
        & " " & Chr(34) _
        & Chr(92) & DbFile & Chr(34) & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "'      Establish the Connection" & vbCrLf
    txtCodeWindow = txtCodeWindow & "       Set " & mstrConnectionObject _
        & "= New ADODB.Connection" & vbCrLf
    txtCodeWindow = txtCodeWindow & "       " & mstrConnectionObject _
        & ".CursorLocation = adUseClient" & vbCrLf
    txtCodeWindow = txtCodeWindow & "       " & mstrConnectionObject _
        & ".ConnectionString = _" & vbCrLf
       
    txtCodeWindow = txtCodeWindow & "             " & Chr(34) & mstrProvider & Chr(34) & Chr(32) & Chr(38) _
            & Chr(32) & Chr(95) & vbCrLf
    
    txtCodeWindow = txtCodeWindow & "             " & Chr(34) & "Data Source=" & Chr(34) _
        & Chr(32) & Chr(38) & Chr(32) & "DbFile" _
        & Chr(32) & Chr(38) & Chr(32) & Chr(34) & Chr(59) & Chr(34) & Chr(32) _
        & Chr(38) & Chr(32) & Chr(95) & vbCrLf
    
    txtCodeWindow = txtCodeWindow & "             " & Chr(34) & "Persist Security Info=False" _
        & Chr(34) & vbCrLf & vbCrLf
        
    txtCodeWindow = txtCodeWindow & "'      Open the Connection" & vbCrLf
    txtCodeWindow = txtCodeWindow & "      " & mstrConnectionObject & ".Open" & vbCrLf & vbCrLf
    
    txtCodeWindow = txtCodeWindow & "'      Once this Connection is opened, it can" & vbCrLf
    txtCodeWindow = txtCodeWindow & "'      be used throughout the application" & vbCrLf & vbCrLf
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Wildcard = False Then
    txtCodeWindow = txtCodeWindow & "'******************************************************************************" & vbCrLf
    txtCodeWindow = txtCodeWindow & "'     The following line is the Master SQL Statement and" & vbCrLf
    txtCodeWindow = txtCodeWindow & "'     it is remarked out to show the actual SQL Statement Selected." & vbCrLf
    txtCodeWindow = txtCodeWindow & "'     Simply switch the remark to the actual SQL Statement" & vbCrLf
    txtCodeWindow = txtCodeWindow & "'     to use the Master SQL Statement" & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "'     SQLstmt = " & Chr(34) & mstrSQL & Chr(34) & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "      SQLstmt = " & Chr(34) & txtSQL & Chr(34) & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "'******************************************************************************" & vbCrLf
 Else
    txtCodeWindow = txtCodeWindow & "      SQLstmt = " & Chr(34) & mstrSQL & Chr(34) & vbCrLf & vbCrLf
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If CnRs1 = True Then
    txtCodeWindow = txtCodeWindow & "'      Get the Records" & vbCrLf
    txtCodeWindow = txtCodeWindow & "      Set rs = New ADODB.Recordset" & vbCrLf
    txtCodeWindow = txtCodeWindow & "      rs.Open SQLstmt, " _
        & mstrConnectionObject & ", adOpenStatic, adLockOptimistic, " _
        & Chr(95) & vbCrLf & Space(10) & " adCmdText" & vbCrLf & vbCrLf
       
    Else
    
    txtCodeWindow = txtCodeWindow & "'      Get the Records" & vbCrLf
    txtCodeWindow = txtCodeWindow & "      Set " & mstrRecordSetObject & " = New ADODB.Recordset" & vbCrLf
    txtCodeWindow = txtCodeWindow & "      " & mstrRecordSetObject & ".Open SQLstmt, " _
        & mstrConnectionObject & ", adOpenStatic, adLockOptimistic, " _
        & Chr(95) & vbCrLf & Space(10) & " adCmdText" & vbCrLf & vbCrLf
       
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    txtCodeWindow.Text = txtCodeWindow & "End Sub" & vbCrLf & vbCrLf
    
    '*************************************************************************
    '             Closing Subroutine Code
    '*************************************************************************
    
    txtCodeWindow.Text = txtCodeWindow & "Private Sub Close_" & mstrConnectionObject & " ()" & vbCrLf
    txtCodeWindow.Text = txtCodeWindow & "     " & mstrConnectionObject & ".Close" & vbCrLf
    txtCodeWindow.Text = txtCodeWindow & "     Set " & mstrConnectionObject & " = Nothing" & vbCrLf
    txtCodeWindow.Text = txtCodeWindow & "End Sub" & vbCrLf & vbCrLf
    
    If GenFlexGridCode Then
        GenerateFlexGridCode
    ElseIf GenDBGridCode Then
        GenerateDBGridCode 'reserved for future to add Update & Delete Button Code
    End If
    
End Sub


Private Sub AddField()
    
    Dim strHoldFieldName As String
    Dim strHoldTableName As String
    Dim intIndex As Integer
    
On Error GoTo HandleErrors

    strHoldFieldName = lstFields.List(lstFields.ListIndex)
    strHoldTableName = lstTables.List(lstTables.ListIndex)
    
    mintRelative = -1
    
    For intIndex = 1 To tv.Nodes.Count
        If tv.Nodes.Item(intIndex).Text = strHoldTableName Then
            mintRelative = intIndex
        End If
    Next intIndex
    'Scan the tables to determine which node to add the child
    'to.
    
    If lstFields.ListIndex <> -1 And mintRelative <> -1 Then
    ' Make sure that a field is selected.
    
        tv.Nodes.Add mintRelative, tvwChild, lstTables.List(lstTables.ListIndex) & strHoldFieldName, strHoldFieldName, 2, 3
        ' Add a child to a node on the tree.
        '   Relative:  The relative will be the node that
        '              contains the table that the field
        '              comes from.
        '   Relation:  The new node is a child.
        '   Key:       The key will be the table name and
        '              field name combined.
        '   Text:      The text will be the name of the
        '              field.
            
    End If
    'If the whole table is selected, i.e. * then
        'unable the ListFields List Box because the
        'SQL Statement Generated will show the entire
        'Data Base Table
        
    If lstFields = "*" Then
            Wildcard = True
            Me.Height = 7960
            chkWhere.Enabled = False
            lstFields.Enabled = True 'False
           'cmdClear.Visible = True
            'cmdExit1.Visible = False
            Call MakeSQLStmt
            cmdShowGrid_Click
            cmdMakeSQLStmt.Enabled = False
            SetButtonForeColor cmdMakeSQLStmt, &H808080
            
            cmdShowGrid.Enabled = False
            SetButtonForeColor cmdShowGrid, &H808080
           
            cmdGenCode.Enabled = True
            SetButtonForeColor cmdShowCode, &H808080
            cmdShowCode.Enabled = False
    Else
            chkNoWhere.Enabled = True
            chkWhere.Enabled = True
    End If
    
    For intIndex = 1 To tv.Nodes.Count
        
    'This refreshes the tree so that new nodes become visible.
           tv.Nodes(intIndex).EnsureVisible
    Next
   
        
    Exit Sub
    'Everything went okay so leave.
    
HandleErrors:
    
    MsgBox Err.Number & vbCrLf & Err.Description
        
End Sub

Private Sub AddTable()

    Dim strHoldTableName As String
    
    On Error GoTo HandleErrors
    ' Make sure that a table is selected.
    If lstTables.ListIndex <> -1 Then
    
        strHoldTableName = lstTables.List(lstTables.ListIndex)
        tv.Nodes.Add , tvwNext, strHoldTableName, strHoldTableName, 1
        ' Add a node to the tree.
        '   Relatative:  There is no relative paramenter.
        '   Relation:    The relation is next in line.
        '   Key:         The key is the name of the table.
        '   Text:        The text displayed in the tree
        '                will be the name of the table.
    
    End If

    Exit Sub
    'Everything went fine so leave.
    
HandleErrors:
    MsgBox Err.Number & vbCrLf & Err.Description

End Sub

Private Sub GenerateDBGridCode()
'   Generate Code for Update Command Button
   
    txtCodeWindow = txtCodeWindow & "Private Sub cmdUpdate_Click()" & vbCrLf
    txtCodeWindow = txtCodeWindow & "'     Remember to add the cmdUpdate Command Button to your Form" & vbCrLf
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If CnRs1 = True Then
        txtCodeWindow = txtCodeWindow & "     rs.Update" & vbCrLf
    Else
        txtCodeWindow = txtCodeWindow & "     " & mstrRecordSetObject & ".Update" & vbCrLf
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    txtCodeWindow = txtCodeWindow & "End Sub" & vbCrLf & vbCrLf
       
    
'   Generate Code for Delete Command Button
    txtCodeWindow = txtCodeWindow & "Private Sub cmdDelete_Click()" & vbCrLf
    txtCodeWindow = txtCodeWindow & "Dim intResponse As Integer" & vbCrLf
    txtCodeWindow = txtCodeWindow & "'      Remember to add the cmdDelete Command Button to your Form" & vbCrLf
    txtCodeWindow = txtCodeWindow & "       Beep" & vbCrLf
    

 txtCodeWindow = txtCodeWindow & "     intResponse = MsgBox(" & Chr(34) _
        & "Delete the Current Record" & Chr(34) & Chr(44) & " " & Chr(95) & vbCrLf
 txtCodeWindow = txtCodeWindow & Space(10) & "vbYesNo + vbQuestion" & Chr(44) & " " & Chr(34) _
    & "Delete Record" & Chr(34) & Chr(41) & vbCrLf & vbCrLf
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  If CnRs1 = True Then
    txtCodeWindow = txtCodeWindow & Space(5) & "If intResponse = vbYes Then" & vbCrLf
    txtCodeWindow = txtCodeWindow & Space(10) & "If Not rs.EditMode = adEditAdd Then " & vbCrLf
    txtCodeWindow = txtCodeWindow & Space(15) & "rs.Delete " & vbCrLf
 Else
    txtCodeWindow = txtCodeWindow & Space(5) & "If intResponse = vbYes Then" & vbCrLf
    txtCodeWindow = txtCodeWindow & Space(10) & "If Not " & mstrRecordSetObject & ".EditMode = adEditAdd Then " & vbCrLf
    txtCodeWindow = txtCodeWindow & Space(15) & mstrRecordSetObject & ".Delete " & vbCrLf
 End If
 
 txtCodeWindow = txtCodeWindow & Space(10) & "End If" & vbCrLf
 txtCodeWindow = txtCodeWindow & Space(5) & "End If" & vbCrLf
 txtCodeWindow = txtCodeWindow & "End Sub" & vbCrLf & vbCrLf
 
End Sub
Private Sub GenerateFlexGridCode()
'********************************************************************
'                     Load FlexGrid Code
'********************************************************************
    txtCodeWindow = txtCodeWindow & "Private Sub ShowFlexGrid()" & vbCrLf
    txtCodeWindow = txtCodeWindow & "Dim c as Integer" & vbCrLf
    txtCodeWindow = txtCodeWindow & "Dim flxgd_row as Integer" & vbCrLf
    txtCodeWindow = txtCodeWindow & "Dim field_wid as Integer" & vbCrLf
    txtCodeWindow = txtCodeWindow & "     ' Use one fixed row and no fixed columns " & vbCrLf
    txtCodeWindow = txtCodeWindow & "     MSFlexGrid1.Rows = 2" & vbCrLf
    txtCodeWindow = txtCodeWindow & "     MSFlexGrid1.FixedRows = 1" & vbCrLf
    txtCodeWindow = txtCodeWindow & "     MSFlexGrid1.FixedCols = 0" & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "     ' Display column headers " & vbCrLf
    txtCodeWindow = txtCodeWindow & "     MSFlexGrid1.Rows = 1 " & vbCrLf
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If CnRs1 = True Then mstrRecordSetObject = "rs"
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    txtCodeWindow = txtCodeWindow & "     MSFlexGrid1.Cols = " & mstrRecordSetObject _
        & ".Fields.Count" & vbCrLf
    txtCodeWindow = txtCodeWindow & "     ReDim col_wid(0 To " & mstrRecordSetObject _
        & ".Fields.Count - 1)" & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "     For c = 0 to (" & mstrRecordSetObject _
        & ".Fields.Count - 1)" & vbCrLf
    txtCodeWindow = txtCodeWindow & "               MSFlexGrid1.TextMatrix(0, c) = " _
        & mstrRecordSetObject & ".Fields(c).Name" & vbCrLf
    txtCodeWindow = txtCodeWindow & "               col_wid(c) = TextWidth(" _
        & mstrRecordSetObject & ".Fields(c).Name)" & vbCrLf
    txtCodeWindow = txtCodeWindow & "     Next c" & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "     'Display the values for each row" & vbCrLf
    txtCodeWindow = txtCodeWindow & "     flxgd_row = 1" & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "     Do While Not " & mstrRecordSetObject & ".EOF" & vbCrLf
    txtCodeWindow = txtCodeWindow & "          MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1" _
        & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "          For c = 0 To (" & mstrRecordSetObject _
        & ".Fields.Count - 1)" & vbCrLf & vbCrLf
            
    txtCodeWindow = txtCodeWindow & "               MSFlexGrid1.TextMatrix(flxgd_row, c) = _" & vbCrLf
    
    
    txtCodeWindow = txtCodeWindow & "               Format(" & mstrRecordSetObject _
        & ".Fields(c).Value, " & Chr(34) & Chr(46) & Space(6) & Chr(34) _
        & Chr(41) & vbCrLf & vbCrLf
    
    txtCodeWindow = txtCodeWindow & "               ' See how big the value is" & vbCrLf
    txtCodeWindow = txtCodeWindow & "               field_wid = TextWidth(" & mstrRecordSetObject _
        & ".Fields(c).Value) " & vbCrLf
    txtCodeWindow = txtCodeWindow & "               If col_wid(c) < field_wid Then col_wid(c) = field_wid" & vbCrLf & vbCrLf
    txtCodeWindow = txtCodeWindow & "          Next c" & vbCrLf & vbCrLf
        
    txtCodeWindow = txtCodeWindow & "          " & mstrRecordSetObject & ".MoveNext" & vbCrLf
    txtCodeWindow = txtCodeWindow & "          flxgd_row = flxgd_row + 1" & vbCrLf & vbCrLf
        
    txtCodeWindow = txtCodeWindow & "     Loop" & vbCrLf & vbCrLf
    
    txtCodeWindow = txtCodeWindow & "     End Sub"
    
End Sub


Private Sub chkNoWhere_Click()

            Me.Height = 7960
            chkWhere.Enabled = False
            lstFields.Enabled = True 'False
            'cmdClear.Visible = True
            'cmdExit1.Visible = False
            Call MakeSQLStmt
            cmdShowGrid_Click
            cmdMakeSQLStmt.Enabled = False
            cmdShowGrid.Enabled = False
            cmdShowCode.Enabled = False
            SetButtonForeColor cmdShowCode, &H808080
    
End Sub

Private Sub chkWhere_Click()
chkNoWhere.Enabled = False
    Call MakeSQLStmt
    txtSQL = txtSQL & " Where"
    
    tv.Enabled = True
    lstFields.Enabled = False
    cmdMakeSQLStmt.Enabled = True
    lblSQLOper.Visible = True
    lstSQLMath.Visible = True
    cmdShowCode.Enabled = False
            SetButtonForeColor cmdShowCode, &H808080
    
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdClear_Click()
Unload Me
Me.Show
Me.Top = 200

cmdClear.BackColor = &H8000000F
 SetButtonForeColor cmdClear, &HC00000
    
End Sub

Private Sub cmdExit_Click()

    Unload Me
   
End Sub

Private Sub MakeSQLStmt()
   Dim strField As String
    Dim strTable As String
    Dim intIndex As Integer
    Dim intCount As Integer
    

    intCount = 0
    For intIndex = 1 To tv.Nodes.Count
        If tv.Nodes(intIndex).Children <> 0 Then
            If intCount > 0 Then
                strTable = strTable & ", "
            End If
            strTable = strTable & "[" & tv.Nodes(intIndex).Text & "]"
            intCount = intCount + 1
        End If
    Next intIndex
    
    
    intCount = 0
    For intIndex = 1 To tv.Nodes.Count
        If tv.Nodes(intIndex).Children = 0 Then
            If intCount > 0 Then
                strField = strField & ", "
            End If
            strField = strField & "[" & tv.Nodes(intIndex).Text & "]"
            If strField = "[*]" Then
                strField = "*"
            End If
            intCount = intCount + 1
        End If
    Next intIndex
  
    txtSQL.Text = "SELECT " & strField & " FROM " & strTable
    
    cmdShowGrid.Enabled = True
    'Me.Height = 5565
End Sub

Private Sub LoadGrid()
            
    On Error GoTo HandleErrors:
        
        dbGrid.Visible = True
    Set rsRecordset = New ADODB.Recordset
    rsRecordset.CursorType = adUseClient
    rsRecordset.LockType = adLockPessimistic
    rsRecordset.Source = txtSQL.Text 'mstrSQL
    rsRecordset.ActiveConnection = connConnection
    rsRecordset.Open
    'Open the recordset that was generated.
    
    Set dbGrid.DataSource = rsRecordset
   ' dbGrid.Visible = True
    'View the generated data
    
    Exit Sub
HandleErrors:
    MsgBox "An invalid attempt has been made to open a database." & _
        "  This action has been cancelled.  Please check your SQL" & _
        " statement", vbOKOnly, "Error"
    'Call cmdClear_Click
    Exit Sub
End Sub



Private Sub cmdGenCode_Click()

    dbGrid.Visible = False
    fraGenCode.Visible = True
    fraProvider.Visible = True
    fraConnOptions.Visible = True
    fraGridOptions.Visible = True
    fraGridOptions.Enabled = False
    
    cmdShowGrid.Enabled = False
    SetButtonForeColor cmdShowGrid, &H808080
    
    cmdGenCode.Enabled = False
    SetButtonForeColor cmdGenCode, &H808080
    
    
End Sub

Private Sub cmdHelp_Click()

    ShowHelpTopic Hlp_SQL_Statement

End Sub

Private Sub cmdMakeSQLStmt_Click()
Call MakeSQLStmt
End Sub

Private Sub cmdOpenDB_Click()

    
    Dim strCheckForDatabase As String

    On Error GoTo HandleErrors
    
     dlgCommon.DialogTitle = "Pick A Database"
    'Give the file selection window a title.
    
    dlgCommon.InitDir = App.Path
    'The file selection window will start in the
    'applications directory.
    
    'Allow the user to view only Access files.
    dlgCommon.Filter = "Access Databases (*.mdb)|*.mdb|" & _
                       "All Files (*.*)|*.*"
    dlgCommon.ShowOpen
    'Open the file selection window.
    
    strCheckForDatabase = Right(dlgCommon.FileName, 4)
    'Select the last four letters of the file selected.
    
    Select Case strCheckForDatabase
       Case vbNullString
            'Do not allow empty strings.
            Exit Sub
            
        Case ".mdb"
             'Assign the chosen file to the path string.
             mstrDatabasePath = dlgCommon.FileName
             'Do not allow the user to select another DB until
            'clear is clicked.
             cmdOpenDB.Enabled = False
             
    End Select
    
   'Make the connection string with the source and selected file.
    mstrConnectionString = mstrProvider & "Data Source=" & _
        mstrDatabasePath
    
    'Open the connection with the selected db.
    Set connConnection = New ADODB.Connection
    connConnection.CursorLocation = adUseClient
    connConnection.Open mstrConnectionString
    
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set rsRecordset = connConnection.OpenSchema(adSchemaTables)
    
    recCount = 0
    Do Until rsRecordset.EOF
        If UCase(Left(rsRecordset!Table_Name, 4)) <> "MSYS" Then
            lstTables.AddItem rsRecordset!Table_Name
             recCount = recCount + 1
        End If
        rsRecordset.MoveNext
   Loop
   
    lblTableCount.Visible = True
    lblTableCount = "(" & recCount & ")" & " - Tables/Queries"
    txtGetDB = mstrDatabasePath
    
    tmpDBstring = txtGetDB
    lblOpeningTables.Visible = False
    Exit Sub
   
HandleErrors:
    MsgBox "Error opening database.  Please try again.  Remember to select" & _
        " the appropriate provider", vbOKCancel, "Error"
    cmdOpenDB.Enabled = True
End Sub

Private Sub cmdShowCode_Click()
txtCodeWindow.Visible = True
End Sub

Private Sub cmdShowGrid_Click()

    txtCodeWindow.Visible = False
    Call LoadGrid
    
End Sub

Private Sub CnRs1_Click()
    fraGridOptions.Enabled = True
    fraGenCode.Enabled = True
    optCodeGen(0).Enabled = True
    optCodeGen(1).Enabled = True
End Sub

Private Sub CnRs2_Click()
    fraGridOptions.Enabled = True
    fraGenCode.Enabled = True
    optCodeGen(0).Enabled = True
    optCodeGen(1).Enabled = True
End Sub



Private Sub Form_Load()
'*** Code added by HelpWriter ***
    SetAppHelp Me.hWnd
'***********************************

    Me.Height = 3585

 'Locate Form on the Screen
Me.Top = 200
Me.Move (Screen.Width - Me.Width) / 2

    mstrProvider = mstrAccessProvider40
    'Fill the SQL Math List Box
    With lstSQLMath
        '.AddItem "AS "
        .AddItem "BETWEEN "
        .AddItem "AND "
        '.AddItem "ORDER BY "
        '.AddItem "GROUB BY "
        '.AddItem "IN "
        '.AddItem "LIKE"
        .AddItem "< "
        .AddItem "> "
        .AddItem "= "
        .AddItem "<= "
        .AddItem ">= "
    End With
    SetButtonForeColor cmdAbout, &HC00000
    SetButtonForeColor cmdHelp, &HC00000
    SetButtonForeColor cmdClear, &HC00000
    
    SetButtonForeColor cmdMakeSQLStmt, &HFF&
    SetButtonForeColor cmdShowGrid, &HC000C0
    
    SetButtonForeColor cmdShowCode, &HC0FFFF
    
    SetButtonForeColor cmdGenCode, &HFF0000
    SetButtonForeColor cmdExit, &H8000&
    End Sub

Private Sub lstSQLMath_Click()
Dim lstItem As String

Me.Height = 7960

lstItem = lstSQLMath.Text


txtSQL = txtSQL & " " & lstItem
txtSQL.SetFocus
End Sub

Private Sub lstFields_Click()
    Call AddField
   
End Sub

Private Sub lstTables_Click()
    Dim intLoop, intLen     As Integer
    Dim strHoldTableName    As String
    Dim strTemp             As String
    Dim strtest             As String
    
    lstFields.Clear 'Clear the list.
    cmdOpenDB.Enabled = False
    'Get the name of the table selected.
    mstrTableName = "[" & lstTables.List(lstTables.ListIndex) & "]"
    
    'Add the wildcard character.
    Set rsRecordset = New ADODB.Recordset
    Set rsRecordset = _
        connConnection.Execute("Select * From [" & lstTables.List(lstTables.ListIndex) & "]", 1, 1)
    
    If rsRecordset.RecordCount <> 0 Then
        lstFields.AddItem "*"
    End If
    
    'Get the names & number of all fields within the selected table.
    For intLoop = 0 To rsRecordset.Fields.Count - 1
        lstFields.AddItem rsRecordset.Fields(intLoop).Name
       
    Next
     'Get the names & number of all fields within the selected table.
     lblFieldsCount.Visible = True
     lblFieldsCount = "(" & rsRecordset.Fields.Count & ") - Fields/Columns"
    ' Generate the name of the table for the code.
    mstrTableName = Mid(mstrTableName, 2, Len(mstrTableName) - 2)
    
    strTemp = "rs" & mstrTableName
    
   intLoop = 1
    strHoldTableName = ""
    intLen = Len(strTemp) + 1
    Do Until intLoop = intLen
        intLoop = intLoop + 1
        strtest = Left(strTemp, 1)
        If strtest = " " Then
            strHoldTableName = strHoldTableName & "_"
        Else
            strHoldTableName = strHoldTableName & Left(strTemp, 1)
        End If
        strTemp = Right(strTemp, Len(strTemp) - 1)
    Loop
    'List number of records in table
   lblRecordCount.Visible = True
   lblRecordCount = "Number of Records in " _
    & mstrTableName & " = " & rsRecordset.RecordCount
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate the name of the table for the code.
    mstrRecordSetObject = strHoldTableName
    mstrTableName = "[" & mstrTableName & "]"
      
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Generate the SQL Source string.
    mstrSQL = "SELECT * FROM " & mstrTableName
        Call AddTable
    lstTables.Enabled = False
    lblFieldsOpening.Visible = False
End Sub



Private Sub optCodeGen_Click(Index As Integer)

    If CnRs1 = False And CnRs2 = False Then
        For Index = 0 To 1
            optCodeGen(Index).Value = False
        Next
            RetVal = MsgBox("You must select a Connection / Recordset Option", 48, "SQL / ADO Code Generator")
            Exit Sub
    End If

    cmdShowCode.Enabled = True
    SetButtonForeColor cmdShowCode, &HC0FFFF
    
    cmdShowGrid.Enabled = True
    SetButtonForeColor cmdShowGrid, &HC000C0
    
    cmdGenCode.Enabled = False
    SetButtonForeColor cmdGenCode, &H808080
    
    fraConnOptions.Visible = False
    fraProvider.Visible = False
    fraGridOptions.Visible = False
    fraGenCode.Visible = False
    dbGrid.Visible = False
    txtCodeWindow.Top = 3960
    txtCodeWindow.Height = 3600
    
Select Case Index
    Case 0
    Dim intI As Integer
    Dim strHold As String
    Dim strtest As String
    
    cmdGenCode.Enabled = False
    'Do not allow code generation again until clear is selected.
    
    mstrDatabaseName = Left(mstrDatabasePath, Len(mstrDatabasePath) - 4)
    strHold = ""
    strtest = ""
    
    Do Until strtest = "\"
        strHold = Right(mstrDatabaseName, 1) & strHold
        strtest = Right(mstrDatabaseName, 1)
        mstrDatabaseName = Left(mstrDatabaseName, Len(mstrDatabaseName) - 1)
       
    Loop
    'Cut away the string until the database name is left.
    If CnRs1.Value = True Then
         mstrConnectionObject = "cn" '& Right(strHold, Len(strHold) - 1)
    Else
        mstrConnectionObject = "conn" & Right(strHold, Len(strHold) - 1)
    End If
    DbFile = Right(strHold, Len(strHold) - 1) + ".mdb"

  
    'Generate the code.
    Call WriteCode
    
    
    'Call LoadGrid
    'View the result of the data.
    cmdMakeSQLStmt.Enabled = False
    
    
    'Do not allow edits after generation.
    Clipboard.Clear


    txtCodeWindow.Visible = True
        
    Case Else
        
        fraGenCode.Visible = False
        fraProvider.Visible = False
        fraConnOptions.Visible = False
        fraGridOptions.Visible = False
        txtCodeWindow.Visible = False
        dbGrid.Visible = True
        Exit Sub
        
End Select

 Clipboard.Clear

'Select Text in txtBox & copy to clipboard
  Clipboard.SetText vbCrLf & vbCrLf & txtCodeWindow, vbCFText
'''''''''''''''''''''''''''''''''''''''''''
     lblClipCode.Visible = True
End Sub

Private Sub optGridOptions_Click(Index As Integer)

    Select Case Index
        Case 0
            mstrDataGrid = True
        
        Case 1
            mstrFlexGrid = True
        
        Case Else
            mstrDataGrid = False
            mstrFlexGrid = False
        
    End Select
End Sub

Private Sub optProvider_Click(Index As Integer)
 
    Select Case Index
        Case 0
            mstrProvider = mstrAccessProvider351
        Case 1
            mstrProvider = mstrAccessProvider40
    End Select
    
End Sub

Private Sub Timer1_Timer()
    
   If lblClipCode.BackColor = &HC0C0C0 Then 'Gray
      lblClipCode.BackColor = &HFF000 'Green
      lblClipCode.ForeColor = &HFF&   'Red
   Else
      lblClipCode.BackColor = &HC0C0C0 'Gray
      lblClipCode.ForeColor = &HFFFF&  'Yellow
   End If
End Sub

Private Sub tv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


  txtSQL = txtSQL & " " & tv.SelectedItem
  
End Sub

Private Sub txtSQL_Change()
    cmdShowGrid.Enabled = True
End Sub

Private Sub txtSQL_GotFocus()

    txtSQL.SelStart = Len(txtSQL)
    
End Sub

Private Sub txtSQL_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    
  
    cmdShowGrid_Click
    SendKeys "{TAB}"
End If
cmdGenCode.Enabled = True

End Sub
Sub Form_Unload(Cancel As Integer)
UnsetButtonForeColor cmdClear

'*** Code added by HelpWriter ***
    QuitHelp
'***********************************

End Sub
