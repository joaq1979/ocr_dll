VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDemo 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox image2 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Left            =   5040
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.PictureBox image1 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Left            =   600
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   5040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblNombre 
      Caption         =   "Label1"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
