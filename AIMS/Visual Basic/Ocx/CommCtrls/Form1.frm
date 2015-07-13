VERSION 5.00
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.0#0"; "CommCtrls.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin CommCtrls.mskDat mskDat1 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
   End
   Begin CommCtrls.CtxtBox CtxtBox1 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommCtrls.ItxtBox ItxtBox1 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommCtrls.NTxtBox NTxtBox1 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommCtrls.GujTxtBox GujTxtBox1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GUJAFONT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
