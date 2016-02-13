VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5B73778E-352B-11D9-91C4-40B155C10000}#7.1#0"; "CommCtrls.ocx"
Begin VB.Form frmPosGui 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame frmHeader 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   195
      TabIndex        =   52
      Top             =   50
      Width           =   1575
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   54
         Top             =   120
         Width           =   795
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   285
         TabIndex        =   53
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.Frame frmMainContainer 
      Height          =   11055
      Left            =   195
      TabIndex        =   1
      Top             =   400
      Width           =   20220
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   7935
         TabIndex        =   5
         Top             =   195
         Width           =   8535
         Begin VB.CommandButton cmdItemNevigate 
            Caption         =   ">"
            Height          =   1550
            Index           =   1
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   5280
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   240
            Width           =   1500
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   1
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   240
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   2
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   3
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   4
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   240
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   6
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   7
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   8
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   9
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   10
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   11
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   12
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   13
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   14
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   15
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   5280
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   16
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   5280
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   17
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   5280
            Width           =   1550
         End
         Begin VB.CommandButton cmdItems 
            Caption         =   "cmdItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Index           =   18
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   5280
            Width           =   1550
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5940
         Left            =   30
         TabIndex        =   2
         Top             =   200
         Width           =   7905
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshTicket 
            Height          =   5055
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   8916
            _Version        =   393216
            ForeColor       =   8388608
            Rows            =   1
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorSel    =   16308668
            ForeColorSel    =   8388608
            BackColorBkg    =   14482428
            FocusRect       =   0
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   1
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHeader 
            Height          =   615
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   1085
            _Version        =   393216
            ForeColor       =   8388608
            Cols            =   1
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorSel    =   16308668
            ForeColorSel    =   8388608
            BackColorBkg    =   14482428
            FocusRect       =   0
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   1
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   30
         TabIndex        =   26
         Top             =   6045
         Width           =   7905
         Begin VB.CommandButton Command1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            TabIndex        =   55
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtTotalAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   200
            Width           =   1215
         End
         Begin VB.TextBox txtTotalQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   200
            Width           =   975
         End
         Begin VB.TextBox txtTotalDiscount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   200
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   30
         TabIndex        =   30
         Top             =   6690
         Width           =   7905
         Begin VB.CommandButton cmdPay 
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   240
            Width           =   7575
         End
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3330
         Left            =   30
         TabIndex        =   32
         Top             =   7695
         Width           =   7905
         Begin CommCtrls.CtxtBox txtBarcode 
            Height          =   855
            Left            =   120
            TabIndex        =   0
            Top             =   2280
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   1508
            Appearance      =   0
            MaxLength       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AllowNull       =   -1  'True
         End
         Begin VB.Label lblBarcodeMsg 
            Alignment       =   1  'Right Justify
            Caption         =   "Item missing!"
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4320
            TabIndex        =   56
            Top             =   1800
            Width           =   3255
         End
         Begin VB.Label lblChangeAmount 
            Alignment       =   2  'Center
            Caption         =   "Change Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   1155
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   5895
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         Left            =   7935
         TabIndex        =   34
         Top             =   7080
         Width           =   8535
         Begin VB.CommandButton cmdCategory 
            Caption         =   "cmdCategory"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   8
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1375
            Width           =   1550
         End
         Begin VB.CommandButton cmdCategory 
            Caption         =   "cmdCategory"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   7
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1375
            Width           =   1550
         End
         Begin VB.CommandButton cmdCategory 
            Caption         =   "cmdCategory"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   6
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1375
            Width           =   1550
         End
         Begin VB.CommandButton cmdCategory 
            Caption         =   "cmdCategory"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1375
            Width           =   1550
         End
         Begin VB.CommandButton cmdCategoryNeviagte 
            Caption         =   ">"
            Height          =   1080
            Index           =   0
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1375
            Width           =   1550
         End
         Begin VB.CommandButton cmdCategory 
            Caption         =   "cmdCategory"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   4
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   200
            Width           =   1550
         End
         Begin VB.CommandButton cmdCategory 
            Caption         =   "cmdCategory"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   3
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   200
            Width           =   1550
         End
         Begin VB.CommandButton cmdCategory 
            Caption         =   "cmdCategory"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   2
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   200
            Width           =   1550
         End
         Begin VB.CommandButton cmdCategory 
            Caption         =   "cmdCategory"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   1
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   200
            Width           =   1550
         End
         Begin VB.CommandButton cmdCategory 
            Caption         =   "cmdCategory"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   200
            Width           =   1550
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   7935
         TabIndex        =   45
         Top             =   9570
         Width           =   8535
         Begin VB.CommandButton cmdPrint 
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   7080
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   240
            Width           =   1275
         End
         Begin VB.CommandButton cmdQty 
            Caption         =   "&Quantity"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   5700
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   240
            Width           =   1275
         End
         Begin VB.CommandButton cmdNoStock 
            Caption         =   "No Stock"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   4305
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton cmdSerchItem 
            Caption         =   "Search Item"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   2910
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton cmdVoidTicket 
            Caption         =   "Void Ticket"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   1515
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   240
            Width           =   1275
         End
         Begin VB.CommandButton cmdVoidItem 
            Caption         =   "Void Item"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   120
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   240
            Width           =   1275
         End
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   240
      Left            =   0
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "frmPosGui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_TicItms As ADODB.Recordset
Dim rs_ItmsLst As ADODB.Recordset
Dim rs_ItmCat As ADODB.Recordset

Dim mTerminalId  As Integer
Dim mdTicketDenom As Double
Dim mActiveEventId As Integer
Dim mActiveEventDescr As String
Dim bPrint As Boolean

Private Enum enmColTicket
    eItmId = 0
    eItmDescr = 1
    eItmQty = 2
    eItmRtlPrc = 3
    eItmDistAmt = 4
    eItmamt = 5
End Enum

Private Enum enmEntry
    eInsert = 1
    eUpdate = 2
End Enum

Const PAYMENT_BARCODE = "00000"

Private Sub SetMshTickets()

    'Set Header Grid------------------------------------------------------------------
    With mshHeader
        .RowHeight(0) = 500
        .Height = 1000
        .Cols = 6
        .ScrollBars = flexScrollBarBoth
        .Enabled = False
        .FontFixed.Name = gGujaratiFontName
        .FontFixed.Size = 12
        
        .ColWidth(enmColTicket.eItmId) = 0
        .ColWidth(enmColTicket.eItmDescr) = 3150
        .ColWidth(enmColTicket.eItmQty) = 600: .ColAlignment(enmColTicket.eItmQty) = flexAlignRightCenter
        .ColWidth(enmColTicket.eItmamt) = 1100: .ColAlignment(enmColTicket.eItmamt) = flexAlignRightCenter
        .ColWidth(enmColTicket.eItmDistAmt) = 1100
        .ColWidth(enmColTicket.eItmRtlPrc) = 1400
        
        .TextMatrix(0, enmColTicket.eItmId) = "ITM_ID"          'Not visible
        .TextMatrix(0, enmColTicket.eItmDescr) = "krJd;"        'Vigat
        .TextMatrix(0, enmColTicket.eItmQty) = "lkd"            'Nang
        .TextMatrix(0, enmColTicket.eItmamt) = "hfb"            'Rakam
        .TextMatrix(0, enmColTicket.eItmDistAmt) = "mntg"       'Sahay
        .TextMatrix(0, enmColTicket.eItmRtlPrc) = "bw¤rfkb;"    'Mul Kinmat
    End With
    '---------------------------------------------------------------------------------
    
    'Set Ticket Grid------------------------------------------------------------------
    With mshTicket
        .FixedCols = 0
        .GridLinesFixed = flexGridInset
        .GridLines = flexGridFlat
        .FocusRect = flexFocusLight
        .HighLight = flexHighlightWithFocus
        .RowHeightMin = 400
        .ScrollBars = flexScrollBarVertical
        
        .Font.Name = gGujaratiFontName
        .Font.Size = 12
        .Cols = 6
        
        .ColWidth(enmColTicket.eItmId) = 0
        .ColWidth(enmColTicket.eItmDescr) = 3150
        .ColWidth(enmColTicket.eItmQty) = 600: .ColAlignment(enmColTicket.eItmQty) = flexAlignRightCenter
        .ColWidth(enmColTicket.eItmamt) = 1100: .ColAlignment(enmColTicket.eItmamt) = flexAlignRightCenter
        
        .ColWidth(enmColTicket.eItmDistAmt) = 1100: .ColAlignment(enmColTicket.eItmDistAmt) = flexAlignRightCenter
        .ColWidth(enmColTicket.eItmRtlPrc) = 1400: .ColAlignment(enmColTicket.eItmRtlPrc) = flexAlignRightCenter
         
        .Move mshHeader.Left, mshHeader.Top + mshHeader.RowHeight(0)
    End With
    '---------------------------------------------------------------------------------
End Sub

Private Sub cmdCategory_Click(Index As Integer)
    
    SetItemListbyCategory Val(cmdCategory(Index).Tag)
    PopulateItemBtns 1

    Dim iCnt As Integer
    For iCnt = 0 To cmdCategory.Count - 1
        cmdCategory(iCnt).FontBold = False
    Next
    cmdCategory(Index).FontBold = True
    
    cmdItemNevigate(1).Visible = cmdItems(cmdItems.Count - 1).Visible
    
End Sub

Private Sub cmdCategoryNeviagte_Click(Index As Integer)
    
    PopulateCategoryBtns rs_ItmCat.AbsolutePage
    
End Sub

Private Sub cmdItemNevigate_Click(Index As Integer)
    
    Select Case Index
        Case 0
            If rs_ItmsLst.AbsolutePage > 2 Then
                PopulateItemBtns rs_ItmsLst.AbsolutePage - 2
            End If
        Case 1
            PopulateItemBtns rs_ItmsLst.AbsolutePage
    End Select
    
End Sub

Private Sub cmdItems_Click(Index As Integer)

    lblBarcodeMsg.Caption = ""
    
    AddItem2Ticket SplitItemDetails(cmdItems.Item(Index).Tag, 0), _
                    cmdItems.Item(Index).Caption, _
                    SplitItemDetails(cmdItems.Item(Index).Tag, 1), _
                    SplitItemDetails(cmdItems.Item(Index).Tag, 2)
    
    
    Dim iCnt As Integer
    For iCnt = 0 To cmdItems.Count - 1
        cmdItems(iCnt).FontBold = False
    Next
    cmdItems(Index).FontBold = True
    
    SetBarcodeLable (True)
    
End Sub

Private Sub SetBarcodeLable(bfound As Boolean)
    If (bfound) Then
        lblBarcodeMsg.Caption = "Added!"
    Else
    
        If (Trim(txtBarcode.Text) = PAYMENT_BARCODE) Then
            lblBarcodeMsg.Caption = ""
        Else
            lblBarcodeMsg.Caption = "Item missing!!"
        End If
    End If
    
    txtBarcode.Text = ""
    txtBarcode.SetFocus

End Sub

Private Sub GetItemByBarcode(s_barcode As String)
    
    If Len(s_barcode) < 4 Then
        SetBarcodeLable (False)
        Exit Sub
    End If
    
    If (Trim(s_barcode) = PAYMENT_BARCODE) Then
        SetBarcodeLable (False)
        cmdPay_Click
        Exit Sub
    End If
    
    Dim rs_BarcodeItm As ADODB.Recordset
    Dim str_ShortName As String
    Set rs_BarcodeItm = New ADODB.Recordset

    SQL = "Exec stpBarcodeItem " & AQ(s_barcode)
    OpenAdoRst rs_BarcodeItm, SQL, adOpenKeyset, , , gCnnMst

    If rs_BarcodeItm.RecordCount <= 0 Then
        SetBarcodeLable (False)

        Exit Sub
    End If

    'Itm_Code ~ Rtl_Prc ~ Disc_Amt
    With rs_BarcodeItm

        If Trim$(.Fields("sizename").Value) = "<unknown>" Then
            'ShortNmae
            str_ShortName = .Fields("shortname").Value & vbCrLf & Format(.Fields("rtl_prc").Value, "###0.00")
        Else
            'Shortname + Size
            str_ShortName = .Fields("shortname").Value & "(" & Trim$(.Fields("sizename").Value) & ")" & vbCrLf & Format(.Fields("rtl_prc").Value, "###0.00")
        End If
    
        AddItem2Ticket .Fields("itm_code").Value, _
                        str_ShortName, _
                        Format(.Fields("rtl_prc").Value, "###0.00"), _
                        Format(.Fields("disc_amt").Value, "###0.00")
    End With
    
    SetBarcodeLable (True)
    
    rs_BarcodeItm.Close
    Set rs_BarcodeItm = Nothing

End Sub

Private Sub cmdPay_Click()

    Dim mbtnOkPressed As Boolean
    mbtnOkPressed = False
    'If Val(cmdPay.Caption) = 0 Then
    '    Exit Sub
    'End If
    
    With frmDenoms
        mdTicketDenom = .Display(Val(cmdPay.Caption), mbtnOkPressed)
    End With
    
    If OperaionMode = enTerminal Or IsNeaturalUserMode Then
        If mbtnOkPressed Then
            SaveSales
        End If
    Else
        MsgBox "Cannot perform Sales on Server", vbExclamation
    End If
    
    If mdTicketDenom = 0 Then Exit Sub
    
    lblChangeAmount.Caption = "Change Amount " & vbCrLf & _
                Format$(mdTicketDenom - Val(cmdPay.Caption), "###0.00")
                
    
    Prepare4NewTiceket
    
End Sub

Private Sub cmdPrint_Click()

    Unload Me
    
End Sub

Private Sub cmdQty_Click()
    Dim mbtnOkPressed As Boolean
    mbtnOkPressed = False
    
    If mshTicket.Rows > 0 Then
        Dim initQty As Integer
        Dim itemDescr As String

        initQty = Val(mshTicket.TextMatrix(mshTicket.RowSel, enmColTicket.eItmQty))
        itemDescr = mshTicket.TextMatrix(mshTicket.RowSel, enmColTicket.eItmDescr) + " - " + _
                    mshTicket.TextMatrix(mshTicket.RowSel, enmColTicket.eItmRtlPrc)
            
        With frmQty
            initQty = .Display(itemDescr, initQty, mbtnOkPressed)
        End With
        
        If mbtnOkPressed Then
            ''Update Qty
            UpdateItem mshTicket.RowSel, initQty
            
        End If
    End If
    
End Sub

Private Sub cmdVoidItem_Click()
    
    RemoveItem
    
    SetBarcodeLable (False)

End Sub

Private Sub cmdVoidTicket_Click()

    Prepare4NewTiceket
    
    SetBarcodeLable (False)
    
End Sub

Private Sub Form_Activate()
    
    cmdCategoryNeviagte(0).Visible = cmdCategory(cmdCategory.Count - 1).Visible

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = vbKeyReturn Then
        cmdPay_Click
    ElseIf KeyCode = vbKeyReturn And Trim$(txtBarcode.Text) <> "" Then
        GetItemByBarcode (txtBarcode.Text)
    'ElseIf KeyCode = vbKeyQ Then
    '    KeyCode = 0
    '    cmdQty_Click
    End If

End Sub

Private Sub Form_Load()

    CenterFrmNonChild Me

    SetMshTickets
    
    SetControls
    
    RebindTicketGrid
    
    Call GetPrinter(Command1)
    bPrint = gPrintEnable

    MP vbDefault
    
End Sub

Private Sub AddItem2Ticket(s_iItmId As Long, s_sItmDescr As String, s_dUnitPrice As Double, s_dDiscAmt As Double)
        
    Dim mItemExist As Boolean
    
    lblChangeAmount.Caption = ""
    
    With rs_TicItms
        mItemExist = IsItemExists(s_iItmId)

        If Not mItemExist Then
            .AddNew
        End If
        
        .Fields("itm_id").Value = Val(s_iItmId)
        .Fields("shortname") = Trim$(Mid$(s_sItmDescr, 1, InStr(1, s_sItmDescr, vbCrLf) - 1))
        .Fields("qty") = Val(.Fields("qty")) + 1
        .Fields("rtl_amt") = Format$(Val(s_dUnitPrice - Val(s_dDiscAmt)) * Val(.Fields("qty")), "###.00")
        .Fields("disc_amt") = Format$(Val(s_dDiscAmt) * Val(.Fields("qty")), "###.00")
        .Fields("rtl_prc") = Format$(Val(s_dUnitPrice), "###.00")
        .Update
    End With
    
    RebindTicketGrid
    
End Sub

Private Function IsItemExists(s_iItmId As Long) As Boolean
    
    Dim i As Integer
    
    IsItemExists = False

    With rs_TicItms
        If .State = adStateClosed Then .Open
        If (.RecordCount <= 0) Then
            Exit Function
        End If
        
        .MoveFirst
        
        While Not .EOF
            If Val(.Fields("itm_id").Value) = Val(s_iItmId) Then
                IsItemExists = True
                Exit Function
            End If
            .MoveNext
        Wend
    End With
    
End Function

Private Sub TicketTotal()

    Dim i As Integer
    Dim iQty As Integer
    Dim dTicketAmt As Double
    Dim dDiscAmt As Double
    
    iQty = 0
    dTicketAmt = 0
    dDiscAmt = 0
    
    With mshTicket
        For i = 0 To .Rows - 1
            iQty = iQty + Val(.TextMatrix(i, enmColTicket.eItmQty))
            dTicketAmt = dTicketAmt + Val(.TextMatrix(i, enmColTicket.eItmamt))
            dDiscAmt = dDiscAmt + Val(.TextMatrix(i, enmColTicket.eItmDistAmt))
        Next
    End With
            
    txtTotalQty.Text = iQty
    txtTotalAmount.Text = Format$(dTicketAmt, "###.00")
    txtTotalDiscount.Text = Format$(dDiscAmt, "###.00")
    
    cmdPay.Caption = txtTotalAmount.Text
    
End Sub

Private Sub SetControls()
    
    Me.WindowState = vbMaximized
    Me.frmMainContainer.BorderStyle = vbNone
    
    Set rs_TicItms = New ADODB.Recordset
    Set rs_ItmsLst = New ADODB.Recordset
    Set rs_ItmCat = New ADODB.Recordset

    ' Set Default Value for Barcode
    lblBarcodeMsg.Caption = ""
    txtBarcode.Text = ""

    'Set Terminal Id from INT file-----------------------------------
    mTerminalId = gTerminalId

    'Load Categories--------------------------------------------------
    SQL = "Exec stpFetchItemCategory " & GetLinkedKbdId
    OpenAdoRst rs_ItmCat, SQL, adOpenKeyset, , , gCnnMst
    PopulateCategoryBtns 1

    'Load Items-------------------------------------------------------
    SetItemListbyCategory Val(cmdCategory(0).Tag)
    PopulateItemBtns 1

    'cmdCategory_Click (0)

    With rs_TicItms
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        
        .Fields.Append "itm_id", adVarChar, 10
        .Fields.Append "shortname", adVarChar, 35
        .Fields.Append "qty", adVarChar, 5
        .Fields.Append "rtl_prc", adVarChar, 12
        .Fields.Append "disc_amt", adVarChar, 12
        .Fields.Append "rtl_amt", adVarChar, 12
    End With

    With mshHeader
        
        txtTotalQty.Width = .ColWidth(enmColTicket.eItmQty) + 10
        txtTotalAmount.Width = .ColWidth(enmColTicket.eItmamt) + 200
        txtTotalDiscount.Width = .ColWidth(enmColTicket.eItmDistAmt) + 200
        
        txtTotalQty.Left = .Left + .ColPos(enmColTicket.eItmQty)
        txtTotalAmount.Left = .Left + .ColPos(enmColTicket.eItmamt)
        txtTotalDiscount.Left = .Left + .ColPos(enmColTicket.eItmDistAmt) + 180
        
    End With
    
    lblChangeAmount.Caption = ""
    cmdItemNevigate(1).Caption = " >>"
    
    'Set Gujarati Font
    txtTotalQty.Font.Name = gGujaratiFontName
    txtTotalAmount.Font.Name = gGujaratiFontName
    txtTotalDiscount.Font.Name = gGujaratiFontName
    cmdPay.Font.Name = gGujaratiFontName
    
    txtTotalQty.Font.Size = 14
    txtTotalAmount.Font.Size = 14
    txtTotalDiscount.Font.Size = 14
    cmdPay.Font.Size = 16
    
    txtTotalQty.Font.Bold = True
    txtTotalAmount.Font.Bold = True
    txtTotalDiscount.Font.Bold = True
    cmdPay.Font.Bold = True
    
    bPrint = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set rs_TicItms = Nothing
    Set rs_ItmsLst = Nothing
    Set rs_ItmCat = Nothing
    
    'Cancel Printer Handler
'    BiCancelStatusBack (mpHandle)
'    If Not BiCloseMonPrinter(mpHandle) = SUCCESS Then
'        MsgBox ("Failed to close printer status monitor.")
'    End If
    
    Set rstPrint = Nothing

End Sub

Private Sub RebindTicketGrid()
    
On Error GoTo errorhandler

    With mshTicket
        .Redraw = False
        
        Set .Recordset = rs_TicItms
        SetMshTickets
        
        'SetGridColGujFont mshTicket, enmColTicket.eItmDescr
        
        .Redraw = True
        
        cmdVoidItem.Enabled = (.Rows > 0)
        cmdVoidTicket.Enabled = (.Rows > 0)
        cmdPay.Enabled = (.Rows > 0)
        cmdQty.Enabled = (.Rows > 0)
        
        ''Make last row selected by default
        If .Rows > 0 Then
            mshTicket.RowSel = (mshTicket.Rows - 1)
        End If
        
    End With

    TicketTotal
    
    'SetGridRowColor mshTicket, mGridTicketBookMark - 1, vbCyan
    
errorhandler:
    
    Select Case Err.Number
        Case -2147217885        'Delete mode
            Resume Next
    End Select

End Sub

Private Sub RemoveItem()

    Dim mItemId As Long
    Dim dDiscAmt As Double
    
    If mshTicket.Rows > 0 Then
        mItemId = Val(mshTicket.TextMatrix(mshTicket.Row, enmColTicket.eItmId))
        
        With rs_TicItms
            If IsItemExists(mItemId) Then
                If Val(.Fields("qty").Value) > 0 Then
                    dDiscAmt = Val(.Fields("disc_amt")) / Val(.Fields("qty").Value)
                End If
                
                .Fields("qty").Value = Val(.Fields("qty").Value) - 1
                .Fields("rtl_amt") = Format$((Val(.Fields("rtl_prc")) - dDiscAmt) * Val(.Fields("qty")), "###.00")
                .Fields("disc_amt") = Format$(dDiscAmt * Val(.Fields("qty")), "###.00")
                
                .Update
                If .Fields("qty").Value = 0 Then
                    .Delete adAffectCurrent
                    .Update
                End If
            End If
        End With
    
    End If
    
    RebindTicketGrid
End Sub


Private Sub UpdateItem(selectedRow As Integer, Qty As Integer)

    Dim mItemId As Long
    Dim dDiscAmt As Double
    
    If mshTicket.Rows > 0 Then
        mItemId = Val(mshTicket.TextMatrix(selectedRow, enmColTicket.eItmId))
        
        With rs_TicItms
            If IsItemExists(mItemId) Then
                If Val(.Fields("qty").Value) > 0 Then
                    dDiscAmt = Val(.Fields("disc_amt")) / Val(.Fields("qty").Value)
                End If
                
                .Fields("qty").Value = Qty
                .Fields("rtl_amt") = Format$((Val(.Fields("rtl_prc")) - dDiscAmt) * Val(.Fields("qty")), "###.00")
                .Fields("disc_amt") = Format$(dDiscAmt * Val(.Fields("qty")), "###.00")
                
                .Update
                If .Fields("qty").Value = 0 Then
                    .Delete adAffectCurrent
                    .Update
                End If
            End If
        End With
    
    End If
    
    RebindTicketGrid
End Sub


Private Sub Prepare4NewTiceket()
    
    With rs_TicItms
        .MoveFirst
        While Not .EOF
            .Delete adAffectCurrent
            .Update
            .MoveNext
        Wend
    End With
    
    RebindTicketGrid
    
End Sub

Private Sub PopulateItemBtns(s_iPage As Integer)

    Dim iCnt As Integer
    Const PAGE_SIZE As Integer = 19
    
    For iCnt = 0 To PAGE_SIZE - 1
        cmdItems(iCnt).Visible = False
    Next

    If s_iPage <= 0 Then Exit Sub

    If rs_ItmsLst.RecordCount <= 0 Then
        Exit Sub
    End If


    With rs_ItmsLst
        cmdItemNevigate(1).Caption = Val(.AbsolutePage) & " >>"
        .PageSize = PAGE_SIZE
        .AbsolutePage = s_iPage
        For iCnt = 0 To .PageSize - 1
        
            cmdItems(iCnt).FontName = gGujaratiFontName
            cmdItems(iCnt).FontSize = 12
            cmdItems(iCnt).FontBold = False

            If .EOF = True Then
                cmdItems(iCnt).Visible = False
            Else
                If Trim$(.Fields("sizename").Value) = "<unknown>" Then
                    'ShortNmae
                    cmdItems(iCnt).Caption = .Fields("shortname").Value & vbCrLf & Format(.Fields("rtl_prc").Value, "###0.00")
                Else
                    'Shortname + Size
                    cmdItems(iCnt).Caption = .Fields("shortname").Value & "(" & Trim$(.Fields("sizename").Value) & ")" & vbCrLf & Format(.Fields("rtl_prc").Value, "###0.00")
                End If
                'Itm_Code ~ Rtl_Prc ~ Disc_Amt
                cmdItems(iCnt).Tag = .Fields("itm_code").Value & _
                                    "~" & Format(.Fields("rtl_prc").Value, "###0.00") & _
                                    "~" & Format(.Fields("disc_amt").Value, "###0.00")
                                    
                cmdItems(iCnt).Visible = True
                .MoveNext
            End If
        Next
        If .AbsolutePage = adPosEOF Then
            .MoveFirst
        End If
    End With
End Sub

Private Sub PopulateCategoryBtns(s_iPage As Integer)

    Dim iCnt As Integer
    Const CAT_PAGE_SIZE As Integer = 9
    
    For iCnt = 0 To CAT_PAGE_SIZE - 1
        cmdCategory(iCnt).Visible = False
    Next
    
    If rs_ItmCat.RecordCount <= 0 Or s_iPage <= 0 Then
        cmdCategoryNeviagte(0).Visible = False
        Exit Sub
    End If


    With rs_ItmCat
        .PageSize = CAT_PAGE_SIZE
        .AbsolutePage = s_iPage
        For iCnt = 0 To .PageSize - 1
            cmdCategory(iCnt).FontName = gGujaratiFontName
            cmdCategory(iCnt).FontSize = 12
            cmdCategory(iCnt).FontBold = False
            
            If .EOF = True Then
                cmdCategory(iCnt).Visible = False
            Else
                cmdCategory(iCnt).Caption = .Fields("shortname").Value
                cmdCategory(iCnt).Tag = .Fields("code").Value
                cmdCategory(iCnt).Visible = True
                .MoveNext
            End If
        Next
        If .AbsolutePage = adPosEOF Then
            .MoveFirst
        End If
    End With
    
End Sub

Private Function SplitItemDetails(s_sItemDetails As String, s_iItemIndex As Integer)
    
    Dim temp
    temp = Split(s_sItemDetails, "~")
    
    If UBound(temp) >= s_iItemIndex Then
        SplitItemDetails = temp(s_iItemIndex)
    Else
        SplitItemDetails = ""
    End If
    
End Function

Private Sub Form_Resize()
    With Shape1
        .BorderWidth = 5
        .Move 0, 0, Me.Width, Me.Height
    End With
    
    With frmHeader
        .Move 40, 150, Me.Width - 90, 405
    End With
    
    Dim sCaption As String
    
    'Set Active Event-------------------------------------------------------
    mActiveEventId = GetActiveEvent

    If mActiveEventId <= 0 Then
        sCaption = Space(10)
    Else
        sCaption = "Event : " & mActiveEventDescr & Space(10)
    End If
    
    sCaption = sCaption & "Terminal Id : " & mTerminalId & Space(10)
    sCaption = sCaption & "User : " & gUserName & "[" & gUser & "]" & Space(10)
    sCaption = sCaption & "Date : " & Format(Date, "dd/MM/yyyy") & Space(5)
    
    'Header Caption
    With lblHeader(0)
        .Caption = sCaption
                                    
        .Move 0, (frmHeader.Height - lblHeader(0).Height) / 2, frmHeader.Width
        .Alignment = vbRightJustify
        .ZOrder vbSendToBack
    End With
    
    With lblHeader(1)
        .Caption = lblHeader(0).Caption
                                    
        .Move lblHeader(0).Left + 20, lblHeader(0).Top + 15, lblHeader(0).Width, lblHeader(0).Height
        .Alignment = vbRightJustify
        .ZOrder vbBringToFront
    End With
    
End Sub



Private Sub SetItemListbyCategory(s_iCategoryId As Integer)
        
    SQL = "Exec stpItemKeybrdList " & s_iCategoryId & "," & GetLinkedKbdId
    OpenAdoRst rs_ItmsLst, SQL, adOpenKeyset, , , gCnnMst

End Sub

Private Function SQLSaveTrn(s_EntryId As String) As String
On Error GoTo errhndl
MP vbHourglass
        
    Dim sSQL As String
    
    sSQL = "IF NOT EXISTS (" & vbCrLf
    sSQL = sSQL & "     SELECT '1' " & vbCrLf
    sSQL = sSQL & "     FROM TerSaltrn " & vbCrLf
    sSQL = sSQL & "     WHERE tran_id = " & AQ(s_EntryId) & vbCrLf
    sSQL = sSQL & "     )" & vbCrLf
    
    sSQL = sSQL & " BEGIN " & vbCrLf
    sSQL = sSQL & "     Insert into TerSaltrn(" & vbCrLf
    sSQL = sSQL & "          tran_id" & vbCrLf
    sSQL = sSQL & "         ,ter_id" & vbCrLf
    sSQL = sSQL & "         ,export_fg" & vbCrLf
    
    sSQL = sSQL & "         ,paid_amt" & vbCrLf
    sSQL = sSQL & "         ,change_amt" & vbCrLf
    
    sSQL = sSQL & "         ,dtadat" & vbCrLf
    sSQL = sSQL & "         ,dtatim" & vbCrLf
    sSQL = sSQL & "         ,dtausr" & vbCrLf
    sSQL = sSQL & "         ,Event_id" & vbCrLf
    sSQL = sSQL & "         ,Trng_fg"
    
    sSQL = sSQL & " ) Values (" & vbCrLf
    
    sSQL = sSQL & "          " & AQ(s_EntryId) & vbCrLf
    sSQL = sSQL & "         ," & mTerminalId & vbCrLf
    sSQL = sSQL & "         ," & "0" & vbCrLf       'False
    
    sSQL = sSQL & "         ," & Val(mdTicketDenom) & vbCrLf
    sSQL = sSQL & "         ," & mdTicketDenom - Val(cmdPay.Caption) & vbCrLf
    
    sSQL = sSQL & "         ," & ConvDatSql(Date) & vbCrLf
    sSQL = sSQL & "         ," & AQ(DtaTime) & vbCrLf
    sSQL = sSQL & "         ," & AQ(gUser) & vbCrLf
    sSQL = sSQL & "         ," & mActiveEventId
    sSQL = sSQL & "         ," & IsTrainingMode
    
    sSQL = sSQL & ")"
    sSQL = sSQL & " END " & vbCrLf
    
    sSQL = sSQL & " ELSE " & vbCrLf
    
    sSQL = sSQL & " BEGIN " & vbCrLf
    sSQL = sSQL & "     Update TerSaltrn" & vbCrLf
    sSQL = sSQL & "     Set  tran_id    = " & AQ(s_EntryId) & vbCrLf
    sSQL = sSQL & "         ,ter_id     = " & mTerminalId & vbCrLf
    sSQL = sSQL & "         ,export_fg  = " & "0" & vbCrLf   'False
    
    sSQL = sSQL & "         ,dtadat     = " & ConvDatSql(Date) & vbCrLf
    sSQL = sSQL & "         ,dtatim     = " & AQ(DtaTime) & vbCrLf
    sSQL = sSQL & "         ,dtausr     = " & AQ(gUser) & vbCrLf
    sSQL = sSQL & "         ,Trng_fg    = " & IsTrainingMode & vbCrLf
    sSQL = sSQL & "     WHERE tran_id = " & AQ(s_EntryId) & vbCrLf
    sSQL = sSQL & " END "
    
    
    SQLSaveTrn = sSQL
    
MP vbDefault
Exit Function
errhndl:
    ErrMsg

End Function

Private Function SQLSaveDet(s_EntryId As String) As String
On Error GoTo errhndl
MP vbHourglass
    
    Dim sSQL As String
    Dim iCnt As Integer
    
    With mshTicket
        For iCnt = 0 To .Rows - 1
        
            sSQL = sSQL & " IF NOT EXISTS (" & vbCrLf
            sSQL = sSQL & "     SELECT '1' " & vbCrLf
            sSQL = sSQL & "     FROM TerSaldet " & vbCrLf
            sSQL = sSQL & "     WHERE tran_id = " & AQ(s_EntryId) & vbCrLf
            sSQL = sSQL & "     AND itm_code = " & Val(.TextMatrix(iCnt, enmColTicket.eItmId)) & vbCrLf
            sSQL = sSQL & "     )" & vbCrLf
            
            sSQL = sSQL & " BEGIN " & vbCrLf
            sSQL = sSQL & "     Insert into TerSaldet(" & vbCrLf
            sSQL = sSQL & "          tran_id" & vbCrLf
            sSQL = sSQL & "         ,tran_seq" & vbCrLf
            sSQL = sSQL & "         ,itm_code" & vbCrLf
            
            sSQL = sSQL & "         ,rtl_prc" & vbCrLf
            sSQL = sSQL & "         ,disc_amt" & vbCrLf
            sSQL = sSQL & "         ,qty" & vbCrLf
            sSQL = sSQL & "         ,amt" & vbCrLf
            'sSQL = sSQL & "         ,Trng_fg" & vbCrLf
            
            sSQL = sSQL & "     ) Values (" & vbCrLf
            
            sSQL = sSQL & "          " & AQ(s_EntryId) & vbCrLf
            sSQL = sSQL & "         ," & iCnt + 1 & vbCrLf
            sSQL = sSQL & "         ," & Val(.TextMatrix(iCnt, enmColTicket.eItmId)) & vbCrLf
            
            sSQL = sSQL & "         ," & Val(.TextMatrix(iCnt, enmColTicket.eItmRtlPrc)) & vbCrLf
            sSQL = sSQL & "         ," & Val(.TextMatrix(iCnt, enmColTicket.eItmDistAmt)) & vbCrLf
            sSQL = sSQL & "         ," & Val(.TextMatrix(iCnt, enmColTicket.eItmQty)) & vbCrLf
            sSQL = sSQL & "         ," & Val(.TextMatrix(iCnt, enmColTicket.eItmamt)) & vbCrLf
            'sSQL = sSQL & "         ," & IsTrainingMode
            
            sSQL = sSQL & ")" & vbCrLf
            sSQL = sSQL & " END " & vbCrLf
            
            sSQL = sSQL & " ELSE " & vbCrLf
            
            sSQL = sSQL & " BEGIN " & vbCrLf
            sSQL = sSQL & "     Update TerSaldet " & vbCrLf
            sSQL = sSQL & "     Set  tran_id = " & AQ(s_EntryId) & vbCrLf
            sSQL = sSQL & "         ,tran_seq = " & iCnt + 1 & vbCrLf
            sSQL = sSQL & "         ,itm_code = " & Val(.TextMatrix(iCnt, enmColTicket.eItmId)) & vbCrLf
            
            sSQL = sSQL & "         ,rtl_prc = " & Val(.TextMatrix(iCnt, enmColTicket.eItmRtlPrc)) & vbCrLf
            sSQL = sSQL & "         ,disc_amt = " & Val(.TextMatrix(iCnt, enmColTicket.eItmDistAmt)) & vbCrLf
            sSQL = sSQL & "         ,qty = " & Val(.TextMatrix(iCnt, enmColTicket.eItmQty)) & vbCrLf
            sSQL = sSQL & "         ,amt = " & Val(.TextMatrix(iCnt, enmColTicket.eItmamt)) & vbCrLf
            'sSQL = sSQL & "         ,Trng_fg = " & IsTrainingMode & vbCrLf
            
            sSQL = sSQL & "     where tran_id = " & AQ(s_EntryId) & vbCrLf
            sSQL = sSQL & "     and itm_code = " & Val(.TextMatrix(iCnt, enmColTicket.eItmId)) & vbCrLf
            
            sSQL = sSQL & " END " & vbCrLf
        Next
    End With
    SQLSaveDet = sSQL
    
MP vbDefault
Exit Function
errhndl:
    ErrMsg

End Function

Private Sub SaveSales()
On Error GoTo errhndl
MP vbHourglass

    Dim sEntryId As String
    Dim str_dt As String
    Dim SQL1 As String
    Dim SQL2 As String
    
    str_dt = GetYYMMDD(Date)
    
    If UserLevel(gUser) = eTraining Then
        str_dt = "990101"
    End If
    
    gCnnMst.BeginTrans
    
        sEntryId = GetNewSalesTranId(str_dt)
    
        SQL1 = "": SQL2 = ""
        SQL1 = SQLSaveTrn(sEntryId) & vbCrLf
        SQL2 = SQLSaveDet(sEntryId) & vbCrLf
        
        gCnnMst.Execute SQL1 & SQL2
        
        InsertUpdateBillSequence eUpdate, str_dt
    
    gCnnMst.CommitTrans
            
    If bPrint Then
        Call PrintPreview(sEntryId)
    End If

MP vbDefault
Exit Sub
errhndl:
    gCnnMst.RollbackTrans
    ErrMsg
End Sub

Private Function GetNewSalesTranId(s_strdt As String) As String

On Error GoTo errhndl
MP vbHourglass

    Dim strtmp As String
    
    Dim rsttmp As ADODB.Recordset
    Set rsttmp = New ADODB.Recordset
    
    SQL = " Select IsNull(seq,0) + 1 "
    SQL = SQL & " From BillSequence "
    SQL = SQL & " Where dtadat = " & AQ(s_strdt)
    SQL = SQL & " and ter_id = " & mTerminalId
    
    rsttmp.Open SQL, gCnnMst
    
    If rsttmp.RecordCount <= 0 Then
        InsertUpdateBillSequence eInsert, s_strdt
        strtmp = 1
    Else
        strtmp = rsttmp.Fields(0).Value
    End If
    
    GetNewSalesTranId = Trim$(s_strdt) & _
                        mTerminalId & _
                        Format(strtmp, "0000")
                        
    
    Set rsttmp = Nothing
    
MP vbDefault
Exit Function
errhndl:
    ErrMsg
End Function

Private Function GetYYMMDD(s_Date As Date) As String
    
    GetYYMMDD = Format(Right(Year(s_Date), 2), "00") & _
                Format(Month(s_Date), "00") & _
                Format(Day(s_Date), "00")
    
End Function


Private Sub InsertUpdateBillSequence(s_InsertUpdate As enmEntry, s_strdt As String)
    
    If s_InsertUpdate = eInsert Then
        SQL = "Insert Into BillSequence ( " & vbCrLf
        SQL = SQL & " dtadat"
        SQL = SQL & ",ter_id"
        SQL = SQL & ",seq"
        SQL = SQL & ",Trng_fg"
        
        SQL = SQL & " ) Values ( "
        SQL = SQL & AQ(s_strdt)
        SQL = SQL & "," & mTerminalId
        SQL = SQL & ",1"
        SQL = SQL & "," & IsTrainingMode
        SQL = SQL & ")"
    
    ElseIf s_InsertUpdate = eUpdate Then
        SQL = "Update BillSequence  " & vbCrLf
        SQL = SQL & " Set   seq = Isnull(seq,0)+1" & vbCrLf
        SQL = SQL & " Where dtadat = " & AQ(s_strdt) & vbCrLf
        SQL = SQL & " and ter_id = " & mTerminalId
    End If
    
    gCnnMst.Execute SQL

End Sub

Private Sub mshTicket_Scroll()
    mshHeader.LeftCol = mshTicket.LeftCol
End Sub


Private Function GetLinkedKbdId()

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    
    SQL = "Select Isnull(Keybrd_id,0) "
    SQL = SQL & " From TerminalConfig"
    SQL = SQL & " Where Code = " & mTerminalId
    
    OpenAdoRst rst, SQL
    
    If rst.RecordCount > 0 Then
        GetLinkedKbdId = rst.Fields(0).Value
    Else
        GetLinkedKbdId = 0
    End If
    
    rst.Close
    Set rst = Nothing
    
End Function


Private Function GetActiveEvent() As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    SQL = "Select [code],[name] from EventMast where actv_fg = 1"
    OpenAdoRst rs, SQL
    
    If rs.RecordCount > 0 Then
        GetActiveEvent = rs.Fields("code").Value
        mActiveEventDescr = rs.Fields("name").Value
    Else
        GetActiveEvent = 0
        mActiveEventDescr = ""
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

' The callback function that will monitor printer/printing status.
Private Sub Command1_Click()
        On Error Resume Next
        
    If (lpdwStatus And ASB_PRINT_SUCCESS) = ASB_PRINT_SUCCESS Or _
       (lpdwStatus And ASB_NO_RESPONSE) = ASB_NO_RESPONSE Or _
       (lpdwStatus And ASB_COVER_OPEN) = ASB_COVER_OPEN Or _
       (lpdwStatus And ASB_AUTOCUTTER_ERR) = ASB_AUTOCUTTER_ERR Or _
       ((lpdwStatus And ASB_PAPER_END_FIRST) = ASB_PAPER_END_FIRST) Or ((lpdwStatus And ASB_PAPER_END_SECOND) = ASB_PAPER_END_SECOND) Then
        isFinish = True
        status = lpdwStatus
    End If
End Sub

