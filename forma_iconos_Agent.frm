VERSION 5.00
Begin VB.Form forma_iconos 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   15375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2565
   ControlBox      =   0   'False
   Icon            =   "forma_iconos_Agent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15375
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1560
      TabIndex        =   61
      Top             =   12120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   58
      Top             =   8040
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   12
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   12
         Left            =   360
         TabIndex        =   60
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   11
      Left            =   1400
      TabIndex        =   55
      Top             =   6960
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   11
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   11
         Left            =   360
         TabIndex        =   57
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   52
      Top             =   6960
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   10
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   10
         Left            =   360
         TabIndex        =   54
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   9
      Left            =   1400
      TabIndex        =   49
      Top             =   5760
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   9
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   9
         Left            =   360
         TabIndex        =   51
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   46
      Top             =   5760
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   8
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   8
         Left            =   360
         TabIndex        =   48
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   7
      Left            =   1400
      TabIndex        =   43
      Top             =   4560
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   7
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   7
         Left            =   360
         TabIndex        =   45
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   40
      Top             =   4560
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   6
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   6
         Left            =   360
         TabIndex        =   42
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   5
      Left            =   1400
      TabIndex        =   37
      Top             =   3360
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   5
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   5
         Left            =   360
         TabIndex        =   39
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   34
      Top             =   3360
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   4
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   4
         Left            =   360
         TabIndex        =   36
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   1400
      TabIndex        =   31
      Top             =   2160
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   3
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   3
         Left            =   360
         TabIndex        =   33
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   2160
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   2
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   2
         Left            =   360
         TabIndex        =   30
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   9960
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808080&
      Caption         =   "Program"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin Project1.lvButtons_H btnbrowser 
      Height          =   585
      Index           =   1
      Left            =   720
      TabIndex        =   16
      Top             =   13800
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   1032
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "forma_iconos_Agent.frx":3336E
      ImgSize         =   48
      cBack           =   16777215
   End
   Begin Project1.lvButtons_H btnoffice 
      Height          =   300
      Left            =   1700
      TabIndex        =   23
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      Caption         =   "Office"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnPasswords 
      Height          =   420
      Left            =   240
      TabIndex        =   22
      Top             =   12920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Caption         =   "My Passwords"
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   4210752
      cGradient       =   4210752
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "forma_iconos_Agent.frx":3689E
      cBack           =   12632319
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   360
      Left            =   240
      ScaleHeight     =   300
      ScaleWidth      =   2115
      TabIndex        =   20
      Top             =   13320
      Width           =   2175
      Begin VB.Label lblIP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VNC:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   40
         Width           =   1695
      End
   End
   Begin Project1.lvButtons_H btncolor 
      Height          =   180
      Index           =   3
      Left            =   1200
      TabIndex        =   6
      Top             =   14520
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483646
   End
   Begin Project1.lvButtons_H btncolor 
      Height          =   180
      Index           =   4
      Left            =   1440
      TabIndex        =   7
      Top             =   14520
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin Project1.lvButtons_H btncolor 
      Height          =   180
      Index           =   6
      Left            =   1920
      TabIndex        =   9
      Top             =   14520
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   4210752
   End
   Begin Project1.lvButtons_H btncolor 
      Height          =   180
      Index           =   5
      Left            =   1680
      TabIndex        =   8
      Top             =   14520
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12648384
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   1400
      TabIndex        =   17
      Top             =   960
      Width           =   1140
      Begin Project1.lvButtons_H btnchrome 
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Chrome"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   -1  'True
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnedge 
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
         Caption         =   "Edge"
         CapAlign        =   1
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   4210752
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
   End
   Begin Project1.lvButtons_H btnbrowser 
      Height          =   570
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   13800
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1005
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "forma_iconos_Agent.frx":36CF0
      ImgSize         =   48
      cBack           =   16777215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1440
      TabIndex        =   12
      Top             =   13680
      Width           =   1215
      Begin Project1.lvButtons_H btnlocaliza 
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         Caption         =   "1"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         cGradient       =   0
         Mode            =   2
         Value           =   -1  'True
         cBack           =   16711680
      End
      Begin Project1.lvButtons_H btnlocaliza 
         Height          =   435
         Index           =   1
         Left            =   480
         TabIndex        =   14
         Top             =   120
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         Caption         =   "2"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   16711680
      End
      Begin VB.Image Image5 
         Height          =   735
         Index           =   1
         Left            =   480
         Picture         =   "forma_iconos_Agent.frx":3A088
         Stretch         =   -1  'True
         Top             =   0
         Width           =   435
      End
      Begin VB.Image Image5 
         Height          =   735
         Index           =   0
         Left            =   0
         Picture         =   "forma_iconos_Agent.frx":3ACEB
         Stretch         =   -1  'True
         Top             =   0
         Width           =   435
      End
   End
   Begin Project1.lvButtons_H btncolor 
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   14520
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   -1  'True
      cBack           =   8421504
   End
   Begin Project1.lvButtons_H btnsalir 
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   10560
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1085
      Caption         =   "Exit"
      CapAlign        =   2
      BackStyle       =   1
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btncolor 
      Height          =   180
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   14520
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   16761087
   End
   Begin Project1.lvButtons_H btncolor 
      Height          =   180
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   14520
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   8438015
   End
   Begin Project1.lvButtons_H btnprogram 
      Height          =   405
      Left            =   2040
      TabIndex        =   2
      Top             =   6360
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   714
      Caption         =   "  P"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   4210752
      cFHover         =   4210752
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgSize         =   40
      cBack           =   -2147483632
   End
   Begin VB.Image btnaspiradora 
      Height          =   735
      Left            =   240
      Picture         =   "forma_iconos_Agent.frx":3B94E
      Stretch         =   -1  'True
      ToolTipText     =   "Clean the Desktop"
      Top             =   11400
      Width           =   1095
   End
   Begin VB.Image btn2 
      Height          =   525
      Index           =   23
      Left            =   1080
      Picture         =   "forma_iconos_Agent.frx":51CEA
      Stretch         =   -1  'True
      ToolTipText     =   "Open Dialpad.com"
      Top             =   12480
      Width           =   495
   End
   Begin VB.Image btn2 
      Height          =   765
      Index           =   22
      Left            =   120
      Picture         =   "forma_iconos_Agent.frx":55EB2
      Stretch         =   -1  'True
      ToolTipText     =   "Making calls on Dialpad"
      Top             =   12240
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   855
      Index           =   21
      Left            =   1485
      Picture         =   "forma_iconos_Agent.frx":58782
      Stretch         =   -1  'True
      ToolTipText     =   "Vehicle Moving Permit"
      Top             =   11235
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   765
      Index           =   20
      Left            =   1620
      Picture         =   "forma_iconos_Agent.frx":5AD5B
      Stretch         =   -1  'True
      ToolTipText     =   "Copy Lastpass' Password  to Clipboard"
      Top             =   12360
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   -120
      Picture         =   "forma_iconos_Agent.frx":5CF3A
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image btn2 
      Height          =   825
      Index           =   8
      Left            =   165
      Picture         =   "forma_iconos_Agent.frx":60191
      Stretch         =   -1  'True
      ToolTipText     =   "Teams"
      Top             =   4965
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   885
      Index           =   19
      Left            =   1485
      Picture         =   "forma_iconos_Agent.frx":62674
      Stretch         =   -1  'True
      ToolTipText     =   "Money Report"
      Top             =   10320
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   975
      Index           =   18
      Left            =   165
      Picture         =   "forma_iconos_Agent.frx":65F87
      Stretch         =   -1  'True
      ToolTipText     =   "UW program"
      Top             =   10320
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   825
      Index           =   17
      Left            =   1440
      Picture         =   "forma_iconos_Agent.frx":69AFE
      Stretch         =   -1  'True
      ToolTipText     =   "Calculator"
      Top             =   9400
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   975
      Index           =   16
      Left            =   165
      Picture         =   "forma_iconos_Agent.frx":6DB45
      Stretch         =   -1  'True
      ToolTipText     =   "Transfer folder"
      Top             =   9360
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   990
      Index           =   15
      Left            =   1440
      Picture         =   "forma_iconos_Agent.frx":71023
      Stretch         =   -1  'True
      ToolTipText     =   "Appointment Scheduler"
      Top             =   8340
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   975
      Index           =   14
      Left            =   120
      Picture         =   "forma_iconos_Agent.frx":75004
      Stretch         =   -1  'True
      ToolTipText     =   "Discrepancy Program"
      Top             =   8400
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   975
      Index           =   13
      Left            =   1395
      Picture         =   "forma_iconos_Agent.frx":78A15
      Stretch         =   -1  'True
      ToolTipText     =   "Vacation program"
      Top             =   7320
      Width           =   1050
   End
   Begin VB.Image btn2 
      Height          =   855
      Index           =   12
      Left            =   210
      Picture         =   "forma_iconos_Agent.frx":7C517
      Stretch         =   -1  'True
      ToolTipText     =   "Authorize.net"
      Top             =   7320
      Width           =   930
   End
   Begin VB.Image btn2 
      Height          =   855
      Index           =   11
      Left            =   1440
      Picture         =   "forma_iconos_Agent.frx":7E681
      Stretch         =   -1  'True
      ToolTipText     =   "Word"
      Top             =   6165
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   825
      Index           =   10
      Left            =   165
      Picture         =   "forma_iconos_Agent.frx":8156A
      Stretch         =   -1  'True
      ToolTipText     =   "Excel"
      Top             =   6120
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   855
      Index           =   9
      Left            =   1485
      Picture         =   "forma_iconos_Agent.frx":84C95
      Stretch         =   -1  'True
      ToolTipText     =   "Onedrive"
      Top             =   4920
      Width           =   885
   End
   Begin VB.Image btn2 
      Height          =   975
      Index           =   0
      Left            =   165
      Picture         =   "forma_iconos_Agent.frx":85EC8
      Stretch         =   -1  'True
      ToolTipText     =   "Clock in"
      Top             =   0
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   855
      Index           =   7
      Left            =   1455
      Picture         =   "forma_iconos_Agent.frx":8A018
      Stretch         =   -1  'True
      ToolTipText     =   "Email"
      Top             =   3720
      Width           =   960
   End
   Begin VB.Image btn2 
      Height          =   855
      Index           =   6
      Left            =   195
      Picture         =   "forma_iconos_Agent.frx":8C5EB
      Stretch         =   -1  'True
      ToolTipText     =   "Heldesk"
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   975
      Index           =   5
      Left            =   1470
      Picture         =   "forma_iconos_Agent.frx":8E268
      Stretch         =   -1  'True
      ToolTipText     =   "Sonar"
      Top             =   2400
      Width           =   945
   End
   Begin VB.Image btn2 
      Height          =   870
      Index           =   4
      Left            =   220
      Picture         =   "forma_iconos_Agent.frx":90F0E
      Stretch         =   -1  'True
      ToolTipText     =   "Eversign"
      Top             =   2565
      Width           =   870
   End
   Begin VB.Image btn2 
      Height          =   855
      Index           =   3
      Left            =   1440
      Picture         =   "forma_iconos_Agent.frx":951A4
      Stretch         =   -1  'True
      ToolTipText     =   "ITC turborater"
      Top             =   1320
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   975
      Index           =   2
      Left            =   180
      Picture         =   "forma_iconos_Agent.frx":97F2A
      Stretch         =   -1  'True
      ToolTipText     =   "LAE System"
      Top             =   1320
      Width           =   975
   End
   Begin VB.Image btn2 
      Height          =   855
      Index           =   1
      Left            =   1470
      Picture         =   "forma_iconos_Agent.frx":9B47B
      Stretch         =   -1  'True
      ToolTipText     =   "JA Login"
      Top             =   120
      Width           =   945
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   120
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hector Navarro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   14880
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created by:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   14880
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v3.2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   13935
      Left            =   1320
      Top             =   -2160
      Width           =   15
   End
   Begin VB.Menu mnufile 
      Caption         =   "Just &Auto"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&End"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuTrayMove 
         Caption         =   "&Move"
      End
      Begin VB.Menu mnuTraySize 
         Caption         =   "&Size"
      End
      Begin VB.Menu mnuTrayMinimize 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu mnuTrayMaximize 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu mnuTraySep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "forma_iconos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ruta_chrome$, SO As Integer, ruta_internet$

Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer

Dim vnc$, seg As Integer, user1$

Public LastState As Integer

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&


Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
  
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
  
Dim OReg As Registro



'Private Declare Function ShellExecute _
'                            Lib "shell32.dll" _
'                            Alias "ShellExecuteA" ( _
'                            ByVal hwnd As Long, _
'                            ByVal lpOperation As String, _
'                            ByVal lpFile As String, _
'                            ByVal lpParameters As String, _
'                            ByVal lpDirectory As String, _
'                            ByVal nShowCmd As Long) _
'                            As Long



 

    Private Const WM_GETTEXT = &HD
    Private Const WM_GETTEXTLENGTH = &HE

    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long



Public Sub Conecta_SQL()
On Error Resume Next
'  Set cn_ptos = New ADODB.Connection
 '  cn_ptos.Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
 
 
 contraseña_ini$ = "Q6XSkLMjy7BUSKdxcE"
 user_ini$ = "payroll"
 bd_ini$ = "laesystemja"
 server_ini$ = "ec2-52-8-179-170.us-west-1.compute.amazonaws.com"   ' "167.114.199.93"  '

 

 With base
   .CursorLocation = adUseClient
   ' .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CallCenter;Data Source=AICO2-HECTOR"
    .Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
   
 End With
End Sub
Public Function HiByte(ByVal wParam As Integer) As Byte
  On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function
Public Function LoByte(ByVal wParam As Integer) As Byte
On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function
Public Sub SocketsCleanup()
On Error Resume Next
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub
Public Function SocketsInitialize() As Boolean
On Error Resume Next

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function


Public Function GetIPHostName() As String
On Error Resume Next
    Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function


Public Function GetIPAddress() As String
On Error Resume Next
   Dim sHostName    As String * 256
   Dim lpHost    As Long
   Dim HOST      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
   End If
    
  'gethostname returns the name of the local host into
  'the buffer specified by the name parameter. The host
  'name is returned as a null-terminated string. The
  'form of the host name is dependent on the Windows
  'Sockets provider - it can be a simple host name, or
  'it can be a fully qualified domain name. However, it
  'is guaranteed that the name returned will be successfully
  'parsed by gethostbyname and WSAAsyncGetHostByName.

  'In actual application, if no local host name has been
  'configured, gethostname must succeed and return a token
  'host name that gethostbyname or WSAAsyncGetHostByName
  'can resolve.
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
              " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   
  'gethostbyname returns a pointer to a HOSTENT structure
  '- a structure allocated by Windows Sockets. The HOSTENT
  'structure contains the results of a successful search
  'for the host specified in the name parameter.

  'The application must never attempt to modify this
  'structure or to free any of its components. Furthermore,
  'only one copy of this structure is allocated per thread,
  'so the application should copy any information it needs
  'before issuing any other Windows Sockets function calls.

  'gethostbyname function cannot resolve IP address strings
  'passed to it. Such a request is treated exactly as if an
  'unknown host name were passed. Use inet_addr to convert
  'an IP address string the string to an actual IP address,
  'then use another function, gethostbyaddr, to obtain the
  'contents of the HOSTENT structure.
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
    
  'to extract the returned IP address, we have to copy
  'the HOST structure and its members
   CopyMemory HOST, lpHost, Len(HOST)
   CopyMemory dwIPAddr, HOST.hAddrList, 4
   
  'create an array to hold the result
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   
  'and with the array, build the actual address,
  'appending a period between members
   For i = 1 To HOST.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
  
  'the routine adds a period to the end of the
  'string, so remove it here
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
    
End Function



Public Sub SetTrayMenuItems(window_state As Integer)
    Select Case window_state
        Case vbMinimized
            mnuTrayMaximize.Enabled = True
            mnuTrayMinimize.Enabled = False
            mnuTrayMove.Enabled = False
            mnuTrayRestore.Enabled = True
            mnuTraySize.Enabled = False
        Case vbMaximized
            mnuTrayMaximize.Enabled = False
            mnuTrayMinimize.Enabled = True
            mnuTrayMove.Enabled = False
            mnuTrayRestore.Enabled = True
            mnuTraySize.Enabled = False
        Case vbNormal
            mnuTrayMaximize.Enabled = True
            mnuTrayMinimize.Enabled = True
            mnuTrayMove.Enabled = True
            mnuTrayRestore.Enabled = False
            mnuTraySize.Enabled = True
    End Select
End Sub


Private Sub btn2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

btn2(Index).Visible = False
seg = 0
Timer1.Enabled = True



End Sub


Private Sub btn2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset

Select Case Index
Case 0
  If btnchrome(0).Value = True Then
     r$ = Shell(ruta_chrome$ + " https://secure5.yourpayrollhr.com/ta/JAI04.clock", vbNormalFocus)
  Else
   r$ = Shell(ruta_internet$ + " https://secure5.yourpayrollhr.com/ta/JAI04.clock", vbNormalFocus)
  End If
     
Case 1
  If btnchrome(1).Value = True Then
  r$ = Shell(ruta_chrome$ + " https://secure5.yourpayrollhr.com/ta/JAI04.login", vbNormalFocus)
  Else
  r$ = Shell(ruta_internet$ + " https://secure5.yourpayrollhr.com/ta/JAI04.login", vbNormalFocus)
  End If
  
Case 2
  If btnchrome(2).Value = True Then
    r$ = Shell(ruta_chrome$ + " https://www.laesystem.com", vbNormalFocus)
  Else
  r$ = Shell(ruta_internet$ + " https://www.laesystem.com", vbNormalFocus)
  End If
  
Case 3
  If btnchrome(3).Value = True Then
  r$ = Shell(ruta_chrome$ + " https://www.turborater.com/login/", vbNormalFocus)
  Else
  r$ = Shell(ruta_internet$ + " https://www.turborater.com/login/", vbNormalFocus)
  End If
  
Case 4
  If btnchrome(4).Value = True Then
  r$ = Shell(ruta_chrome$ + " https://justautoins.eversign.com/dashboard", vbNormalFocus)
  Else
  r$ = Shell(ruta_internet$ + " https://justautoins.eversign.com/dashboard", vbNormalFocus)
  End If
  
Case 5
  If btnchrome(5).Value = True Then
  r$ = Shell(ruta_chrome$ + " dashboard.sendsonar.com/users/sign_in", vbNormalFocus)
  Else
  r$ = Shell(ruta_internet$ + " dashboard.sendsonar.com/users/sign_in", vbNormalFocus)
  End If
  
Case 6
  If btnchrome(6).Value = True Then
  r$ = Shell(ruta_chrome$ + " https://forms.clickup.com/2206515/f/23atk-300/HMZDTQLTRLJ0PYIPS9", vbNormalFocus)
  Else
  r$ = Shell(ruta_internet$ + " https://forms.clickup.com/2206515/f/23atk-300/HMZDTQLTRLJ0PYIPS9", vbNormalFocus)
  End If
  
Case 7
  If btnchrome(7).Value = True Then
  r$ = Shell(ruta_chrome$ + " https://outlook.office.com", vbNormalFocus)
  Else
  r$ = Shell(ruta_internet$ + " https://outlook.office.com", vbNormalFocus)
  End If
  
Case 8

 If Check1.Value = 0 Then

  If btnchrome(8).Value = True Then
    r$ = Shell(ruta_chrome$ + " https://teams.microsoft.com", vbNormalFocus)
  Else
    r$ = Shell(ruta_internet$ + " https://teams.microsoft.com", vbNormalFocus)
  End If
  
 Else
  If Dir$("C:\Users\station\AppData\Local\Microsoft\Teams\Update.exe") <> "" Then
     r$ = Shell("C:\Users\station\AppData\Local\Microsoft\Teams\Update.exe --processStart " + Chr$(34) + "Teams.exe" + Chr$(34), vbNormalFocus)
  Else
     If btnchrome(8).Value = True Then
       r$ = Shell(ruta_chrome$ + " https://teams.microsoft.com", vbNormalFocus)
     Else
       r$ = Shell(ruta_internet$ + " https://teams.microsoft.com", vbNormalFocus)
     End If
  
  End If
 
 End If
  
Case 9
  If btnchrome(9).Value = True Then
    r$ = Shell(ruta_chrome$ + " https://justautoins0-my.sharepoint.com/", vbNormalFocus)
  Else
    r$ = Shell(ruta_internet$ + " https://justautoins0-my.sharepoint.com/", vbNormalFocus)
  End If
  
Case 10
  If btnchrome(10).Value = True Then
  r$ = Shell(ruta_chrome$ + " https://www.office.com/launch/excel?auth=2", vbNormalFocus)
  Else
  r$ = Shell(ruta_internet$ + " https://www.office.com/launch/excel?auth=2", vbNormalFocus)
  End If
  
Case 11
  If btnchrome(11).Value = True Then
  r$ = Shell(ruta_chrome$ + " https://www.office.com/launch/word?auth=2", vbNormalFocus)
  Else
  r$ = Shell(ruta_internet$ + " https://www.office.com/launch/word?auth=2", vbNormalFocus)
  End If
Case 12
 
  If btnchrome(12).Value = True Then
    r$ = Shell(ruta_chrome$ + " https://login.authorize.net/", vbNormalFocus)
  Else
    r$ = Shell(ruta_internet$ + " https://login.authorize.net/", vbNormalFocus)
  End If
  
 Case 13
  
    r$ = Shell("C:\vacations\vacations.exe", vbNormalFocus)
    
 Case 14
 
    r$ = Shell("C:\Discrepancy\discrepancy.exe", vbNormalFocus)
  
 Case 15
 
    r$ = Shell("C:\callcenter\CallCenter.exe", vbNormalFocus)
  
 Case 16
  
     r$ = Shell("C:\Windows\explorer.exe " + Chr$(34) + "C:\transfer" + Chr$(34), vbNormalFocus)
 
 Case 17
 
  r$ = Shell("C:\Windows\system32\calc.exe", vbNormalFocus)
  
 Case 18
 
  r$ = Shell("C:\uw\uw.exe", vbNormalFocus)
  
 Case 19
  
  r$ = Shell("cmd /c taskkill /f /im money.exe")
  
  r$ = Shell("C:\Money\money.exe", vbNormalFocus)
  r$ = Shell("C:\Money\money.exe", vbNormalFocus)
 
  Case 20   ' lastpass
  
    id_employee$ = "59"
    Conecta_SQL
    
    If Left(vnc$, 2) = "84" Then
    
       If Dir$("c:\iconos\oficina.dat") <> "" Then
            nf = FreeFile
            Open "c:\iconos\oficina.dat" For Input Shared As #nf
            Lock #nf
            Line Input #nf, n$
            Unlock #nf
            Close #nf
     
            If n$ = "HAVEN" Then
                sSelect = "SELECT usr, password From appsaccess where idemployee='" + id_employee$ + "' and program='84'"
            Else
                sSelect = "SELECT usr, password From appsaccess where idemployee='" + id_employee$ + "' and program='45'"
            End If
       Else
            sSelect = "SELECT usr, password From appsaccess where idemployee='" + id_employee$ + "' and program='84'"
       
       End If
  
  
    
    Else
 
        sSelect = "SELECT usr, password From appsaccess where idemployee='" + id_employee$ + "' and program='" + Left(vnc$, 2) + "'"
        
    End If
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
      
    usuario$ = Rs(0)
    Contrasena$ = Rs(1)
    
    Rs.Close
    
    MsgBox "The user is: " + UCase(usuario$) + Chr$(13) + "Password: " + Contrasena$ + Chr$(13) + Chr$(13) + "The password was copied to Clipboard. " + Chr$(13) + "Simply, press <Ctrl> + <V> to paste it.", 64, "Attention"
    
    Clipboard.Clear
    Clipboard.SetText Contrasena$
    
    base.Close
    
  
  Case 21
  
     r$ = Shell("C:\dmv\DMV_MOV_PERM.exe", vbNormalFocus)
     
  Case 22
  
    r$ = Shell("C:\Users\" + user1$ + "\AppData\Local\dialpad\Dialpad.exe", vbNormalFocus)
    
   Case 23
  
    r$ = Shell(ruta_chrome$ + " https://dialpad.com/login", vbNormalFocus)
  
 End Select
 
 btn2(Index).Visible = True

 
End Sub

Private Sub btnaspiradora_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnaspiradora.Visible = False
End Sub


Private Sub btnaspiradora_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
' crea directorios

' user1$
 r$ = "c:\users\" + user1$ + "\desktop"
 ' MsgBox r$
 If Dir$(r$ + "\PDFs") = "" Then
    MkDir r$ + "\PDFs"
 End If
 
 If Dir$(r$ + "\Images") = "" Then
   MkDir r$ + "\Images"
 End If
 
 File1.Path = "c:\"
 File1.Pattern = "*.pdf"
 File1.Path = r$ '+ "\PDFs"
 
 For t = 0 To File1.ListCount - 1
   n$ = File1.List(t)
   FileCopy r$ + "\" + n$, r$ + "\PDFs\" + n$
   Kill r$ + "\" + n$
 Next t
 
 
   
   
   
 File1.Path = "c:\"
 File1.Pattern = "*.JPG"
 File1.Path = r$ '+ "\Images"
  
   
 For t = 0 To File1.ListCount - 1
   n$ = File1.List(t)
   FileCopy r$ + "\" + n$, r$ + "\Images\" + n$
   Kill r$ + "\" + n$
 Next t
   
   
   
   
 File1.Path = "c:\"
 File1.Pattern = "*.BMP"
 File1.Path = r$ '+ "\Images"
  
   
 For t = 0 To File1.ListCount - 1
   n$ = File1.List(t)
   FileCopy r$ + "\" + n$, r$ + "\Images\" + n$
   Kill r$ + "\" + n$
 Next t
   
   
   
 File1.Path = "c:\"
 File1.Pattern = "*.PNG"
 File1.Path = r$ '+ "\Images"
  
   
 For t = 0 To File1.ListCount - 1
   n$ = File1.List(t)
   FileCopy r$ + "\" + n$, r$ + "\Images\" + n$
   Kill r$ + "\" + n$
 Next t
   
   
   
   
 File1.Path = "c:\"
 File1.Pattern = "*.GIF"
 File1.Path = r$ '+ "\Images\"
  
   
 For t = 0 To File1.ListCount - 1
   n$ = File1.List(t)
   FileCopy r$ + "\" + n$, r$ + "\Images\" + n$
   Kill r$ + "\" + n$
 Next t




btnaspiradora.Visible = True
End Sub


Private Sub btnbrowser_Click(Index As Integer)
On Error Resume Next


    Dim sChromePath As String
    Dim sProgramFiles As String
    '
    ' check for 32/64 bit version
    '
    sProgramFiles = Environ("ProgramFiles")
    sChromePath = sProgramFiles & "\Google\Chrome\Application\chrome.exe"










If Index = 0 Then
  
  r$ = Shell(ruta_chrome$, vbNormalFocus)
  Clipboard.Clear
  Clipboard.SetText "chrome://settings/content/all"  '+ Chr$(13) + Chr$(10)
  AppActivate r$
  
  Application.SendKeys ("^v~")
  'DoEvents
  
  
   'Dim r As Long
   'r = ShellExecute(0, "open", sChromePath, " chrome://settings/content/all", vbNullString, 1)
   
   ' + " --disable-gpu-vsync "
   'r = Shell(ruta_chrome$ + " --chrome://settings/content/all", vbNormalFocus)
   'AppActivate r
   'SendKeys "chrome://settings/content/all", True
   
   
Else
  r$ = Shell(ruta_internet$ + " edge://settings/siteData", vbNormalFocus)
  Clipboard.Clear
  Clipboard.SetText "edge://settings/siteData"
  
  AppActivate r$
  
  Application.SendKeys ("^v~")
  'DoEvents
  
  
  
  
  
  
  
End If

End Sub


Private Sub btnchrome_Click(Index As Integer)
On Error Resume Next

nf = FreeFile
n$ = "c:\iconos\chk1-" + Format(Index, "00")
Open n$ For Output Shared As #nf
Lock #nf
Print #nf, btnchrome(Index).Value
Unlock #nf
Close #nf

End Sub

Private Sub btncolor_Click(Index As Integer)
On Error Resume Next
forma_iconos.BackColor = btncolor(Index).BackColor
Frame1.BackColor = btncolor(Index).BackColor
Check1.BackColor = btncolor(Index).BackColor

For t = 0 To 12
  Frame2(t).BackColor = btncolor(Index).BackColor
Next t

nf = FreeFile
Open "c:\iconos\backgr" For Output Shared As #nf
Lock #nf
Print #nf, btncolor(Index).BackColor
Print #nf, Str(Index)
Unlock #nf
Close nf

End Sub



Private Sub btnedge_Click(Index As Integer)
On Error Resume Next

nf = FreeFile
n$ = "c:\iconos\chk1-" + Format(Index, "00")
Open n$ For Output Shared As #nf
Lock #nf
Print #nf, btnchrome(Index).Value
Unlock #nf
Close #nf

End Sub

Private Sub btnlocaliza_Click(Index As Integer)

posicion = Index
If Index = 0 Then
Left = Screen.Width - Width
Else
Left = Screen.Width + (Screen.Width - Width)

End If


End Sub

Private Sub btnoffice_Click()
On Error Resume Next

nf = FreeFile
Open "c:\iconos\oficina.dat" For Output Shared As #nf

r$ = MsgBox("Are you assigned to the Covina office?", 4, "Attention")
If r$ = "6" Then
  n$ = "COVINA"
Else
  n$ = "HAVEN"
End If


  Lock #nf
  Print #nf, n$
  Unlock #nf
  
  Close #nf


End Sub

Private Sub btnPasswords_Click()
On Error Resume Next



Load Forma_accesar
Forma_accesar.Show 1

End Sub


Private Sub btnsalir_Click()
On Error Resume Next


End
End Sub





Private Sub Check1_Click()
On Error Resume Next

nf = FreeFile
Open "c:\iconos\Teams.dat" For Output Shared As #nf
Lock #nf
Print #nf, Check1.Value
Unlock #nf
Close #nf




End Sub

Private Sub Form_Load()

On Error Resume Next
Left = (Screen.Width - Width) ' / 2
Top = 0  ' (Screen.Height - Height) - 600

If (App.PrevInstance = True) Then
  'base.Close
  End
End If


If Dir$("c:\iconos\Teams.dat") <> "" Then
  nf = FreeFile
  Open "c:\iconos\Teams.dat" For Input Shared As #nf
  Lock #nf
  Line Input #nf, c
  Unlock #nf
  Close #nf
  Check1.Value = c
  
End If



If Dir$("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") <> "" Then
  SO = 1 ' 64 bits
  ruta_chrome$ = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
ElseIf Dir$("C:\Program Files\Google\Chrome\Application\chrome.exe") <> "" Then
  SO = 0 ' 32 bits
  ruta_chrome$ = "C:\Program Files\Google\Chrome\Application\chrome.exe"
  
 
Else
        If Dir$("c:\iconos\path_goo") <> "" Then

                nf = FreeFile
                Open "c:\iconos\path_goo" For Input Shared As #nf
                Lock #nf
                Line Input #nf, ruta_chrome$
                Unlock #nf
                Close nf
                
                SO = 3
    
        End If


End If



'b$ = GetIPHostName()
a$ = GetIPAddress()


 Dim strUserName As String
 
  strUserName = String(100, Chr$(0))
  'Get the username
  GetUserName strUserName, 100
  'strip the rest of the buffer
  strUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
  
  user1$ = strUserName
 
  
  
  

r$ = ""
valor = 0
For Y = Len(a$) To 1 Step -1
   If Mid$(a$, Y, 1) = "." Then
      If valor <= 1 Then
         r$ = Mid$(a$, Y, 1) + r$
         valor = valor + 1
      End If
   Else
    If valor <= 1 Then
     r$ = Mid$(a$, Y, 1) + r$
    End If
   End If
Next Y

vnc$ = Right(r$, Len(r$) - 1)


lblIP.Caption = "my vnc:  " + vnc$

If Left$(vnc$, 2) = "84" Then
   btnoffice.Visible = True
End If




If Dir$("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe") <> "" Then
  
  'ruta_internet$ = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
  ruta_internet$ = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
Else
  
  ruta_internet$ = "C:\Program Files\Internet Explorer\iexplore.exe"
End If


Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      ' Size of Form in Pixels at design resolution
      
          DesignX = 1280
            
          DesignY = 900 '800
      
      
      RePosForm = True   ' Flag for positioning Form
      DoResize = False   ' Flag for Resize Event
      ' Set up the screen values
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
      Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
      If DesignX = 800 Then
        ScaleFactorX = (Xpixels / DesignX)  ' 0.78
        ScaleFactorY = (Ypixels / DesignY)  ' 0.78
      Else
               
        
        If Xpixels <= 1366 Then  ' Si es laptop
      
           ScaleFactorX = 1360 / DesignX  ' 1360
           ScaleFactorY = 680 / DesignY   ' 1024
        
        Else  ' Si es Desktop con monitor de alta resolucion
          
           ScaleFactorX = 1280 / DesignX '1360 / DesignX
           ScaleFactorY = 860 / DesignY ' 1024 / DesignY
                
        End If
        
        
      End If
      
      ScaleMode = 1  ' twips
      'Exit Sub  ' uncomment to see how Form1 looks without resizing
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      'Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
       '"  by " + Str$(Ypixels)
      If DesignX = 800 Then
        forma_main.Height = 9000 'Me.Height ' Remember the current size
        forma_main.Width = 12000 'Me.Width
      Else
        Height = Me.Height ' Remember the current size
        Width = Me.Width
      
      End If
primeravez = 0


carga_valores




If Dir$("c:\iconos\backgr") <> "" Then

  nf = FreeFile
  Open "c:\iconos\backgr" For Input Shared As #nf
  Lock #nf
  Line Input #nf, r$
  Line Input #nf, r2$
  Unlock #nf
  Close nf

  forma_iconos.BackColor = r$
  btncolor(Val(r2$)).Value = True
  Frame1.BackColor = r$

  For t = 0 To 12
     Frame2(t).BackColor = r$
  Next t

End If

Set OReg = New Registro
Call OReg.EstablecerValor(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "BARRA_AGENT", "c:\iconos\barra_agent.exe", REG_SZ)





If WindowState = vbMinimized Then
        LastState = vbNormal
Else
        LastState = WindowState
End If

    AddToTray Me, mnuTray

    SetTrayTip "Agent Navigation Bar"
End Sub


Private Sub Form_Resize()
On Error Resume Next
On Error Resume Next
Dim ScaleFactorX As Single, ScaleFactorY As Single

If primeravez = 0 Then


primeravez = 1
      If Not DoResize Then  ' To avoid infinite loop
         DoResize = True
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
      ScaleFactorY = Me.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width
End If
primeravez = 1

SetTrayMenuItems WindowState

    If WindowState <> vbMinimized Then _
        LastState = WindowState
End Sub




Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image1.Visible = False
End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

pos = InStr(1, vnc$, ".")
r$ = Left(vnc$, pos - 1)

Select Case Val(r$)
Case 39 ' Arleta
  If btnchrome(8).Value = True Then
     a$ = Shell(ruta_chrome$ + " https://teams.microsoft.com/_#/conversations/19:8503da54405d408991204acf81b0a02d@thread.v2?ctx=chat", vbNormalFocus)
  Else
     a$ = Shell(ruta_internet$ + " https://teams.microsoft.com/_#/conversations/19:8503da54405d408991204acf81b0a02d@thread.v2?ctx=chat", vbNormalFocus)
  End If


Case 43 ' Compton
  If btnchrome(8).Value = True Then
     a$ = Shell(ruta_chrome$ + " https://teams.microsoft.com/_#/conversations/19:4a76a34081d84f5cbdc3706879008984@thread.v2?ctx=chat", vbNormalFocus)
  Else
     a$ = Shell(ruta_internet$ + " https://teams.microsoft.com/_#/conversations/19:4a76a34081d84f5cbdc3706879008984@thread.v2?ctx=chat", vbNormalFocus)
  End If

Case 84  ' Haven

  If Dir$("c:\iconos\oficina.dat") <> "" Then
     nf = FreeFile
     Open "c:\iconos\oficina.dat" For Input Shared As #nf
     Lock #nf
     Line Input #nf, n$
     Unlock #nf
     Close #nf
     
     If n$ = "HAVEN" Then
          If btnchrome(8).Value = True Then
             a$ = Shell(ruta_chrome$ + " https://teams.microsoft.com/_#/conversations/19:fdcadd3f75dc465e8bf2211340a4bd24@thread.v2?ctx=chat", vbNormalFocus)
          Else
             a$ = Shell(ruta_internet$ + " https://teams.microsoft.com/_#/conversations/19:fdcadd3f75dc465e8bf2211340a4bd24@thread.v2?ctx=chat", vbNormalFocus)
          End If
     
     Else
          If btnchrome(8).Value = True Then
             a$ = Shell(ruta_chrome$ + " https://teams.microsoft.com/_#/conversations/19:3ae64664f20e47558ff2c18d91147963@thread.v2?ctx=chat", vbNormalFocus)
          Else
             a$ = Shell(ruta_internet$ + " https://teams.microsoft.com/_#/conversations/19:3ae64664f20e47558ff2c18d91147963@thread.v2?ctx=chat", vbNormalFocus)
          End If
     
     End If
  Else
  
    If btnchrome(8).Value = True Then
       a$ = Shell(ruta_chrome$ + " https://teams.microsoft.com/_#/conversations/19:fdcadd3f75dc465e8bf2211340a4bd24@thread.v2?ctx=chat", vbNormalFocus)
    Else
       a$ = Shell(ruta_internet$ + " https://teams.microsoft.com/_#/conversations/19:fdcadd3f75dc465e8bf2211340a4bd24@thread.v2?ctx=chat", vbNormalFocus)
    End If
  
  
  End If


  
Case 49  ' Florence
  If btnchrome(8).Value = True Then
     a$ = Shell(ruta_chrome$ + " https://teams.microsoft.com/_#/conversations/19:04f68636fe5f49b0b40a86385513c958@thread.v2?ctx=chat", vbNormalFocus)
  Else
     a$ = Shell(ruta_internet$ + " https://teams.microsoft.com/_#/conversations/19:04f68636fe5f49b0b40a86385513c958@thread.v2?ctx=chat", vbNormalFocus)
  End If
  
Case 46  ' Ponderosa
  If btnchrome(8).Value = True Then
     a$ = Shell(ruta_chrome$ + " https://teams.microsoft.com/_#/conversations/19:preview-f400c577-ae06-4c7c-b4c8-d6a63999b36c?ctx=chat", vbNormalFocus)
  Else
     a$ = Shell(ruta_internet$ + " https://teams.microsoft.com/_#/conversations/19:preview-f400c577-ae06-4c7c-b4c8-d6a63999b36c?ctx=chat", vbNormalFocus)
  End If
  
Case 47  ' SB
  If btnchrome(8).Value = True Then
     a$ = Shell(ruta_chrome$ + " https://teams.microsoft.com/_#/conversations/19:2baf76aa9aa4492899dbae06993dc932@thread.v2?ctx=chat", vbNormalFocus)
  Else
     a$ = Shell(ruta_internet$ + " https://teams.microsoft.com/_#/conversations/19:2baf76aa9aa4492899dbae06993dc932@thread.v2?ctx=chat", vbNormalFocus)
  End If
  
Case 41  ' Santa Ana
  If btnchrome(8).Value = True Then
     a$ = Shell(ruta_chrome$ + " https://teams.microsoft.com/_#/conversations/19:8534130428e140d78373163555a127c8@thread.v2?ctx=chat", vbNormalFocus)
  Else
     a$ = Shell(ruta_internet$ + " https://teams.microsoft.com/_#/conversations/19:8534130428e140d78373163555a127c8@thread.v2?ctx=chat", vbNormalFocus)
  End If
  
Case 23  ' Whittier
  If btnchrome(8).Value = True Then
     a$ = Shell(ruta_chrome$ + " https://teams.microsoft.com/_#/conversations/19:6136fdf5c8444c5993f2a517ce619d95@thread.v2?ctx=chat", vbNormalFocus)
  Else
     a$ = Shell(ruta_internet$ + " https://teams.microsoft.com/_#/conversations/19:6136fdf5c8444c5993f2a517ce619d95@thread.v2?ctx=chat", vbNormalFocus)
  End If
  
  
End Select




  
Image1.Visible = True
End Sub


Private Sub Image2_Click()
End
End Sub



Private Sub mnuFileExit_Click()
    forma_iconos.WindowState = 1
End
End Sub

Private Sub mnuTrayClose_Click()
    Unload Me
End Sub


Private Sub mnuTrayMaximize_Click()
On Error Resume Next

    WindowState = vbMaximized
    If Not Visible Then Me.Show
End Sub


Private Sub mnuTrayMinimize_Click()
    WindowState = vbMinimized
End Sub


Private Sub mnuTrayMove_Click()
On Error Resume Next

    SendMessage hwnd, WM_SYSCOMMAND, _
        SC_MOVE, 0&
End Sub


Private Sub mnuTrayRestore_Click()
On Error Resume Next

    SendMessage hwnd, WM_SYSCOMMAND, _
        SC_RESTORE, 0&
End Sub


Private Sub mnuTraySize_Click()
On Error Resume Next

    SendMessage hwnd, WM_SYSCOMMAND, _
        SC_SIZE, 0&
End Sub



Public Sub carga_valores()
On Error Resume Next

For t = 0 To 12

 
  
  
    
  nf = FreeFile
  n$ = "c:\iconos\chk1-" + Format(t, "00")
  Open n$ For Input Shared As #nf
  Lock #nf
  Line Input #nf, b$
  Unlock #nf
  Close #nf
  
  btnchrome(t).Value = b$
  
  If UCase(b$) = "FALSE" Then
    btnedge(t).Value = True
  Else
    btnedge(t).Value = False
  End If
  
  
  
Next t

  
End Sub


Public Sub Obtener_titulo_barra()
Dim dhWnd As Long
        Dim chWnd As Long

        Dim Web_Caption As String * 256
        Dim Web_hWnd As Long

        Dim URL As String * 256
        Dim URL_hWnd As Long

        dhWnd = GetDesktopWindow
        chWnd = FindWindowEx(dhWnd, 0, "Chrome_WidgetWin_1", vbNullString)
        Web_hWnd = FindWindowEx(dhWnd, chWnd, "Chrome_WidgetWin_1", vbNullString)
        URL_hWnd = FindWindowEx(Web_hWnd, 0, "Chrome_OmniboxView", vbNullString)

        Call SendMessage(Web_hWnd, WM_GETTEXT, 256, ByVal Web_Caption)
        Call SendMessage(URL_hWnd, WM_GETTEXT, 256, ByVal URL)

        MsgBox Split(Web_Caption, Chr(0))(0) & vbCrLf & Split(URL, Chr(0))(0)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
seg = seg + 1

If seg >= 3 Then
   
   For t = 0 To 21
      btn2(t).Visible = True
   Next t
   seg = 0
   Timer1.Enabled = False

End If


End Sub


