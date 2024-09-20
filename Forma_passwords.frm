VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Forma_passwords 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Passwords"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5520
   ControlBox      =   0   'False
   Icon            =   "Forma_passwords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   4935
      Begin Project1.lvButtons_H btnborrar1 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
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
         Mode            =   0
         Value           =   0   'False
         Image           =   "Forma_passwords.frx":0442
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btncopy 
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Caption         =   "Copy"
         CapAlign        =   2
         BackStyle       =   3
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
      Begin Project1.lvButtons_H btnsave 
         Height          =   495
         Left            =   4200
         TabIndex        =   7
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         Caption         =   "Save"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin VB.TextBox txtpassword 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtuser 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         MaxLength       =   60
         TabIndex        =   5
         Top             =   480
         Width           =   2895
      End
      Begin Project1.lvButtons_H btncopy2 
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Caption         =   "Copy"
         CapAlign        =   2
         BackStyle       =   3
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
      Begin Project1.lvButtons_H btnborrar2 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
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
         Mode            =   0
         Value           =   0   'False
         Image           =   "Forma_passwords.frx":0DA4
         cBack           =   12632256
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
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
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000A&
      Height          =   1575
      Left            =   3720
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   16
      Top             =   120
      Width           =   1575
      Begin VB.Image Image1 
         Height          =   1335
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   1335
      TabIndex        =   15
      Top             =   3720
      Width           =   1335
   End
   Begin Project1.lvButtons_H btncerrar 
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   3960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      Caption         =   "OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ImageList im1 
      Left            =   2040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":1706
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":5866
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":9BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":D155
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":FEEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":14191
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":16E47
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":1942A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":1B91D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":1CB60
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":1ECDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":227EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":2620D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":2A1FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":2DD85
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":316A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forma_passwords.frx":33897
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboprograms 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblid 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "IDEmp:"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program:"
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
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   660
   End
End
Attribute VB_Name = "Forma_passwords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
Private Sub btnborrar1_Click()
On Error Resume Next
txtuser.Text = ""
txtuser.SetFocus
End Sub

Private Sub btnborrar2_Click()
On Error Resume Next
txtpassword.Text = ""
txtpassword.SetFocus
End Sub


Private Sub btncerrar_Click()
On Error Resume Next

base.Close
Unload Me
End Sub

Private Sub btncopy_Click()
On Error Resume Next

Clipboard.Clear
  Clipboard.SetText txtuser.Text
End Sub


Private Sub btncopy2_Click()
On Error Resume Next

Clipboard.Clear
  Clipboard.SetText txtpassword.Text
End Sub


Private Sub btnsave_Click()
On Error Resume Next
If cboprograms.ListIndex = -1 Then
   Exit Sub
End If


Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    
  id_employee$ = lblid.Caption
  
  Conecta_SQL
 
    
sSelect = "SELECT idappsaccess From appsaccess where idemployee='" + id_employee$ + "' and program='" + Format(cboprograms.ListIndex, "#0") + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_apps$ = Rs(0)
    
    Rs.Close
    
    
    
    
    
    If id_apps$ = "" Then
         
         sSelect = "insert into appsaccess (idemployee, program, usr, password)  VALUES ('" & _
         id_employee$ + "', '" + Format(cboprograms.ListIndex, "#0") + "', '" + txtuser.Text + "', '" + txtpassword.Text + "')"
    
         Rs.Open sSelect, base, adOpenUnspecified
    
         Rs.Close
         
     Else
        
         sSelect = "update appsaccess set idemployee='" + id_employee$ + "', program='" + Format(cboprograms.ListIndex, "#0") + "', usr='" + txtuser.Text + "', Password='" & _
         txtpassword.Text + "' where idappsaccess='" + id_apps$ + "'"

         Rs.Open sSelect, base, adOpenUnspecified
         Rs.Close
        
     End If
     
    
cboprograms.ListIndex = -1
txtuser.Text = ""
txtpassword.Text = ""
Image1.Picture = LoadPicture()

base.Close
End Sub


Private Sub cboprograms_Click()
On Error Resume Next

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    
txtuser.Text = ""
txtpassword.Text = ""
    

Image1.Picture = im1.ListImages(cboprograms.ListIndex + 1).Picture

id_employee$ = lblid.Caption

 Conecta_SQL
 
 sSelect = "SELECT usr, password From appsaccess where idemployee='" + id_employee$ + "' and program='" + Format(cboprograms.ListIndex, "#0") + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    usuario$ = Rs(0)
    Contrasena$ = Rs(1)
    
    Rs.Close


   txtuser.Text = RTrim(usuario$)
   txtpassword.Text = RTrim(Contrasena$)


txtuser.SetFocus

base.Close


End Sub


Private Sub Form_Load()
On Error Resume Next

cboprograms.Clear

cboprograms.AddItem "Clock in"
cboprograms.AddItem "JA Login"
cboprograms.AddItem "LAE System"
cboprograms.AddItem "ITC Turborater"
cboprograms.AddItem "Eversign"
cboprograms.AddItem "Sonar"
cboprograms.AddItem "Email Outlook"
cboprograms.AddItem "Teams"
cboprograms.AddItem "Onedrive"
cboprograms.AddItem "Authorize"
cboprograms.AddItem "Vacation Program"
cboprograms.AddItem "Discrepancy Program"
cboprograms.AddItem "Appointment Scheduler"
cboprograms.AddItem "UW Program"
cboprograms.AddItem "Money Report Program"
cboprograms.AddItem "LastPass"
cboprograms.AddItem "Dialpad"


Top = 0

If posicion = 0 Then
  Left = Screen.Width - Width
Else
  Left = Screen.Width + (Screen.Width - Width)

End If


lblid.Caption = transfiere$

Conecta_SQL

End Sub
