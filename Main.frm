VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BTCentral CA"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6975
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin CA.DMSXpButton Command7 
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sign Application Signing Certificate"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Aboutcmd 
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "About"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Command8 
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Revoke Certificate(s)"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Command5 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sign Email Certificate"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Command4 
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Email Certificate"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Command3 
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sign Server Certificate"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Command2 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Server Certificate"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Command1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Generate CA Certificate"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Exitcmd 
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Command6 
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Application Signing Certificate"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Revoke a certificate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create App signing certificate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   16
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Configure App signing certificate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create Email certificate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Configure Email certificate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create Server certificate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Configure server certificate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create the CA certificate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Aboutcmd_Click()
Unload Me
Main.Show
frmAbout.Show
End Sub

Private Sub Command1_Click()
Unload Me
Main.Show
CAConfig.Show
End Sub

Private Sub Command2_Click()
Unload Me
Main.Show
ServerConfig.Show
End Sub

Private Sub Command3_Click()
    Dim cmd As String
    cmd = "./openssl.exe ca -config CA.ini -msie_hack -out servercert.cer -infiles serverreq.pem"
    If RunProcess(cmd, SW_NORMAL) Then
        Command1.Enabled = True
    Else
        MsgBox "Error, Server Certificate not made!"
    End If
End Sub

Private Sub Command4_Click()
Unload Me
Main.Show
EmailConfig.Show
End Sub

Private Sub Command5_Click()
Unload Me
Main.Show
mkEmail.Show
End Sub

Private Sub Command6_Click()
Unload Me
Main.Show
AppConfig.Show
End Sub

Private Sub Command7_Click()
    Dim cmd As String
    cmd = "./openssl.exe ca -config CA.ini -msie_hack -out appcert.cer -infiles appreq.pem"
    If RunProcess(cmd, SW_NORMAL) Then
        Command1.Enabled = True
    Else
        MsgBox "Error, Application Signing Certificate not made!"
    End If
End Sub

Private Sub Command8_Click()
Unload Me
Main.Show
Revoke.Show
End Sub

Private Sub Exitcmd_Click()
End
End Sub
