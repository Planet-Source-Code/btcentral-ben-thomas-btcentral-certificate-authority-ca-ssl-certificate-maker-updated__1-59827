VERSION 5.00
Begin VB.Form mkEmail 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email Certificate Maker"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5550
   Icon            =   "mkEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin CA.DMSXpButton Command2 
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Step 2"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Command1 
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Step 1"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CA.DMSXpButton Command3 
      Height          =   255
      Left            =   2310
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Back"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To install the email certificate please double click email.p12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click the button step 2 to create the email certificate"
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
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click the button step 1 to generate the required files."
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "mkEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim cmd As String
    cmd = "./openssl.exe ca -config CA.ini -msie_hack -out emailcert.cer -infiles emailreq.pem"
    If RunProcess(cmd, SW_NORMAL) Then
        Command1.Enabled = True
    Else
        MsgBox "Error, Email Certificate not made!"
    End If
End Sub

Private Sub Command2_Click()
    Dim cmd2 As String
    cmd2 = "./openssl.exe pkcs12 -export -out email.p12 -in emailcert.cer -inkey emailkey.pem"
    If RunProcess(cmd2, SW_NORMAL) Then
        Command1.Enabled = True
    Else
        MsgBox "Error, Email Certificate not made!"
    End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
