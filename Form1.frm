VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Menggunakan Property Mask di MaskEdBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Not IsDate(MaskEdBox1.Text) Then
     MsgBox "Tanggal tidak valid!", vbCritical, _
            "Tidak Valid"
     MaskEdBox1.SetFocus
     SendKeys "{Home}+{End}"
     Exit Sub
  End If
  MsgBox MaskEdBox1.Text
  MsgBox MaskEdBox2.Text
  MsgBox MaskEdBox3.Text
End Sub

Private Sub Form_Load()
   MaskEdBox1.Mask = "##-##-####"    'Tanggal
  'Bisa juga dengan:
  'MaskEdBox1.Mask = "99-99-9999"   'Tanggal
   MaskEdBox2.Mask = "(###)-#######" 'Nomor telepon
  'Bisa juga dengan:
  'MaskEdBox2.Mask = "(999)-9999999" 'Nomor telepon
   MaskEdBox3.Mask = "????????"  'Hanya karakter huruf
  'sebanyak 8 karakter, dan tidak boleh ada
  'mengandung spasi
End Sub



