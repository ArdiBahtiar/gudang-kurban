VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Aplikasi Penghitung Stok"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16170
   FillColor       =   &H00008000&
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   16800
      TabIndex        =   34
      Text            =   "174"
      Top             =   7920
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   16680
      TabIndex        =   33
      Text            =   "10"
      Top             =   8640
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   16800
      TabIndex        =   32
      Text            =   "13"
      Top             =   9480
      Width           =   1815
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Ambil"
      Height          =   615
      Left            =   18840
      TabIndex        =   31
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Ambil (1)"
      Height          =   615
      Left            =   18840
      TabIndex        =   30
      Top             =   8760
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Ambil"
      Height          =   615
      Left            =   18840
      TabIndex        =   29
      Top             =   9600
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Ambil"
      Height          =   615
      Left            =   11160
      TabIndex        =   28
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Ambil (4)"
      Height          =   615
      Left            =   12240
      TabIndex        =   27
      Top             =   8880
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Ambil (1)"
      Height          =   615
      Left            =   11160
      TabIndex        =   26
      Top             =   8880
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Ambil"
      Height          =   615
      Left            =   11160
      TabIndex        =   25
      Top             =   8040
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   8520
      TabIndex        =   24
      Text            =   "80"
      Top             =   9600
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   8520
      TabIndex        =   23
      Text            =   "48"
      Top             =   8760
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   8520
      TabIndex        =   22
      Text            =   "12"
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton Command7 
      Caption         =   "+5"
      Height          =   615
      Left            =   2520
      TabIndex        =   17
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "+2"
      Height          =   615
      Left            =   1440
      TabIndex        =   16
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "+1"
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+100"
      Height          =   615
      Left            =   3600
      TabIndex        =   14
      Top             =   9360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+50"
      Height          =   615
      Left            =   2520
      TabIndex        =   13
      Top             =   9360
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+20"
      Height          =   615
      Left            =   1440
      TabIndex        =   12
      Top             =   9360
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+10"
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   9360
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   8400
      TabIndex        =   7
      Text            =   "0"
      Top             =   5160
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   8400
      TabIndex        =   6
      Text            =   "0"
      Top             =   4080
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   8400
      TabIndex        =   5
      Text            =   "1953"
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   10
      X1              =   5880
      X2              =   19800
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   10
      X1              =   5400
      X2              =   5400
      Y1              =   7320
      Y2              =   10200
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Balungan :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   13320
      TabIndex        =   37
      Top             =   7800
      Width           =   3495
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Buntut     :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   13320
      TabIndex        =   36
      Top             =   8640
      Width           =   3495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Gajih       :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   13320
      TabIndex        =   35
      Top             =   9480
      Width           =   3495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Babat   :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   5760
      TabIndex        =   21
      Top             =   9600
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Kaki     :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   5760
      TabIndex        =   20
      Top             =   8760
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Kepala :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   5760
      TabIndex        =   19
      Top             =   7920
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Lain - Lain :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   5760
      TabIndex        =   18
      Top             =   7080
      Width           =   4095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bungkus"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   12360
      TabIndex        =   10
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bungkus"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   12360
      TabIndex        =   9
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Bungkus"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   12360
      TabIndex        =   8
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Stok Keluar                 :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   720
      TabIndex        =   4
      Top             =   3960
      Width           =   7215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Stok Gudang Saat Ini :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   5160
      Width           =   7215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Stok Awal                    :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Top             =   2760
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Yayasan Al-Maghfirah Surabaya"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   5640
      TabIndex        =   1
      Top             =   1080
      Width           =   10335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Panitia Qurban 1441H"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   6960
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = Val(Text2.Text) + 10
Text3.Text = Val(Text1.Text) - Val(Text2.Text)

End Sub

Private Sub Command10_Click()
Text5.Text = Val(Text5.Text) - 4
End Sub

Private Sub Command11_Click()
Text6.Text = Val(Text6.Text) - 1
End Sub

Private Sub Command12_Click()
Text7.Text = Val(Text7.Text) - 1
End Sub

Private Sub Command14_Click()
Text8.Text = Val(Text8.Text) - 1
End Sub

Private Sub Command15_Click()
Text9.Text = Val(Text9.Text) - 1
End Sub

Private Sub Command2_Click()
Text2.Text = Val(Text2.Text) + 20
Text3.Text = Val(Text1.Text) - Val(Text2.Text)

End Sub

Private Sub Command3_Click()
Text2.Text = Val(Text2.Text) + 50
Text3.Text = Val(Text1.Text) - Val(Text2.Text)

End Sub

Private Sub Command4_Click()
Text2.Text = Val(Text2.Text) + 100
Text3.Text = Val(Text1.Text) - Val(Text2.Text)

End Sub

Private Sub Command5_Click()
Text2.Text = Val(Text2.Text) + 1
Text3.Text = Val(Text1.Text) - Val(Text2.Text)
End Sub

Private Sub Command6_Click()
Text2.Text = Val(Text2.Text) + 2
Text3.Text = Val(Text1.Text) - Val(Text2.Text)
End Sub

Private Sub Command7_Click()
Text2.Text = Val(Text2.Text) + 5
Text3.Text = Val(Text1.Text) - Val(Text2.Text)
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Command8_Click()
Text4.Text = Val(Text4.Text) - 1
End Sub

Private Sub Command9_Click()
Text5.Text = Val(Text5.Text) - 1
End Sub

