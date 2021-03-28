VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form2"
   ClientHeight    =   3120
   ClientLeft      =   5172
   ClientTop       =   3012
   ClientWidth     =   6252
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   13.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3120
   ScaleWidth      =   6252
   Begin Project1.WLabel WLabel5 
      Height          =   1524
      Left            =   4500
      TabIndex        =   4
      Top             =   828
      Width           =   516
      _ExtentX        =   910
      _ExtentY        =   2688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "WLabel5"
      Orientation     =   3
   End
   Begin Project1.WLabel WLabel4 
      Height          =   696
      Left            =   1656
      TabIndex        =   3
      Top             =   1656
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1228
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "WLabel4"
      Orientation     =   2
   End
   Begin Project1.WLabel WLabel3 
      Height          =   1560
      Left            =   3276
      TabIndex        =   2
      Top             =   792
      Width           =   948
      _ExtentX        =   1672
      _ExtentY        =   2752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "WLabel3"
      Alignment       =   1
      Orientation     =   1
   End
   Begin Project1.WLabel WLabel2 
      Height          =   480
      Left            =   1908
      TabIndex        =   1
      Top             =   792
      Width           =   1236
      _ExtentX        =   2180
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "WLabel2"
   End
   Begin Project1.WLabel WLabel1 
      Height          =   1572
      Left            =   648
      TabIndex        =   0
      Top             =   720
      Width           =   420
      _ExtentX        =   550
      _ExtentY        =   1947
      BackColor       =   -2147483646
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "WLabel1"
      AutoSize        =   -1  'True
      Orientation     =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserControl11_Click()
    Caption = Rnd
End Sub

Private Sub Command1_Click()
    Label1.Caption = "22"
    VLabel1.Caption = "22"
End Sub

Private Sub Label1_Change()
Stop
End Sub

Private Sub VLabel1_Change()
Stop
End Sub

