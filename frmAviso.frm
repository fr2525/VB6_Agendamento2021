VERSION 5.00
Begin VB.Form frmAviso 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aviso:"
   ClientHeight    =   4440
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9795
   Icon            =   "frmAviso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4260
      MaskColor       =   &H8000000A&
      TabIndex        =   0
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Pet Shop dos Pets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3165
      TabIndex        =   5
      Top             =   600
      Width           =   3285
   End
   Begin VB.Label lblAviso4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   9555
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAviso3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2340
      Width           =   9555
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAviso2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   9555
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAviso1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   9555
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
    Unload Me
End Sub
