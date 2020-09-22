VERSION 5.00
Begin VB.Form frmSTART 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stegosaurus Step-by-Step by Clint M. LaFever and Mischa Balen"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optDECRYPT 
      Caption         =   "Extract (decrypt) a message."
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   5415
   End
   Begin VB.OptionButton optENCRYPT 
      Caption         =   "Create (encrypt) a message."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Value           =   -1  'True
      Width           =   5415
   End
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdNEXT 
      Caption         =   "&Next >>"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblMSG 
      Caption         =   "Please indicate which action you wish to perform."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label lblSUBTITLE 
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label lblTITLE 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Stegosaurus."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image imgICON 
      Height          =   480
      Left            =   120
      Picture         =   "frmSTART.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape shpBOX 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmSTART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCANCEL_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdNEXT_Click()
    On Error Resume Next
    Dim frm As frmCHOOSE
    Set frm = New frmCHOOSE
    If Me.optENCRYPT.Value = True Then
        frm.LoadIt atENCRYPT
    Else
        frm.LoadIt atDECRYPT
    End If
    frm.Show
    Unload Me
End Sub
