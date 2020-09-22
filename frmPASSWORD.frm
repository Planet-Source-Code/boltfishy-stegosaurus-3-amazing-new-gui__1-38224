VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPASSWORD 
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
   Begin VB.CommandButton cmdBACK 
      Caption         =   "<< &Back"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgBAR 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   32
      Scrolling       =   1
   End
   Begin VB.TextBox txtPASSWORD 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1680
      Width           =   5535
   End
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdNEXT 
      Caption         =   "&Next >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblTITLE 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password Entry."
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
      TabIndex        =   4
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image imgICON 
      Height          =   480
      Left            =   120
      Picture         =   "frmPASSWORD.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblMSG 
      Caption         =   "Please supply the password to use for this image."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label lblSUBTITLE 
      BackStyle       =   0  'Transparent
      Caption         =   "Step 3."
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
      TabIndex        =   5
      Top             =   960
      Width           =   5535
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
Attribute VB_Name = "frmPASSWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngACTIONTYPE As ActionTypes
'------------------------------------------------------------
' Author:  Clint M. LaFever - [lafeverc@saic.com]
' Date: August,23 2002 @ 11:26:51
'------------------------------------------------------------
Public Sub LoadIt(at As ActionTypes)
    On Error GoTo ErrorLoadIt
    mlngACTIONTYPE = at
    Exit Sub
ErrorLoadIt:
    MsgBox Err & ":Error in call to LoadIt()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub cmdBACK_Click()
    On Error Resume Next
    Dim frm As frmCHOOSE
    Set frm = New frmCHOOSE
    frm.LoadIt mlngACTIONTYPE
    frm.Show
    Unload Me
End Sub
Private Sub cmdCANCEL_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Sub cmdNEXT_Click()
    On Error Resume Next
    Dim frm As frmMSG
    Set frm = New frmMSG
    frm.LoadIt mlngACTIONTYPE, Me.txtPASSWORD.Text
    frm.Show
    Unload Me
End Sub
Private Sub txtPASSWORD_Change()
    On Error Resume Next
    Me.prgBAR.Value = Len(txtPASSWORD.Text)
    If Len(txtPASSWORD.Text) > 0 Then
        Me.cmdNEXT.Enabled = True
    Else
        Me.cmdNEXT.Enabled = False
    End If
End Sub
