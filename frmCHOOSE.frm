VERSION 5.00
Begin VB.Form frmCHOOSE 
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
   Begin VB.CommandButton cmdBROWSE 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtFILE 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   5535
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
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblTITLE 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Locate Image File."
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
      TabIndex        =   6
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image imgICON 
      Height          =   480
      Left            =   120
      Picture         =   "frmCHOOSE.frx":0000
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
   Begin VB.Label lblSUBTITLE 
      BackStyle       =   0  'Transparent
      Caption         =   "Step 2."
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
   Begin VB.Label lblMSG 
      Caption         =   "Please locate and select the image file to use for this action."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5760
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmCHOOSE"
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
    frmSTART.Show
    Unload Me
End Sub

Private Sub cmdBROWSE_Click()
    On Error Resume Next
    Dim obj As CDLG, fNAME As String
    Set obj = New CDLG
    If mlngACTIONTYPE = atENCRYPT Then
        obj.VBGetOpenFileName fNAME, , , , , , "Image Files|*.jpg;*.bmp|All Files|*.*", , CurDir, "Locate Image", , Me.hwnd
    Else
        obj.VBGetOpenFileName fNAME, , , , , , "Bitmap Files|*.bmp", , CurDir, "Locate Image", "*.bmp", Me.hwnd
    End If
    If fNAME <> "" Then
        Me.txtFILE.Text = fNAME
    End If
End Sub
Private Sub cmdCANCEL_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Sub cmdNEXT_Click()
    On Error Resume Next
    Dim frm As frmPASSWORD
    Set TheImage = LoadPicture(Me.txtFILE.Text)
    Set frm = New frmPASSWORD
    frm.LoadIt mlngACTIONTYPE
    frm.Show
    Unload Me
End Sub
Private Sub txtFILE_Change()
    On Error Resume Next
    If txtFILE.Text = "" Then
        Me.cmdNEXT.Enabled = False
    Else
        Me.cmdNEXT.Enabled = True
    End If
End Sub
