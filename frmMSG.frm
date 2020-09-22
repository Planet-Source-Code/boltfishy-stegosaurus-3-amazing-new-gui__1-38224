VERSION 5.00
Begin VB.Form frmMSG 
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
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton cmdBACK 
      Caption         =   "<< &Back"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtMSG 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1680
      Width           =   5535
   End
   Begin VB.CommandButton cmdFINISH 
      Caption         =   "&Finish"
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
      Caption         =   "Message."
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
      Picture         =   "frmMSG.frx":0000
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
      Caption         =   "Step 4."
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
      TabIndex        =   6
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label lblMSG 
      Caption         =   "Please supply the message to encrypt into the selected image."
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
      TabIndex        =   5
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
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngACTIONTYPE As ActionTypes
Private mstrPASS As String
'------------------------------------------------------------
' Author:  Clint M. LaFever - [lafeverc@saic.com]
' Date: August,23 2002 @ 11:26:51
'------------------------------------------------------------
Public Sub LoadIt(at As ActionTypes, p As String)
    On Error GoTo ErrorLoadIt
    mlngACTIONTYPE = at
    mstrPASS = p
    Exit Sub
ErrorLoadIt:
    MsgBox Err & ":Error in call to LoadIt()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub cmdBACK_Click()
    On Error Resume Next
    Dim frm As frmPASSWORD
    Set frm = New frmPASSWORD
    frm.LoadIt mlngACTIONTYPE
    frm.Show
    Unload Me
End Sub
Private Sub cmdCANCEL_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdFINISH_Click()
    On Error Resume Next
    If mlngACTIONTYPE = atENCRYPT Then
        EncodeIt txtMSG.Text
    End If
    Unload Me
End Sub
Private Sub Form_Load()
    On Error Resume Next
    With Me.picImage
        .AutoRedraw = True
        .ScaleMode = vbPixels
    End With
    Set Me.picImage.Picture = TheImage
    If mlngACTIONTYPE = atENCRYPT Then
        Me.lblMSG.Caption = "Please supply the message to encrypt into the selected image."
        Me.lblTITLE.Caption = "Message to encrypt."
    Else
        Me.lblMSG.Caption = "Below is the hidden message in the selected image."
        Me.lblTITLE.Caption = "Decyrpted Message."
        DecodeIt
    End If
End Sub
Private Sub EncodeIt(m As String) 'encode file
On Error GoTo ErrSub
Dim msg As String
Dim i As Integer
Dim used_positions As Collection
Dim iWidth As Integer
Dim iHeight As Integer
Dim show_pixels As Boolean
Me.txtMSG.Text = cRC4(txtMSG.Text, mstrPASS)
Rnd -1
Randomize NumericPassword(mstrPASS)

iWidth = picImage.ScaleWidth
iHeight = picImage.ScaleHeight
msg = Left$(txtMSG.Text, 255)

Set used_positions = New Collection

'encode the message length
EncodeByte CByte(Len(msg)), used_positions, iWidth, iHeight

For i = 1 To Len(msg) 'length of message
    EncodeByte Asc(Mid$(msg, i, 1)), used_positions, iWidth, iHeight
Next i

picImage.Picture = picImage.Image
SaveIt
ErrSub: 'error handling
If Err.Number <> 0 Then
    MsgBox (Err.Number & vbCrLf & Err.Description), vbCritical + vbOKOnly, "Error)"
Exit Sub
End If
End Sub
'------------------------------------------------------------
' Author:  Clint M. LaFever - [lafeverc@saic.com]
' Date: August,23 2002 @ 12:14:32
'------------------------------------------------------------
Private Sub SaveIt()
    On Error GoTo ErrorSaveIt
    Dim obj As CDLG, fNAME As String
    Set obj = New CDLG
    obj.VBGetSaveFileName fNAME, , , "Bitmap Files|*.bmp", , CurDir, "Save Image", "*.bmp", Me.hwnd
    If fNAME <> "" Then
        SavePicture picImage.Picture, fNAME
    End If
    Exit Sub
ErrorSaveIt:
    MsgBox Err & ":Error in call to SaveIt()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Function NumericPassword(ByVal password As String) As Long
    Dim Value As Long
    Dim ch As Long
    Dim shift1 As Long
    Dim shift2 As Long
    Dim i As Integer
    Dim str_len As Integer
    shift1 = 3
    shift2 = 17
    str_len = Len(password)
    For i = 1 To str_len
        ch = Asc(Mid$(password, i, 1))
        Value = Value Xor (ch * 2 ^ shift1)
        Value = Value Xor (ch * 2 ^ shift2)
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = Value
End Function


Private Sub EncodeByte(ByVal Value As Byte, ByVal used_positions As Collection, ByVal iWidth As Integer, ByVal iHeight As Integer)
On Error GoTo ErrSub 'error handling

Dim i As Integer
Dim byte_mask As Integer
Dim r As Integer
Dim c As Integer
Dim pixel As Integer
Dim clrr As Byte
Dim clrg As Byte
Dim clrb As Byte
Dim color_mask As Integer

byte_mask = 1
For i = 1 To 8
    'pick a random pixel and RGB
    PickPosition used_positions, iWidth, iHeight, r, c, pixel

'find out each specific pixel's colouring
UnRGB picImage.Point(r, c), clrr, clrg, clrb

If Value And byte_mask Then 'the value to be stored
    color_mask = 1 'mask colouring? yes...
Else
    color_mask = 0 'or no...
End If

Select Case pixel 'update with the new colour
Case 0
    clrr = (clrr And &HFE) Or color_mask
Case 1
    clrg = (clrg And &HFE) Or color_mask
Case 2
    clrb = (clrb And &HFE) Or color_mask
End Select

picImage.PSet (r, c), RGB(clrr, clrg, clrb) 'new colour
byte_mask = byte_mask * 2

Next i

ErrSub: 'error handling
If Err.Number <> 0 Then
    MsgBox (Err.Number & vbCrLf & Err.Description), vbCritical + vbOKOnly, "Error)"
    Exit Sub
End If
End Sub

Private Sub PickPosition(ByVal used_positions As Collection, ByVal iWidth As Integer, ByVal iHeight As Integer, ByRef r As Integer, ByRef c As Integer, ByRef pixel As Integer)
'find an unused combination (R,C, Pixel)

Dim position_code As String

On Error Resume Next 'error handling

Do  'pick a position
    r = Int(Rnd * iWidth)
    c = Int(Rnd * iHeight)
    pixel = Int(Rnd * 3)

    'find out if we can use the position or not
    position_code = "(" & r & "," & c & "," & pixel & ")"
    used_positions.Add position_code, position_code

If Err.Number = 0 Then Exit Do
    Err.Clear
    
Loop

End Sub

Private Sub UnRGB(ByVal Color As OLE_COLOR, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte)
'sub to return the colour's values

    r = Color And &HFF&
    g = (Color And &HFF00&) \ &H100&
    b = (Color And &HFF0000) \ &H10000
    
End Sub

Private Sub DecodeIt() 'decode file
On Error GoTo ErrSub 'error handling

Dim msg_length As Byte
Dim msg As String
Dim ch As Byte
Dim i As Integer
Dim used_positions As Collection
Dim iWidth As Integer
Dim iHeight As Integer
Dim show_pixels As Boolean


Rnd -1
Randomize NumericPassword(mstrPASS)
'randomize the password

iWidth = picImage.ScaleWidth
iHeight = picImage.ScaleHeight
Set used_positions = New Collection

'decode the message length
msg_length = DecodeByte(used_positions, iWidth, iHeight)

For i = 1 To msg_length 'decode the message
    ch = DecodeByte(used_positions, iWidth, iHeight) 'by using the used positions...
    msg = msg & Chr$(ch)
Next i

picImage.Picture = picImage.Image
picImage.Refresh
txtMSG.Text = msg 'set the message


txtMSG.Text = cRC4(txtMSG.Text, mstrPASS) 'decode the message using RC4




ErrSub: 'error handling
If Err.Number <> 0 Then
    MsgBox (Err.Number & vbCrLf & Err.Description), vbCritical + vbOKOnly, "Error)"
Exit Sub
End If

End Sub

Private Function DecodeByte(ByVal used_positions As Collection, ByVal iWidth As Integer, ByVal iHeight As Integer) As Byte
On Error GoTo ErrSub 'error handling

Dim Value As Integer
Dim i As Integer
Dim byte_mask As Integer
Dim r As Integer
Dim c As Integer
Dim pixel As Integer
Dim clrr As Byte
Dim clrg As Byte
Dim clrb As Byte
Dim color_mask As Integer

byte_mask = 1
For i = 1 To 8
    'pick a random pixel and RGB
    PickPosition used_positions, iWidth, iHeight, r, c, pixel

'find out each specific pixel's colouring
UnRGB picImage.Point(r, c), clrr, clrg, clrb

Select Case pixel 'update with the new colour
Case 0
    color_mask = (clrr And &H1)
Case 1
    color_mask = (clrg And &H1)
Case 2
    color_mask = (clrb And &H1)
End Select

If color_mask Then
    Value = Value Or byte_mask
End If

byte_mask = byte_mask * 2

Next i

DecodeByte = CByte(Value)

ErrSub: 'error handling
If Err.Number <> 0 Then
    MsgBox (Err.Number & vbCrLf & Err.Description), vbCritical + vbOKOnly, "Error)"
End If

End Function

