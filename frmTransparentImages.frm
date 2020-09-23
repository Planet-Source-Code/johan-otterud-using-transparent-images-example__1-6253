VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transparent Images Example by Johan Otterud (v.1.0)"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Create mask (important! click here first)"
      Height          =   495
      Left            =   1500
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Flipped Vertical"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   1035
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Flipped Horizontal"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Picture"
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   1395
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      Height          =   3915
      Left            =   60
      Picture         =   "frmTransparentImages.frx":0000
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   433
      TabIndex        =   2
      Top             =   1440
      Width           =   6555
      Begin VB.CommandButton Command3 
         Caption         =   "Use image from Windows Clipboard..."
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Press left mouse button to draw and right mouse button to flip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   555
         Left            =   0
         TabIndex        =   10
         Top             =   3600
         Width           =   6375
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3780
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   1
      Top             =   360
      Width           =   1380
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   5220
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   0
      Top             =   360
      Width           =   1380
   End
   Begin VB.PictureBox PictureOriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      Picture         =   "frmTransparentImages.frx":63D38
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   6
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label Label2 
      Caption         =   "This is what a mask looks like:"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   2715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SRCAND = &H8800C6
Const SRCCOPY = &HCC0020
Const SRCERASE = &H440328
Const SRCINVERT = &H660046
Const SRCPAINT = &HEE0086

'// Used for painting the sprite and the mask on the image
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'// Used for flipping horizontal and vertical
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'// Used when replacing the "Transparent color" with black and white
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


Private Sub CreateMasks(MaskBox1 As PictureBox, MaskBox2 As PictureBox, OriginalPicture As PictureBox, ColorToMakeTransparent As Long)
MaskBox1.Picture = OriginalPicture.Picture
MaskBox2.Picture = OriginalPicture.Picture
'// Make sure ScaleMode 3 is used!
If Not MaskBox1.ScaleMode = 3 Then MsgBox "All the Pictures must use ScaleMode: 3 - Pixels!": Exit Sub
If Not MaskBox2.ScaleMode = 3 Then MsgBox "All the Pictures must use ScaleMode: 3 - Pixels!": Exit Sub
If Not OriginalPicture.ScaleMode = 3 Then MsgBox "All the Pictures must use ScaleMode: 3 - Pixels!": Exit Sub


'// This Sub creates the mask and the sprite, it replaces the color specified
'   with white and black which VB then uses when making certain colors transparent
Dim FoundColor As Long
    
    For J = 0 To MaskBox1.Height / 15
        For I = 0 To MaskBox1.Width / 15
        FoundColor = GetPixel(MaskBox1.hdc, I, J)
            If FoundColor = ColorToMakeTransparent Then
            SetPixel MaskBox1.hdc, I, J, vbWhite
            End If
        Next I
    Next J

    For J = 0 To MaskBox2.Height / 15
        For I = 0 To MaskBox2.Width / 15
        FoundColor = GetPixel(MaskBox2.hdc, I, J)
            If FoundColor = ColorToMakeTransparent Then
            SetPixel MaskBox2.hdc, I, J, vbBlack
            End If
        Next I
    Next J

'// This code isn't really nescessary
MaskBox1.Refresh
MaskBox2.Refresh

End Sub

Private Sub PaintTransparentImage(PaintInside As PictureBox, MaskBox As PictureBox, SpriteBox As PictureBox, X As Long, Y As Long)
Dim RESULT_BACK As Integer
'// Paint the mask
RESULT_BACK = BitBlt(PaintInside.hdc, X, Y, MaskBox.ScaleWidth, MaskBox.ScaleHeight, MaskBox.hdc, 0, 0, SRCAND)
'// Paint the sprite
RESULT_BACK = BitBlt(PaintInside.hdc, X, Y, SpriteBox.ScaleWidth, SpriteBox.ScaleHeight, SpriteBox.hdc, 0, 0, SRCPAINT)
'// Important when using AUTOREDRAW=TRUE you will need
'   to Refresh the picture or you won't see anything!
PaintInside.Refresh

End Sub

Private Sub FlipPictureHorizontal(PictureToFlipHorizontal As PictureBox)
    '// Flip Horizontal
    Dim RESULT_BACK As Integer
    RESULT_BACK = StretchBlt(PictureToFlipHorizontal.hdc, PictureToFlipHorizontal.ScaleWidth, 0, -PictureToFlipHorizontal.ScaleWidth, PictureToFlipHorizontal.ScaleHeight, PictureToFlipHorizontal.hdc, 0, 0, PictureToFlipHorizontal.ScaleWidth, PictureToFlipHorizontal.ScaleHeight, SRCCOPY)
    '// Important when using AUTOREDRAW=TRUE you will need
    '   to Refresh the picture or you won't see anything!
    PictureToFlipHorizontal.Refresh
    
End Sub

Private Sub FlipPictureVertical(PictureToFlipVertical As PictureBox)
    '// Flip Vertical
    Dim RESULT_BACK As Integer
    RESULT_BACK = StretchBlt(PictureToFlipVertical.hdc, 0, PictureToFlipVertical.ScaleHeight, PictureToFlipVertical.ScaleWidth, -PictureToFlipVertical.ScaleHeight, PictureToFlipVertical.hdc, 0, 0, PictureToFlipVertical.ScaleWidth, PictureToFlipVertical.ScaleHeight, SRCCOPY)
    '// Important when using AUTOREDRAW=TRUE you will need
    '   to Refresh the picture or you won't see anything!
    PictureToFlipVertical.Refresh
    
End Sub
Private Sub Check1_Click()

   '// Flipping The Mask
   FlipPictureHorizontal Picture1
   '// Flipping The Sprite
   FlipPictureHorizontal Picture2

End Sub

Private Sub Check2_Click()

    '// Flipping The Mask
    FlipPictureVertical Picture1
    '// Flipping The Sprite
    FlipPictureVertical Picture2
    
End Sub


Private Sub Command1_Click()
'// Save Picture
Dim Answer As String
Answer = InputBox("Where do you wish to save your image?", "Save image...", "c:\image.bmp")
If Answer = "" Then Exit Sub

SavePicture Picture3.Image, Answer

End Sub


Private Sub Command2_Click()
Check1.Enabled = True: Check2.Enabled = True

'// Creates the mask with my mask sub, in this case the transparent color is
'   the color located in the left-top corner of the picture (Position (1,1))
'   (if you want to select a transparent color of your own replace
'   PictureOriginal.Point(1, 1) with for instance vbWhite

CreateMasks Picture1, Picture2, PictureOriginal, PictureOriginal.Point(1, 1)   '<-- The Transparent Color

End Sub

Private Sub Command3_Click()
'// First we check so that there really is a bitmap in the
'   clipboard
If Clipboard.GetData(2) Then
PictureOriginal.Picture = Clipboard.GetData(2)
Else
MsgBox "No Image Available In The Clipboard!", 64
End If

'// Reset the mask
Picture1.Picture = Nothing
Picture2.Picture = Nothing

End Sub


Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'// OOps you haven't clicked the Create mask button yet!
If Picture1.Picture = 0 Then
MsgBox "Press the Create Mask button first or you won't see any results! (Off course this can be done automatically in your program/project)", 48
Exit Sub
End If
'* In this sub the mask and the sprite is painted on top of each other with the
'* "PaintTransparentImage" Sub and forms the transparent image in that way,
'* the "x - (Pic..." line specifies the LEFT position of the transparent image and
'* the RIGHT position is specified by the "y - (Pic..." (what it does is that it takes
'* the mouse position and divides it with 15 and then with 2 to center it over the
'* mouse cursor, you're free to replace this with a position of your own.

    '// Occurs only when LEFT mouse-button has been clicked
    If Button = 1 Then
    '// Paints the Transparent Image in Picture3 with the "PaintTransparentImage" Sub
    PaintTransparentImage Picture3, Picture1, Picture2, X - (Picture1.Width / 15) / 2, Y - (Picture1.Height / 15) / 2
    End If
    
    '// Flip horizontal, vertical or both when clicking the RIGHT mouse-button
    If Button = 2 Then
    Static whatNow As Integer
    Check1.Value = 0
    Check2.Value = 0
        If whatNow = 0 Then
        Check1.Value = 1
        End If
        If whatNow = 1 Then
        Check2.Value = 1
        End If
        If whatNow = 2 Then
        Check1.Value = 1
        Check2.Value = 1
        End If
        If whatNow = 3 Then
        Check1.Value = 0
        Check2.Value = 0
        whatNow = -1
        End If
        whatNow = whatNow + 1
    End If
End Sub


Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '// Occurs only when left mouse-button has been clicked
    If Button = 1 Then
     Dim RESULT_BACK As Integer
    '// Paint the mask
     RESULT_BACK = BitBlt(Picture3.hdc, X - (Picture1.Width / 15) / 2, Y - (Picture1.Height / 15) / 2, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hdc, 0, 0, SRCAND)
    '// Paint the sprite
     RESULT_BACK = BitBlt(Picture3.hdc, X - (Picture1.Width / 15) / 2, Y - (Picture1.Height / 15) / 2, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, SRCPAINT)
    '// Important when using AUTOREDRAW=TRUE you will need
    '   to Refresh the picture or you won't see anything!
     Picture3.Refresh
    End If
    
End Sub


