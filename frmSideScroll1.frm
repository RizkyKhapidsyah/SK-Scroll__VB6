VERSION 5.00
Begin VB.Form frmSideScroll1 
   AutoRedraw      =   -1  'True
   Caption         =   "Side Scrolling"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerScroll 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2520
      Top             =   4080
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   735
   End
End
Attribute VB_Name = "frmSideScroll1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Side Scrolling
'
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
'****************************************

Dim BackDC As Long

'Back ground dimensions
Const BackHeight As Long = 250
Const BackLength As Long = 750

'The width of the scrolling screen
Const ScrollWidth As Long = 250


Private Sub cmdExit_Click()

DeleteGeneratedDC BackDC

Unload Me
Set frmSideScroll1 = Nothing

End Sub

Private Sub cmdStart_Click()

TimerScroll.Enabled = True

End Sub

Private Sub cmdStop_Click()

TimerScroll.Enabled = False

End Sub

Private Sub Form_Load()

'Load the background
BackDC = GenerateDC(App.Path & "\side.bmp")

'dimension the form
Me.Move Me.Left, Me.Top, 250 * Screen.TwipsPerPixelX, Me.Height

End Sub
'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function

Private Sub TimerScroll_Timer()
Static X As Long
Dim GlueWidth As Long, EndScroll As Long

If X + ScrollWidth > BackLength Then 'We ned to glue at the beginnig again
    'Calculate the remaining width
    GlueWidth = X + ScrollWidth - BackLength
    EndScroll = ScrollWidth - GlueWidth
    
    'Blit the first part
    BitBlt Me.hdc, 0, 0, EndScroll, BackHeight, BackDC, X, 0, vbSrcCopy
    'Now draw from the beginning again
    BitBlt Me.hdc, EndScroll, 0, GlueWidth, BackHeight, BackDC, 0, 0, vbSrcCopy
Else

    
    BitBlt Me.hdc, 0, 0, ScrollWidth, BackHeight, BackDC, X, 0, vbSrcCopy

End If

Me.Refresh

X = (X Mod BackLength) + 10

End Sub
