VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "RGB -> YUV -> RGB"
   ClientHeight    =   8190
   ClientLeft      =   195
   ClientTop       =   1605
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10200
   Begin VB.CommandButton cmdProcess 
      Caption         =   "go"
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
   Begin VB.PictureBox picYUV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   5040
      ScaleHeight     =   3630
      ScaleWidth      =   4830
      TabIndex        =   4
      Top             =   4200
      Width           =   4830
   End
   Begin VB.PictureBox picY 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   120
      ScaleHeight     =   3630
      ScaleWidth      =   4830
      TabIndex        =   2
      Top             =   4200
      Width           =   4830
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   120
      ScaleHeight     =   3600
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   4800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "YUV -> RGB"
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   3960
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Y (Luma)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Source image"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Code written by Marius Hudea. Permission to use this code in your projects granted
' as long as my contribution is mentioned somewhere. You should optimize the conversion routines
' if you plan to use them in something that needs speed.
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const CREATE_NEW As Long = 1
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3
Private Const OPEN_ALWAYS As Long = 4
Private Const OPEN_IF_EXISTS As Long = (&H1)
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_BEGIN As Long = 0
Private Const FILE_END As Long = 2
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const GENERIC_WRITE As Long = &H40000000
Private Const GENERIC_READ As Long = &H80000000

Private Type typeRGBQuad
 R As Byte
 G As Byte
 B As Byte
 x As Byte ' reserved
End Type

Private Type typeYUV
 Y As Byte ' Luminance
 U As Byte ' Chroma b (Cb)
 V As Byte ' Chroma r (Cr)
End Type

Private imgRGB() As typeRGBQuad
Private imgRGB2() As typeRGBQuad
Private imgYUV() As typeYUV

'Conversion routines for RGB->YUV and YUV->RGB
'
' coefficients        Rec.601 Rec.709 FCC
' Kr : Red channel    0.299   0.2125  0.3
' Kg : Green channel  0.587   0.7154  0.59
' Kb : Blue channel   0.114   0.0721  0.11
'
' 601 is the standard that is used.
'
' The formulas look scary, but really they're not. Stand back and relax.
'
'y - 16 = (Kr*219/255)*r + (Kg*219/255)*g + (Kb*219/255)*b
'u - 128 = - Kr*112/255*r/(1-Kb) - Kg*112/255*g/(1-Kb) + 112/255*b
'v - 128 = 112/255*r - Kg*112/255*g/(1-Kr) - Kb*112/255*b/(1-Kr)
'
' for more information check :
'
' http://www.avisynth.org/mediawiki/wiki/Color_conversions
' http://www.intersil.com/data/an/an9717.pdf

'
'This function is used to make sure computed values fall within the standard ranges.
' Y = 16..235 ; U,V = 16..240; R,G,B = 0..255
'
Private Function CheckRange(min As Byte, max As Byte, val As Long) As Byte
If val < min Then
  CheckRange = min: Exit Function
End If
If val > max Then
  CheckRange = max: Exit Function
End If
CheckRange = CByte(val)
End Function

Private Sub cmdProcess_Click()
Dim i As Long
Dim j As Long
Dim imgw As Long
Dim imgh As Long
Dim pixel As Long
Dim R As Byte
Dim G As Byte
Dim B As Byte
Dim Y As Byte
Dim U As Byte
Dim V As Byte

Dim handle As Long
Dim bwrote As Long

Dim x(1 To 8, 0 To 255) As Double

'Determine the width and height of the source image
imgw = picSource.Width \ Screen.TwipsPerPixelX - 1
imgh = picSource.Height \ Screen.TwipsPerPixelY - 1

ReDim imgRGB(0 To imgw, 0 To imgh)                      ' original RGB image
ReDim imgYUV(0 To imgw, 0 To imgh)                      ' original in YUV format
ReDim imgRGB2(0 To imgw, 0 To imgh)                     ' YUV converted back to RGB

'read the file to memory for faster operation
For j = 0 To imgh
 For i = 0 To imgw
  pixel = GetPixel(picSource.hdc, i, j)                 ' gets the color of a pixel
  CopyMemory imgRGB(i, j), pixel, 4                     ' copies the 4 bytes at once into the typeRGBQuad
 Next i
Next j

'convert to YUV (YCbCr)
' Formulas pasted again, to be easier to understand what happens
' Kr = 0.299 Kg = 0.587 Kb = 0.114
'
'y - 16 = (Kr*219/255)*r + (Kg*219/255)*g + (Kb*219/255)*b
'u - 128 = - Kr*112/255*r/(1-Kb) - Kg*112/255*g/(1-Kb) + 112/255*b
'v - 128 = 112/255*r - Kg*112/255*g/(1-Kr) - Kb*112/255*b/(1-Kr)
'
' As you see, there are a lot of fractions and multiplications.
' If we have large images or a sequence of images, it would be wise to
' compute everything from the start and just create some small tables
' with the results. Our ecuations will become :

'Y = Y = 0.257R´ + 0.504G´ + 0.098B´ + 16
'U = Cb = -0.148R´ - 0.291G´ + 0.439B´ + 128
'V = Cr = 0.439R´ - 0.368G´ - 0.071B´ + 128

For i = 0 To 255
  x(1, i) = 0.257 * i   ' 0.299 * 219/255 = 0.25678 = 0.257
  x(2, i) = 0.504 * i   ' 0.587 * 219/255 = 0.5041 = 0.504
  x(3, i) = 0.098 * i   ' you should get the ideea by now
  x(4, i) = 0.148 * i
  x(5, i) = 0.291 * i
  x(6, i) = 0.439 * i
  x(7, i) = 0.368 * i
  x(8, i) = 0.071 * i
Next i

' Now, we can actually use the tables created above to convert the image to yuv
For j = 0 To imgh
 For i = 0 To imgw
  R = imgRGB(i, j).R    ' this is extra but this way it's easier to read the formulas later on
  G = imgRGB(i, j).G
  B = imgRGB(i, j).B
  imgYUV(i, j).Y = CheckRange(16, 235, Round(x(1, R) + x(2, G) + x(3, B), 0) + 16)
  imgYUV(i, j).U = CheckRange(16, 240, Round(-x(4, R) - x(5, G) + x(6, B), 0) + 128)
  imgYUV(i, j).V = CheckRange(16, 240, Round(x(6, R) - x(7, G) - x(8, B), 0) + 128)
 Next i
Next j
'
' The image is now converted in YUV and stored in the array.
' The backwards process uses the following formulas to convert from YUV to RGB
'
'R = 1.164(Y - 16) + 1.596(Cr - 128)
'G = 1.164(Y - 16) - 0.813(Cr - 128) - 0.392(Cb - 128)
'B = 1.164(Y - 16) + 2.017(Cb - 128)
'
'Like before, we're going to build tables that will ease the process.
'You will maybe notice that I start from 0 instead of 16 (Y,U,V start from 16), this is just to be on
'the safe side. We're not trying to make the fastest and most optimized code here.
'
For i = 0 To 255
 x(1, i) = 1.164 * (i - 16)
 x(2, i) = 1.596 * (i - 128)
 x(3, i) = 0.813 * (i - 128)
 x(4, i) = 0.392 * (i - 128)
 x(5, i) = 2.017 * (i - 128)
Next i
'
' Now, we can actually use the tables created above to convert the image back to RGB
'
For j = 0 To imgh
 For i = 0 To imgw
  Y = imgYUV(i, j).Y
  U = imgYUV(i, j).U
  V = imgYUV(i, j).V
  imgRGB2(i, j).R = CheckRange(0, 255, Round(x(1, Y) + x(2, V), 0))
  imgRGB2(i, j).G = CheckRange(0, 255, Round(x(1, Y) - x(3, V) - x(4, U), 0))
  imgRGB2(i, j).B = CheckRange(0, 255, Round(x(1, Y) + x(5, U), 0))
 Next i
Next j
'
' Now,let's do some nice drawings to see what we have done.
' A nice grayscale image can be obtained by using the luma component in the YUV format
' as red,green and blue
'
For j = 0 To imgh
 For i = 0 To imgw
 Y = imgYUV(i, j).Y
  SetPixel picY.hdc, i, j, RGB(Y, Y, Y)
 Next i
Next j
picY.Refresh

'
' In the second picture box, we'll draw the RGB image obtained from YUV.
' If you use Print Screen and zoom at pixel level , you'll see that some pixels in
' the RGB image obtained from YUV are not *exactly* the same as the ones in the RGB image.
' This is normal, it's a very very small sacrifice, almost unnoticeable.
' All video conversion as a first step in converting the images to
' YUY2 (HuffYUV) or YV12 (DVD, MPG, XVID)
'

For j = 0 To imgh
 For i = 0 To imgw
  SetPixel picYUV.hdc, i, j, RGB(imgRGB2(i, j).R, imgRGB2(i, j).G, imgRGB2(i, j).B)
 Next i
Next j
picYUV.Refresh

'We're going to create a file on the disk with the YUV conversion results.
' Open the file
handle = CreateFile("c:\testyuv.bin", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, CREATE_ALWAYS, 0&, 0&)
For j = 0 To imgh
 For i = 0 To imgw
  WriteFile handle, imgYUV(i, j).Y, 1, bwrote, 0&
  WriteFile handle, imgYUV(i, j).U, 1, bwrote, 0&
  WriteFile handle, imgYUV(i, j).V, 1, bwrote, 0&
 Next i
Next j

CloseHandle handle  'file can now be closed.

MsgBox "Done"
End Sub

Private Sub Form_Load()
On Error Resume Next
picSource.Picture = LoadPicture(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "picture.bmp")
End Sub
