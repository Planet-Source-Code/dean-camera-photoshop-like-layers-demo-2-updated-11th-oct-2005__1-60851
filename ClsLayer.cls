VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'                   PHOTOSHOP-LIKE ALPHABLENDED LAYERS DEMO
'                             BY DEAN CAMERA, 2005
'
'                         Email: dean_camera@hotmail.com
'
' This code is presented in its entirety as free. Feel free to use and abuse it but I would
' appreciate a mention somewhere in your program. A few million dollars of your resulting profit
' wouldn't go amiss either...
'
' SUPPORTED OSes: Windows 2000, 98, ME, XP
'
' ================================================================================================
'
' DESCRIPTION OF THIS FILE:
'
' This class holds the properties of each layer. While it is true that in its current form it
' could be compressed into a indexed array of "Layer" types inside the ClsLayerHandler class,
' I opted for this method as it allows for easy expansion later.
'
' A blank (transparent) bitmap is created automatically with the initialisation of this class.
' This is also automatically painted by the clsLayerHandler class with the current background
' colour (set by the BGColour global variable) but that code may be ommitted if desired. This
' automatic bitmap creation ensures that no glitches will occur if the layer is painted before
' anything is painted to it. This bitmap can be accessed via your program and thus can be used
' for standard drawing functions - perhaps for a multilayered drawing program. In any case, this
' bitmap is DELETED and then recreated if an image is loaded via the LoadFile method.
'
' When a new layer is created, it is by default made visible (visible flag set to true). Despite
' this, the alpha value is not set automatically - thus the layer WILL be invisible - and it should
' be set upon creation by the main program.

Public BitmapHandle As Long
Public Index As Integer

Public Alpha As Integer
Public Visible As Boolean

Public OffsetX As Long
Public OffsetY As Long

Private Sub Class_Initialize()
    Dim ScreenDC As Long

    Visible = True ' Default for new layers is visible

    ScreenDC = GetDC(0) ' Get the DC of the screen
    BitmapHandle = CreateCompatibleBitmap(ScreenDC, FrmTestLayers.picImage.Width, FrmTestLayers.picImage.Height) ' Create a new bitmap compatible with the screen
    ReleaseDC 0&, ScreenDC ' Release control of the Screen's DC; prevents a GDI leak
End Sub

Private Sub Class_Terminate()
    DeleteObject BitmapHandle ' Delete the created tempoary bitmap
End Sub

Sub LoadFile(FileName As String)
    If (Dir(FileName) = "") Then Exit Sub ' File not found, exit sub

    DeleteObject BitmapHandle ' Delete the created bitmap

    BitmapHandle = LoadImage(App.hInstance, FileName, IMAGE_BITMAP, FrmTestLayers.picImage.Width, FrmTestLayers.picImage.Height, LR_LOADFROMFILE) ' Load and stretch the file into a new bitmap

    Alpha = 255 ' Set a default Alpha value (no transparency)
End Sub
