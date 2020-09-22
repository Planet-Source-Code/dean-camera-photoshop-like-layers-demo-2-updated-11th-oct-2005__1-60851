VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLayerHandler"
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
' NOTE: The sizes of the bitmaps created are set by the width and height of the pictureboxes on the
'       demonstration form (the BitBlt and other operation sizes are adjusted to the bitmap size
'       progmatically). If you are using this code in another project, you will need to change these
'       picturebox references to global variables or parameters of your choosing.
'
' Tested and is 100% GDI-leak free.
'
' ================================================================================================
'
' DESCRIPTION OF THIS FILE:
'
' This class manages the creation, deletion and drawing of the graphical layers. Several events
' have also been added for convience.
'
' All layer classes are stored in a public collection named Layers. The properties of any layer
' cannot be accessed directly through this layer; the layer data must first be copied by the user
' into a Layer class structure (see the code of the demonstration form).
'
' The base background colour can be set via the public variable BGColour. Changing this property
' does NOT redraw the layers (you must do this manually) as I have made no presumtions as to whether
' you will be continuously drawing to the same destination DC (hence the reason why you must supply
' the destination DC each time.
'
' The CreateLayer function returns a new layer (with default properties set) and returns this as a
' clsLayer class. You must then create your own tempoary layer class and set this to the CreateLayer
' function return to be able to set the new layer's properties.
'
' Two drawing functions have been added; DrawLayers and DrawSingleLayer. As the name suggests, the
' first sub will render all the layers in the collection (with alphablending) onto the destination
' DC as supplied in the sub parameters. For an outline of the drawing process, see the comments on
' the demonstration form. The DrawSingleLayer sub will (astoundingly!) draw a single layer - the
' layer is chosen by its index - onto the specified destination DC.
'
' You can copy the contents of one bitmap to another via the CopyBitmapToBitmap sub. This is useful
' for copying one layer to another (or duplicating a layer, like in the demonstration). Note that
' this sub uses bitmaps instead of DCs; the function will select the bitmaps automatically.
'
' A CopyDCtoBitmap subroutine is also provided and is used in the demonstration for flattening the
' image, as is the InvertLayer sub.
'
' For layer manipulations, you can use the RenderLayerToTempDC function which will return the handle
' to a DC that can be used for layer image manipulations. You will need to use the CopyDCToBitmap to
' save the results after the manipulation (in the demo simple circles are drawn using this function on
' mouse click). NB: The contents of the temp DC are clobbered (overwritten) on layer render, thus any
' manipulations must be saved back into the layer before redering takes place.

Private TempDC As Long
Private TempBitmap As Long
Private TempLayerDC As Long

Public Layers As Collection
Public BGColour As Long

Public Enum LayerRotateType
    FlipHorizontal = 1
    FlipVertical = 2
End Enum

#If False Then ' These lines stop the IDE from messing up case of the Enum members above
    Private FlipHorizontal, FlipVertical
#End If

Public Event StartRender() '  Fires at the start of the DrawLayers Routine
Public Event FinishRender() ' Fires at the end of the DrawLayers Routine
Public Event RenderingLayer(Layer As Integer, SingleLayerRender As Boolean) '   Firers before a layer is drawn
Public Event RenderedLayer(LayerNum As Integer, SingleLayerRender As Boolean) ' Firers after a layer is drawn

Private Sub Class_Initialize()
    Dim ScreenDC As Long

    ScreenDC = GetDC(0) ' Get the DC for the screen

    TempLayerDC = CreateCompatibleDC(ScreenDC) ' Create a tempoary DC for the layer's bitmaps
    TempDC = CreateCompatibleDC(ScreenDC) ' Create a tempoary DC for the rendered layers

    TempBitmap = CreateCompatibleBitmap(ScreenDC, FrmTestLayers.picImage.Width, FrmTestLayers.picImage.Height) ' Creat a bitmap for the rendered layers DC
    SelectObject TempDC, TempBitmap ' Select the created Bitmap into the tempoary DC's context

    ReleaseDC 0&, ScreenDC ' Release the control of the screen's DC; prevents a GDI leak

    BGColour = vbWhite ' Set the default background colour to white

    Set Layers = New Collection ' Create a new layers collection
End Sub

Private Sub Class_Terminate()
    DeleteDC TempDC ' Delete the rendered layers DC
    DeleteDC TempLayerDC ' Delete the layer DC

    DeleteObject TempBitmap ' Delete the rendered layers bitmap

    Set Layers = Nothing ' Delete the layers collection
End Sub

Public Function CreateLayer() As ClsLayer
    Dim cLayer As ClsLayer
    Dim hBrush As Long, rRect As RECT
    Dim BitmapInfo As BITMAP

    Set cLayer = New ClsLayer ' Set the tempoary layer to a new layer
    Layers.Add cLayer ' Add the new layer to the layers collection
    cLayer.Index = Layers.Count ' Set the index for the new layer

    SelectObject TempLayerDC, cLayer.BitmapHandle ' Load the created layer's bitmap into a tempoary DC
    GetObject cLayer.BitmapHandle, Len(BitmapInfo), BitmapInfo ' Get the size information of the created bitmap

    hBrush = CreateSolidBrush(BGColour) ' Create a new solid brush with the user-selected background colour
    rRect.Bottom = BitmapInfo.bmHeight ' \ Set the rectangle values to the
    rRect.Right = BitmapInfo.bmWidth '   / dimentions of the temp bitmap
    FillRect TempLayerDC, rRect, hBrush ' Fill the created layer's bitmap with solid white brush
    DeleteObject hBrush ' Delete the solid background brush

    Set CreateLayer = cLayer ' Set the function return value to the created layer object
End Function

Public Sub DeleteLayer(Index As Long)
    If Index <= 0 Or Index > Layers.Count Then Exit Sub ' If index invalid, exit sub

    Layers.Remove Index ' Delete the collection entry
End Sub

Public Sub DrawSingleLayer(DestDC As Long, Index As Integer, Width As Long, Height As Long, Optional NoOffset As Boolean)
    Dim cLayer As ClsLayer
    Dim BitmapInfo As BITMAP

    If DestDC = 0 Then Exit Sub ' If DestDC invalid, exit sub

    RaiseEvent RenderingLayer(Index, True) ' Raise the RenderingLayer event

    GetObject TempBitmap, Len(BitmapInfo), BitmapInfo ' Get the size information of the temp bitmap
    SelectObject TempDC, TempBitmap ' Make sure the TempDC's bitmap is in the TempDC's context

    Set cLayer = Layers(Index) ' Set the tempoary layer to the passed layer index
    SelectObject TempLayerDC, cLayer.BitmapHandle ' Select the layer's bitmap into the tempoary layer DC

    StretchBlt DestDC, IIf(NoOffset, 0, cLayer.OffsetX), IIf(NoOffset, 0, cLayer.OffsetY), Width, Height, TempLayerDC, 0, 0, BitmapInfo.bmWidth, BitmapInfo.bmHeight, vbSrcCopy ' Stretch the layer bitmap onto the destination DC

    RaiseEvent RenderedLayer(Index, True) ' Raise the RenderedLayer event

    Set cLayer = Nothing
End Sub

Public Sub DrawLayers(DestDC As Long)
    Dim CurrLayer As ClsLayer
    Dim BF As BLENDFUNCTION, lBF As Long
    Dim hBrush As Long, rRect As RECT
    Dim BitmapInfo As BITMAP

    If DestDC = 0 Then Exit Sub ' If DestDC invalid, exit sub

    RaiseEvent StartRender

    GetObject TempBitmap, Len(BitmapInfo), BitmapInfo ' Get the size information of the temp bitmap
    SelectObject TempDC, TempBitmap ' Make sure the TempDC's bitmap is in the TempDC's context

    hBrush = CreateSolidBrush(BGColour) ' Create a new solid brush with the user-selected background colour
    rRect.Bottom = BitmapInfo.bmHeight ' \ Set the rectangle values to the
    rRect.Right = BitmapInfo.bmWidth '   / dimentions of the temp bitmap
    FillRect TempDC, rRect, hBrush ' Fill the DestDC with solid white brush
    DeleteObject hBrush ' Delete the solid background brush

    For Each CurrLayer In Layers
        RaiseEvent RenderingLayer(CurrLayer.Index, False)

        If CurrLayer.Visible Then
            SelectObject TempLayerDC, CurrLayer.BitmapHandle ' Select the current layer's bitmap into the tempoary layer DC

            With BF
                .BlendOp = AC_SRC_OVER
                .BlendFlags = 0
                .SourceConstantAlpha = CurrLayer.Alpha ' Set the blending Alpha value to the Alpha value of the current layer
                .AlphaFormat = 0
            End With
            RtlMoveMemory lBF, BF, 4 ' Copy the blending structure above to a long

            AlphaBlend TempDC, CurrLayer.OffsetX, CurrLayer.OffsetY, BitmapInfo.bmWidth, BitmapInfo.bmHeight, TempLayerDC, 0, 0, BitmapInfo.bmWidth, BitmapInfo.bmHeight, lBF ' Blend the layer onto the temp DC
        End If

        RaiseEvent RenderedLayer(CurrLayer.Index, False)
    Next

    BitBlt DestDC, 0, 0, BitmapInfo.bmWidth, BitmapInfo.bmHeight, TempDC, 0, 0, vbSrcCopy  ' Copy the rendered layer DC onto the Dest DC

    RaiseEvent FinishRender
End Sub

Public Sub CopyBitmapToBitmap(BitmapSource As Long, BitmapDest As Long)
    Dim BitmapInfo As BITMAP

    If BitmapSource = 0 Or BitmapDest = 0 Then Exit Sub ' If either bitmap is invalid, exit the sub

    SelectObject TempLayerDC, BitmapSource ' Select the source bitmap into a tempoary DC
    SelectObject TempDC, BitmapDest ' Select the destination bitmap into a tempoary DC

    GetObject BitmapSource, Len(BitmapInfo), BitmapInfo ' Get the size information of the source bitmap

    BitBlt TempDC, 0, 0, BitmapInfo.bmWidth, BitmapInfo.bmHeight, TempLayerDC, 0, 0, vbSrcCopy ' Copy the bitmaps over

    SelectObject TempDC, TempBitmap ' Restore the TempDC's bitmap (TempLayerDC does not need to be restored)
End Sub

Public Sub CopyDCToBitmap(DCSource As Long, BitmapDest As Long)
    Dim BitmapInfo As BITMAP

    If DCSource = 0 Or BitmapDest = 0 Then Exit Sub ' If either bitmap is invalid, exit the sub

    SelectObject TempLayerDC, BitmapDest ' Select the destination bitmap into a tempoary DC

    GetObject BitmapDest, Len(BitmapInfo), BitmapInfo ' Get the size information of the destination bitmap

    BitBlt TempLayerDC, 0, 0, BitmapInfo.bmWidth, BitmapInfo.bmHeight, DCSource, 0, 0, vbSrcCopy ' Copy the bitmaps over

    SelectObject TempDC, TempBitmap ' Restore the TempDC's bitmap (TempLayerDC does not need to be restored)
End Sub

Public Function RenderLayerToTempDC(FromIndex As Long) As Long
    Dim cLayer As ClsLayer

    Set cLayer = Layers(FromIndex) ' Set the tempoary layer object to the passed layer index
    SelectObject TempLayerDC, cLayer.BitmapHandle ' Make sure the TempDC's bitmap is in the TempDC's context
    Set cLayer = Nothing
    
    RenderLayerToTempDC = TempLayerDC
End Function

Public Sub FlipLayer(RotationType As LayerRotateType, Index As Long)
    Dim cLayer As ClsLayer
    Dim BitmapInfo As BITMAP

    If Index <= 0 Or Index > Layers.Count Then Exit Sub ' If index invalid, exit sub

    Set cLayer = Layers(Index) ' Set the layer reference to the layer to be manipulated
    SelectObject TempLayerDC, cLayer.Index ' Select the layer's bitmap into the context of TempLayerDC
    GetObject cLayer.BitmapHandle, Len(BitmapInfo), BitmapInfo ' Get the size information of the layer bitmap

    Select Case RotationType ' Determine the type of flip operation
        Case FlipHorizontal
            StretchBlt TempLayerDC, BitmapInfo.bmWidth, 0, -BitmapInfo.bmWidth, BitmapInfo.bmHeight, TempLayerDC, 0, 0, BitmapInfo.bmWidth, BitmapInfo.bmHeight, vbSrcCopy
        Case FlipVertical
            StretchBlt TempLayerDC, 0, BitmapInfo.bmHeight, BitmapInfo.bmWidth, -BitmapInfo.bmHeight, TempLayerDC, 0, 0, BitmapInfo.bmWidth, BitmapInfo.bmHeight, vbSrcCopy
    End Select
End Sub

Public Sub InvertLayer(Index As Long)
    Dim rRect As RECT
    Dim BitmapInfo As BITMAP
    Dim cLayer As ClsLayer

    Set cLayer = Layers(Index) ' Set the tempoary layer object to the passed layer index
    
    SelectObject TempLayerDC, cLayer.BitmapHandle ' Select the layer bitmap into a tempoary DC
    GetObject cLayer.BitmapHandle, Len(BitmapInfo), BitmapInfo ' Get the size information of the layer bitmap

    rRect.Bottom = BitmapInfo.bmHeight ' \ Set the rectangle values to the
    rRect.Right = BitmapInfo.bmWidth '   / dimentions of the temp bitmap
    
    InvertRect TempLayerDC, rRect ' Invert the layer bitmap

    Set cLayer = Nothing
End Sub

Public Sub MoveLayer(FromIndex As Long, ToIndex As Long)
    Dim cLayer As ClsLayer
    Dim MoveAmount As Long
    Dim NewIndex As Long

    MoveAmount = ToIndex - FromIndex ' Calculate the difference between indexes
    
    If FromIndex < 1 Or FromIndex > Layers.Count Then Exit Sub ' FromIndex invalid, exit sub
    If ToIndex < 1 Or ToIndex > Layers.Count Then Exit Sub ' ToIndex invalid, exit sub
    If MoveAmount = 0 Then Exit Sub ' Exit sub if both indexes are the same
    
    If MoveAmount < 0 Then
        Layers.Add Layers(FromIndex), , ToIndex ' Re-add the selected layer to the destination index
        Layers.Remove FromIndex - MoveAmount ' Delete the original copy of the selected layer from the collection
    Else
        Layers.Add Layers(FromIndex), , , ToIndex ' Re-add the selected layer to the destination index
        Layers.Remove FromIndex  ' Delete the original copy of the selected layer from the collection
    End If
    
    For NewIndex = 1 To Layers.Count ' Look at each of the layers in the collection
        Set cLayer = Layers.Item(NewIndex) ' Get the layer
                
        cLayer.Index = NewIndex ' Make sure the collection index and the layer's indexes coincide
    Next
    
    Set cLayer = Nothing
End Sub
