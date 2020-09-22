VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTestLayers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Layer Test  Program (V2) - By Dean Camera"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9615
   Icon            =   "FrmTestLayers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInvertLayer 
      Caption         =   "Invert Layer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   23
      Top             =   1320
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CDialogs 
      Left            =   8160
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDuplicateLayer 
      Caption         =   "Duplicate Layer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdFlattenImage 
      Caption         =   "Flatten Image"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdChooseBG 
      Caption         =   "Set Background Colour"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cmdDeleteLayer 
      Caption         =   "Delete Layer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   600
      Width           =   2175
   End
   Begin VB.Frame frmLayer 
      Caption         =   "Layer"
      Height          =   3975
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdMove 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmTestLayers.frx":1042
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmTestLayers.frx":1384
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton cmdFlip 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmTestLayers.frx":16C6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmdFlip 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   2040
         Picture         =   "FrmTestLayers.frx":1A08
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
         Width           =   375
      End
      Begin VB.PictureBox picLayerPreview 
         Height          =   1396
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   1755
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox cmbLayerAlpha 
         Enabled         =   0   'False
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "FrmTestLayers.frx":1D4A
         Left            =   1320
         List            =   "FrmTestLayers.frx":1D4C
         TabIndex        =   10
         Text            =   "cmbLayerAlpha"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmbSelLayer 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   1095
      End
      Begin VB.CheckBox chkVisible 
         Alignment       =   1  'Right Justify
         Caption         =   "Visible:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2235
      End
      Begin VB.Label lblLayerIndex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label lblSelLayerIndex 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Layer Index: "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblSelectedLayer 
         Caption         =   "Selected Layer:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblLayerAlpha 
         Caption         =   "Layer Alpha:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1140
         Width           =   975
      End
   End
   Begin MSComctlLib.ProgressBar pgbRenderProgress 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4275
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdAddLayer 
      Caption         =   "Add Layer"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox picImage 
      Height          =   3426
      Left            =   120
      ScaleHeight     =   3360
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LEFT CLICK: Draw Circle on current layer RIGHT CLICK: Move selected layer"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   22
      Top             =   30
      Width           =   2895
   End
   Begin VB.Label lblRatioInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmTestLayers.frx":1D4E
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Shape shpBGColour 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Left            =   7320
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblLayersCap 
      Alignment       =   2  'Center
      Caption         =   "Layers: 0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
End
Attribute VB_Name = "FrmTestLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' NB: The demo code on this form is NOT optimised for best performance.
'
' ================================================================================================
'
' DESCRIPTION OF THIS FILE:
'
' This project will emulate the alpha-blended (transparent) layers, as used in the popular image editing program
' "Photoshop". For simplicity, loaded images are strected to the height and width of the picturebox on this form.
' Also for simplicity, some of the programming has been expanded to it's functionally equivelent but easier to
' understand code. This means that in some places, certain variables and such COULD have been optimised out but
' have been left in to make the project easier for VB/GDI novices.
'
' When you add a layer, a new ClsLayer blank Bitmap is created. This allows you to immediatly manipulate the layer's
' bitmap, although in this case the OpenFile routine is called to load a pre-created bitmap picture. Once called,
' the OpenFile sub will re-create the layer's bitmap and load the selected image. When a layer image is loaded,
' this demo calls the DrawLayers subroutine of the ClsLayerHandler object to refresh the screen. When this function
' is called, a tempoary DC is cleared (filled with a white brush) and each layer's bitmap is loaded into a created
' LayerDC, where it is alphablended with the main tempoary DC. Once all layers have been rendered, this tempoary DC
' is BitBlted - copied - into yet another DC (called "FinalDC" on this form).
'
' This FinalDC is itself BitBlted onto the picturebox when the picturebox's Paint method is called. The reason
' why two tempoary DC's are used is simple; it reduces flicker when a layer is moved. At the moment, all the layers
' are rendered before they are put into the FinalDC for BitBlting to the screen, thus a redraw only occurs after all
' layers are rendered. If each layer is alphablended onto the FinalDC directly, a Paint command will occur each time
' the user moves their mouse, whether or not the layers have finished rendering (causing flicker).
'
' I think this implementation is clever because only a Bitmap is created for each layer (as opposed to both a Bitmap
' and a DC), saving memory. The current layer that is being rendered has its Bitmap placed into the context of a
' single tempoary DC created in the LayerHandler class. On the downside, the flickerless system i'm using means that
' the speed of the program (when moving a layer) drops with each layer added. With more effort a system could be
' imposed so that all the layers benieth the selected layer (and possibly above the selected layer) are rendered onto
' yet another tempoary DC, so that only a single AlphaBlend call and two BitBlt's are needed each time a layer is
' moved. This is definetly possible but would increase the complexity of this demonstartion somewhat. If enough
' requests are made I shall add this function and resubmit.
'
' Because the equivelent GDIAlphaBlend (which uses the GDI.dll library) function is used instead of the normal
' AlphaBlend function (which uses the msimg32.dll library), only two DLLs are used in this program, User32.dll and
' GDI.dll, both of which are part of the standard Windows OS.
'
' Let me know what you think by emailing me (address at top of headder).
'
'  %%%%%%%%%%%                           ------------                           ------
' % Layer(s)  % ==   SelectObject   ==> | Layer Temp | ==    AlphaBlend    ==> | Temp |
' % Bitmap(s) % == (For each layer) ==> |     DC     | == (For each layer) ==> |  DC  |
'  %%%%%%%%%%%                           ------------                           ------
'                                                                                 ||
'  ------------                               -------                             ||
' | Picturebox | <==        BitBlt        == | Final | <==         BitBlt       =='|
' |    hDC     | <== (Upon Paint Request) == |  DC   | <== (When rendering done) =='
'  ------------                               -------
'
'
' UPDATE: Fixed many GDI memory leaks (due to not releasing GetDC(0)), added some new features

Option Explicit

Public WithEvents LayerHandler As ClsLayerHandler
Attribute LayerHandler.VB_VarHelpID = -1

Private FinalDC As Long
Private FinalBitmap As Long

Private PrevOffsetX As Long
Private PrevOffsetY As Long

Private Sub cmbLayerAlpha_Click()
    ShowLayerChanges ' Drop-down value clicked, show the changes
End Sub

Private Sub cmdChooseBG_Click()
    CDialogs.ShowColor ' Show the choose colour dialog
    shpBGColour.BackColor = CDialogs.Color ' Show the background colour preview
    LayerHandler.BGColour = CDialogs.Color ' Set the background layer colour

    LayerHandler.DrawLayers FinalDC ' Re-render the layers with the new background colour
    picImage_Paint ' Refresh the picturebox
End Sub

Private Sub cmdDeleteLayer_Click()
    Dim LayerIndex As Long
    Dim ButtonIndex As Long

    LayerHandler.DeleteLayer Int(cmbSelLayer.Text) 'Delete the selcted layer

    cmbSelLayer.Clear ' Remove all the layer numbers from the select layer combobox
    lblLayerIndex.Caption = "N/A" ' Change the selected layer index label to "N/A"

    For LayerIndex = 1 To LayerHandler.Layers.Count
        cmbSelLayer.AddItem LayerIndex ' Add each remaining layer number back the the select layer combobox
    Next

    ClearLayerPreview ' Disable the layer preview items and clear the layer preview picturebox

    If LayerHandler.Layers.Count = 0 Then ' No layers left
        cmbSelLayer.Enabled = False ' Disable the select layer combobox
        cmbLayerAlpha.Enabled = False ' Disable the select layer alpha combobox
        cmdFlattenImage.Enabled = False ' Disable the flatten image button
    End If

    cmdDuplicateLayer.Enabled = False ' Diasable the duplicate layer button
    cmdInvertLayer.Enabled = False ' Disable the invert layer button

    lblLayersCap.Caption = "Layers: " & LayerHandler.Layers.Count ' Show the total layers

    LayerHandler.DrawLayers FinalDC ' Render all the remaining layers onto the tempoary DC "FinalDC"
    picImage_Paint ' Force a refresh of the picturebox
End Sub

Private Sub cmdDuplicateLayer_Click()
    Dim CurrentLayer As ClsLayer
    Dim NewLayer As ClsLayer

    If cmbSelLayer.Text = vbNullString Then Exit Sub ' Exit sub if no layer selected

    Set CurrentLayer = LayerHandler.Layers(Int(cmbSelLayer.Text)) ' Get the currently selected layer
    Set NewLayer = LayerHandler.CreateLayer ' Create a new layer

    LayerHandler.CopyBitmapToBitmap CurrentLayer.BitmapHandle, NewLayer.BitmapHandle ' Copy the current layer to the created new one

    NewLayer.Alpha = CurrentLayer.Alpha ' Set the new layer's alpha to that of the old one

    cmbSelLayer.AddItem LayerHandler.Layers.Count  ' Add the new layer to the select layer combobox

    LayerHandler.DrawLayers FinalDC ' Render all the layers onto the tempoary "FinalDC"
    picImage_Paint ' Force a refresh of the picturebox - FinalDC is BitBlted onto the picturebox via the Paint method

    lblLayersCap.Caption = "Layers: 1" ' Change the caption of the label to show that there is now only one layer

    Set NewLayer = Nothing
    Set CurrentLayer = Nothing
End Sub

Private Sub cmdFlattenImage_Click()
    Dim cLayer As ClsLayer

    LayerHandler.DrawLayers FinalDC ' Render all the layers onto the tempoary "FinalDC"

    Set LayerHandler.Layers = New Collection ' Destroy all the current layers

    Set cLayer = LayerHandler.CreateLayer ' Create a new layer
    LayerHandler.CopyDCToBitmap FinalDC, cLayer.BitmapHandle ' Copy the rendered layers to the new layer (note this uses CopyDCtoBitmap instead of CopyBitmapToBitmap: the FinalBitmap
    '                                                          bitmap can only be in the context of one DC at a time thus the latter function won't work)

    cLayer.Alpha = 255 ' Set the new layer's alpha value to maximum (no transparency)

    cmbSelLayer.Clear ' Remove all the layer numbers from the select layer combobox
    cmbSelLayer.AddItem "1" ' Add the (now) single remining layer, the merged one we just created

    lblLayersCap.Caption = "Layers: 1" ' Change the caption of the label to show that there is now only one layer

    ClearLayerPreview ' Disable the layer preview items and clear the layer preview picturebox
    cmdInvertLayer.Enabled = False ' Disable the Invert Layer button
    cmdDuplicateLayer.Enabled = False ' Disable the Duplicate Layer button

    Set cLayer = Nothing ' Delete the reference to the created layer
End Sub

Private Sub cmdFlip_Click(Index As Integer)
    If cmbSelLayer.Text = vbNullString Then Exit Sub ' Exit sub if no layer selected

    If Index = 0 Then ' Flip Horizontal button
        LayerHandler.FlipLayer FlipHorizontal, Int(cmbSelLayer.Text) ' Flip horizontally the selected layer
    Else ' Flip Vertical button
        LayerHandler.FlipLayer FlipVertical, Int(cmbSelLayer.Text) ' Flip vertically the selected layer
    End If

    LayerHandler.DrawLayers FinalDC ' Redraw all the layers
    picImage_Paint ' Refresh the layer picturebox
    picLayerPreview_Paint ' Refresh the preview picturebox
End Sub

Private Sub cmdInvertLayer_Click()
    LayerHandler.InvertLayer Int(cmbSelLayer.Text)

    ShowLayerChanges ' Refresh the layers and the layer frame details
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then ' Move layer up button pressed
        LayerHandler.MoveLayer Int(cmbSelLayer.Text), Int(cmbSelLayer.Text) + 1
    Else ' Move layer down button pressed
        LayerHandler.MoveLayer Int(cmbSelLayer.Text), Int(cmbSelLayer.Text) - 1
    End If

    ShowLayerChanges ' Refresh the layers and the layer frame details
End Sub

Private Sub Form_Load()
    Dim ScreenDC As Long

    Me.ScaleMode = vbPixels

    cmbLayerAlpha.AddItem "255 (100%)" ' \
    cmbLayerAlpha.AddItem "191 (75%)" '   | Add the percentage sample
    cmbLayerAlpha.AddItem "128 (50%)" '  | values to the alpha combobox
    cmbLayerAlpha.AddItem "64   (25%)" '  /

    Set LayerHandler = New ClsLayerHandler ' Create a new LayerHandler

    ScreenDC = GetDC(0) ' Get the DC of the Screen

    FinalDC = CreateCompatibleDC(ScreenDC) ' Create a new DC, compatible with the screen
    FinalBitmap = CreateCompatibleBitmap(ScreenDC, picImage.Width, picImage.Height) ' Create a bitmap for the created DC, compatible with the screen
    SelectObject FinalDC, FinalBitmap ' Select the created Bitmap into the created DC's context

    ReleaseDC 0&, ScreenDC ' Release the DC of the Screen; prevents a GDI leak

    SetStretchBltMode picLayerPreview.hdc, 3 ' This set the strech copy mode to ColorOnColor ensuring the preview picturebox looks like the original

    LayerHandler.DrawLayers FinalDC ' Render the white background (no layers are present yet) to the temp DC
    picImage_Paint ' Force a picturebox paint to show the white background
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set LayerHandler = Nothing

    DeleteDC FinalDC ' Delete the tempoary DC
    DeleteObject FinalBitmap ' Delete the tempoary Bitmap

    Set FrmTestLayers = Nothing ' Completly destroy this form
End Sub

Private Sub cmbSelLayer_Click()
    Dim cLayer As ClsLayer
    Dim ButtonIndex As Long

    Set cLayer = LayerHandler.Layers(Int(cmbSelLayer.Text)) ' Set the tempoary layer object to the selected layer

    picLayerPreview_Paint ' Force the preview picturebox to repaint
    cmbLayerAlpha.Text = cLayer.Alpha ' Set the alpha textbox to the selected layer's alpha amount
    chkVisible.Value = IIf(cLayer.Visible, 1, 0) ' Set the Visible checkbox value to that of the layer
    lblLayerIndex.Caption = cLayer.Index ' Set the layer index label caption to the index of the currently selected layer

    cmdInvertLayer.Enabled = True ' Enable the invert layer button
    cmdDeleteLayer.Enabled = True ' Enable the delete layer button
    cmbLayerAlpha.Enabled = True '  Enable the change layer alpha combobox
    cmdDuplicateLayer.Enabled = True ' Enable the duplicate layer button
    chkVisible.Enabled = True ' Enable the visible checkbox

    cmdFlip(0).Enabled = True ' Enable the Flip Horizontal button
    cmdFlip(1).Enabled = True ' Enable the Flip Vertical button
    cmdMove(0).Enabled = True ' Enable then Move Layer up button
    cmdMove(1).Enabled = True ' Enable then Move Layer down button

    Set cLayer = Nothing
End Sub

Private Sub cmdAddLayer_Click()
    Dim cLayer As ClsLayer

    CDialogs.Filter = "Bitmap Files (*.bmp)|*.bmp"
    CDialogs.ShowOpen ' Show the Open file dialogue

    DoEvents ' Redraw the "Add Layer" button to prevent visual glitch

    If CDialogs.FileName <> vbNullString Then
        Set cLayer = LayerHandler.CreateLayer ' Create a new layer
        cLayer.LoadFile CDialogs.FileName ' Load the chosen bitmap into the new layer's bitmap
        cLayer.Alpha = 128 ' Set a default Alpha value
        cLayer.Visible = True ' Set the default of Visible to True

        LayerHandler.DrawLayers FinalDC ' Render all the layers onto the tempoary "FinalDC"
        picImage_Paint ' Force a refresh of the picturebox - FinalDC is BitBlted onto the picturebox via the Paint method

        lblLayersCap.Caption = "Layers: " & LayerHandler.Layers.Count ' Show the total layers

        cmbSelLayer.Enabled = True ' Enable the Select Layer combobox
        cmdFlattenImage.Enabled = True ' Enable the flatten image button

        cmbSelLayer.AddItem LayerHandler.Layers.Count ' Add the new layer to the Select Layer combobox
    End If

    CDialogs.FileName = vbNullString
End Sub

Private Sub LayerHandler_RenderedLayer(LayerNum As Integer, SingleLayerRender As Boolean)
    If LayerNum <= pgbRenderProgress.Max Then ' Only set the progressbar's value if it is less than or equal to the the maximum (this would occur when a layer other than the maximum is deleted)
        If SingleLayerRender = True Then ' If the RenderSingleLayer method called (preview picturebox is being rendered)
            pgbRenderProgress.Value = pgbRenderProgress.Max ' Set the progress bar's value to maximum
        Else ' All layers are being rendered
            pgbRenderProgress.Value = LayerNum ' Set the progressbar's value to the layer currently being rendered
        End If
    End If
End Sub

Private Sub LayerHandler_StartRender()
    If LayerHandler.Layers.Count > 0 Then ' Only set the progressbar's maximum if at least one layer present
        pgbRenderProgress.Max = LayerHandler.Layers.Count
    End If
End Sub

Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PrevOffsetX = X ' \ Save the current mouse position
    PrevOffsetY = Y ' / into tempoary variables

    If cmbSelLayer.Text <> vbNullString Then ' Layer selected
        If Button = 1 Then ' Left mouse button clicked
            DrawCircle X, Y ' Draw a circle onto the currently selected layer

            LayerHandler.DrawLayers FinalDC ' Draw the layers
            picImage_Paint ' Refresh the picturebox
            picLayerPreview_Paint ' Refresh the preview picturebox
        End If
    End If
End Sub

Private Sub picImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cLayer As ClsLayer

    If Button = 0 Then Exit Sub ' No button pressed, exit sub
    If cmbSelLayer.Text = vbNullString Then Exit Sub ' No layer selected, exit sub

    If Button = 1 Then ' Left button pressed - draw circles
        DrawCircle X, Y ' Draw a circle at the mouse pointer onto the selected layer

        picLayerPreview_Paint ' Refresh the preview picturebox
    Else ' Right button pressed, move layer
        Set cLayer = LayerHandler.Layers(Int(cmbSelLayer.Text)) ' Set the tempoary layer to the selected layer

        picImage.MousePointer = vbSizeAll ' Change the mouse cursor to a size all (a.k.a "Move") pointer

        cLayer.OffsetX = cLayer.OffsetX + ((X - PrevOffsetX) / Screen.TwipsPerPixelX) ' Move the layer by the X amount that the mouse has been moved
        cLayer.OffsetY = cLayer.OffsetY + ((Y - PrevOffsetY) / Screen.TwipsPerPixelY) ' Move the layer by the Y amount that the mouse has been moved

        PrevOffsetX = X ' \ Save the new current mouse position
        PrevOffsetY = Y ' / into tempoary variables

        Set cLayer = Nothing
    End If

    LayerHandler.DrawLayers FinalDC ' Draw the layers
    picImage_Paint ' Refresh the picturebox
End Sub

Private Sub picImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picImage.MousePointer = vbDefault ' Change the mouse cursor back to the default cursor
End Sub

Private Sub picImage_Paint()
    ' Painting of the picturebox is very quick because it just BitBlt's previously rendered data from a
    ' the "FinalDC" DC to the picturebox. This method stops redrawing flicker.

    BitBlt picImage.hdc, 0, 0, picImage.Width, picImage.Height, FinalDC, 0, 0, vbSrcCopy ' Copy the FinalDC to the picturebox
End Sub

Private Sub chkVisible_Click()
    ShowLayerChanges ' Checkbox value changed, update the layer properties and re-render
End Sub

Sub ShowLayerChanges() ' This sub updates the peoperties of the selected layer and re-renders all layers to reflect the changes
    Dim cLayer As ClsLayer
    Dim NewAlphaValue As Integer

    If cmbSelLayer.Text = vbNullString Then Exit Sub ' If no layer selected, exit sub

    Set cLayer = LayerHandler.Layers(Int(cmbSelLayer.Text)) ' Set the tempary layer to the selected layer

    If cmbLayerAlpha.Text <> vbNullString Then ' If an alpha value specified
        If InStr(1, cmbLayerAlpha.Text, "(") Then ' Bracketed percentage value present (user clicked dropdown sample value)
            NewAlphaValue = Int(Mid$(cmbLayerAlpha.Text, 1, InStr(1, cmbLayerAlpha.Text, "(") - 1)) ' Remove bracketed percentage value and get actual value
        Else
            NewAlphaValue = Int(cmbLayerAlpha.Text) ' Set the variable to the entered value
        End If

        If NewAlphaValue >= 0 And NewAlphaValue <= 255 Then ' Alpha value valid (0-255)
            cLayer.Alpha = Int(NewAlphaValue) ' Set the alpha value of the selected layer to the new alpha value
        End If
    End If

    cLayer.Visible = chkVisible.Value ' Set the layer visible attribute

    lblLayerIndex.Caption = cLayer.Index ' Set the layer index label caption to the index of the currently selected layer

    LayerHandler.DrawLayers FinalDC ' Render all the layers onto the tempoary DC "FinalDC"
    picImage_Paint ' Force a refresh of the picturebox
    picLayerPreview_Paint ' Refresh the preview picturebox

    Set cLayer = Nothing
End Sub

Private Sub picLayerPreview_Paint()
    If cmbSelLayer.Text <> "" Then ' If a layer has been selected
        LayerHandler.DrawSingleLayer picLayerPreview.hdc, Int(cmbSelLayer.Text), picLayerPreview.Width / Screen.TwipsPerPixelX, picLayerPreview.Height / Screen.TwipsPerPixelY, True ' Render the selected layer onto the preview picturebox
    End If
End Sub

Sub ClearLayerPreview()
    picLayerPreview.Cls ' Clear the preview picturebox since the layer has been deleted
    cmbLayerAlpha.Enabled = False ' Disable the select layer alpha combobox
    chkVisible.Enabled = False ' Disable the visible checkbox
    cmdDeleteLayer.Enabled = False ' Disable the Delete Layer button
    cmdFlip(0).Enabled = False ' Disable the Flip Horizontal button
    cmdFlip(1).Enabled = False ' Disable the Flip Vertical button
    cmdMove(0).Enabled = False ' Disable then Move Layer up button
    cmdMove(1).Enabled = False ' Disable then Move Layer down button
End Sub

Sub DrawCircle(X As Single, Y As Single)
    Dim TempDC As Long
    Dim TempBrush As Long
    Dim PrevObj As Long
    Dim cLayer As ClsLayer

    Set cLayer = LayerHandler.Layers(Int(cmbSelLayer.Text)) ' Set the tempoary layer to the selected layer

    TempDC = LayerHandler.RenderLayerToTempDC(cLayer.Index) ' Render the selected layer to a DC for manipulation

    TempBrush = CreateSolidBrush(vbRed) ' Create a new red brush
    PrevObj = SelectObject(TempDC, TempBrush) ' Load the created brush into the temp DC, save the handle for the previous brush

    Dim DrawX As Long
    Dim DrawY As Long

    DrawX = X / Screen.TwipsPerPixelX ' Calculate the X mouse coordinate
    DrawY = Y / Screen.TwipsPerPixelY ' Calculate the Y mouse coordinate

    DrawX = DrawX - cLayer.OffsetX ' Subtract the layer X position offset
    DrawY = DrawY - cLayer.OffsetY ' Subtract the layer Y position offset

    Ellipse TempDC, DrawX - 4, DrawY - 4, DrawX + 4, DrawY + 4 ' Draw the ellipse

    SelectObject TempDC, PrevObj ' Put the previous brush object back into the temp DC
    DeleteObject TempBrush ' Delete the created brush
    LayerHandler.CopyDCToBitmap TempDC, cLayer.BitmapHandle ' Put the changed temp DC bitmap back into the layer bitmap

    Set cLayer = Nothing
End Sub
