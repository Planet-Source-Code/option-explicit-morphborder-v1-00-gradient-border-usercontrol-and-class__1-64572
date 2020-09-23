VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMorphBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'* cMorphBorder v1.00 - VB6 form/control border display class.           *
'* Written by Matthew R. Usner, March, 2006.                             *
'* Copyright ©2006 - 2007, Matthew R. Usner.  All rights reserved.       *
'* Last - 10 March 2006 - fixed small glitch with even border widths.    *
'*************************************************************************
'* A small class that places a 3D gradient border around any control     *
'* that has the following properties: hDC, ScaleMode, ScaleHeight,       *
'* ScaleWidth and AutoRedraw.  Works with both intrinsic VB controls and *
'* usercontrols.                                                         *
'* Note:  If running in the IDE, DO NOT use the stop button in the IDE   *
'* toolbar.  Use Unload Me, not End, in code.                            *
'*************************************************************************
'* Legal:  Redistribution of this code, whole or in part, as source code *
'* or in binary form, alone or as part of a larger distribution or prod- *
'* uct, is forbidden for any commercial or for-profit use without the    *
'* author's explicit written permission.                                 *
'*                                                                       *
'* Non-commercial redistribution of this code, as source code or in      *
'* binary form, with or without modification, is permitted provided that *
'* the following conditions are met:                                     *
'*                                                                       *
'* Redistributions of source code must include this list of conditions,  *
'* and the following acknowledgment:                                     *
'*                                                                       *
'* This code was developed by Matthew R. Usner.                          *
'* Source code, written in Visual Basic, is freely available for non-    *
'* commercial, non-profit use.                                           *
'*                                                                       *
'* Redistributions in binary form, as part of a larger project, must     *
'* include the above acknowledgment in the end-user documentation.       *
'* Alternatively, the above acknowledgment may appear in the software    *
'* itself, if and where such third-party acknowledgments normally appear.*
'*************************************************************************
'* Credits and Thanks:                                                   *
'* LaVolpe, for the segment creation code.                               *
'* Carles P.V., for the gradient generation code.                        *
'*************************************************************************

Option Explicit    ' USE IT!  USE IT!!  USE IT!!!

' declares for creating, selecting, coloring and destroying the shaped border segment regions.
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long

' other graphics api declares.
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'  used in creating trapezoidal border segments.
Private Type POINTAPI
   X                              As Long
   Y                              As Long
End Type

'  used to define various graphics areas.
Private Type RECT
   Left                           As Long
   Top                            As Long
   Right                          As Long
   Bottom                         As Long
End Type

'  declares for gradient painting and bitmap tiling.
Private Type BITMAPINFOHEADER
   biSize                         As Long
   biWidth                        As Long
   biHeight                       As Long
   biPlanes                       As Integer
   biBitCount                     As Integer
   biCompression                  As Long
   biSizeImage                    As Long
   biXPelsPerMeter                As Long
   biYPelsPerMeter                As Long
   biClrUsed                      As Long
   biClrImportant                 As Long
End Type

' gradient generation constants.
Private Const DIB_RGB_COLORS      As Long = 0
Private Const PI                  As Single = 3.14159265358979
Private Const TO_DEG              As Single = 180 / PI
Private Const TO_RAD              As Single = PI / 180
Private Const INT_ROT             As Long = 1000

'  gradient information for horizontal and vertical border segments.
Private SegV1uBIH                  As BITMAPINFOHEADER
Private SegV1lBits()               As Long
Private SegV2uBIH                  As BITMAPINFOHEADER
Private SegV2lBits()               As Long
Private SegH1uBIH                  As BITMAPINFOHEADER
Private SegH1lBits()               As Long
Private SegH2uBIH                  As BITMAPINFOHEADER
Private SegH2lBits()               As Long

' constants defining the four border segments.
Private Const TOP_SEGMENT         As Long = 0
Private Const RIGHT_SEGMENT       As Long = 1
Private Const BOTTOM_SEGMENT      As Long = 2
Private Const LEFT_SEGMENT        As Long = 3

' holds region pointers for border segments.
Private BorderSegment(0 To 3)     As Long

' declares for virtual horizontal segment gradient bitmap.
Private VirtualDC_SegH            As Long                    ' handle of the created DC.
Private mMemoryBitmap_SegH        As Long                    ' handle of the created bitmap.
Private mOriginalBitmap_SegH      As Long                    ' used in destroying virtual DC.

' declares for virtual horizontal segment gradient bitmap.
Private VirtualDC_SegV            As Long                    ' handle of the created DC.
Private mMemoryBitmap_SegV        As Long                    ' handle of the created bitmap.
Private mOriginalBitmap_SegV      As Long                    ' used in destroying virtual DC.

Public Sub DisplayBorder(ByVal TargetDC As Long, ByVal TargetWidth As Long, ByVal TargetHeight As Long, _
                         ByVal BorderWidth As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal bMOut As Boolean)

'*************************************************************************
'* displays the four border segments at the specified X,Y location.      *
'*************************************************************************

'  initialize gradients and border segments.  I do this here so form borders
'  can be automatically recalculated on the fly when the form is resized.
   InitBorder TargetDC, TargetWidth, TargetHeight, BorderWidth, Color1, Color2, bMOut

'  display each border segment.
   DisplaySegment TargetDC, TargetWidth, TargetHeight, BorderWidth, TOP_SEGMENT, 0, 0, bMOut
   DisplaySegment TargetDC, TargetWidth, TargetHeight, BorderWidth, LEFT_SEGMENT, 0, 0, bMOut
   DisplaySegment TargetDC, TargetWidth, TargetHeight, BorderWidth, RIGHT_SEGMENT, TargetWidth - BorderWidth, 0, bMOut
   DisplaySegment TargetDC, TargetWidth, TargetHeight, BorderWidth, BOTTOM_SEGMENT, -1, TargetHeight - BorderWidth, bMOut

End Sub

Private Sub DisplaySegment(ByVal TargetDC As Long, ByVal TargetWidth As Long, ByVal TargetHeight As Long, _
                           ByVal BorderWidth As Long, ByVal SegmentNdx As Long, _
                           ByVal StartX As Long, ByVal StartY As Long, ByVal bMOut As Boolean)

'*************************************************************************
'* displays one border segment.  Border segment gradients are displayed  *
'* to virtual bitmaps on the fly so that correct gradient orientation    *
'* is maintained if the .MiddleOut property is set to False.             *
'*************************************************************************

'  position the border segment region in the correct location.
   MoveRegionToXY BorderSegment(SegmentNdx), StartX, StartY

   Select Case SegmentNdx

      Case LEFT_SEGMENT
         PaintVerticalGradient BorderWidth, TargetHeight, SegV1uBIH, SegV1lBits()
         BlitToRegion VirtualDC_SegV, TargetDC, BorderWidth, TargetHeight, BorderSegment(SegmentNdx), StartX, StartY

      Case RIGHT_SEGMENT
         If bMOut Then
            PaintVerticalGradient BorderWidth, TargetHeight, SegV1uBIH, SegV1lBits()
         Else
            PaintVerticalGradient BorderWidth, TargetHeight, SegV2uBIH, SegV2lBits()
         End If
         BlitToRegion VirtualDC_SegV, TargetDC, BorderWidth, TargetHeight, BorderSegment(SegmentNdx), StartX, StartY

      Case TOP_SEGMENT
         PaintHorizontalGradient BorderWidth, TargetWidth, SegH1uBIH, SegH1lBits()
         BlitToRegion VirtualDC_SegH, TargetDC, TargetWidth, BorderWidth, BorderSegment(SegmentNdx), StartX, StartY

      Case BOTTOM_SEGMENT
         If bMOut Then
            PaintHorizontalGradient BorderWidth, TargetWidth, SegH1uBIH, SegH1lBits()
         Else
            PaintHorizontalGradient BorderWidth, TargetWidth, SegH2uBIH, SegH2lBits()
         End If
         BlitToRegion VirtualDC_SegH, TargetDC, TargetWidth, BorderWidth, BorderSegment(SegmentNdx), StartX, StartY

   End Select

'  reset the region location to (0, 0) to prepare for the next time the segment is moved.
   MoveRegionToXY BorderSegment(SegmentNdx), 0, 0

End Sub

Private Sub MoveRegionToXY(ByVal Rgn As Long, ByVal X As Long, ByVal Y As Long)

'*************************************************************************
'* moves the supplied region to absolute X,Y coordinates.                *
'*************************************************************************

   Dim r As RECT    ' holds current X and Y coordinates of region.

'  get the current X,Y coordinates of the region.
   GetRgnBox Rgn, r

'  shift the region to 0,0 then to X,Y.
   OffsetRgn Rgn, -r.Left + X, -r.Top + Y

End Sub

Private Sub BlitToRegion(ByVal SourceDC As Long, DestDC As Long, lWidth As Long, lHeight As Long, Region As Long, ByVal XPos As Long, ByVal YPos As Long)

'*************************************************************************
'* blits the contents of a source DC to a non-rectangular region in a    *
'* destination DC.  A clipping region is selected in the destination DC, *
'* then the source DC is blitted to that location.  Technique is used in *
'* this control to blit to the trapezoid-shaped border regions.  Thanks  *
'* to LaVolpe for his help in tweaking this routine.                     *
'*************************************************************************

   Dim r              As Long    ' bitblt function call return.
   Dim ClippingRegion As Long    ' clipping region for bitblt.

'  move the region to the desired position.
   MoveRegionToXY Region, XPos, YPos

'  select a clipping region consisting of the segment parameter.
   ClippingRegion = SelectClipRgn(DestDC, Region)

'  blit the virtual bitmap to the control or form.  Since the clipping region has been
'  selected, only that region-shaped portion of the background will actually be drawn.
   r = BitBlt(DestDC, XPos, YPos, lWidth, lHeight, SourceDC, 0, 0, vbSrcCopy)

'  remove the clipping region constraint from the control.
   SelectClipRgn DestDC, ByVal 0&

'  reset the region coordinates to 0,0.
   MoveRegionToXY Region, 0, 0

End Sub

Private Sub InitBorder(ByVal TargetDC As Long, ByVal TargetWidth As Long, ByVal TargetHeight As Long, _
                       ByVal BorderWidth As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal bMOut As Boolean)

'*************************************************************************
'* creates virtual bitmaps, gradients and border segments.               *
'*************************************************************************

'  create virtual bitmaps that will hold the vertical and horizontal border segment gradient bitmaps.
   CreateVirtualDCs TargetDC, BorderWidth, TargetWidth, TargetHeight

'  determine border segment gradients.
   CalculateGradients BorderWidth, TargetWidth, TargetHeight, Color1, Color2, bMOut

'  create the four border segments.
   CreateBorderSegments BorderWidth, TargetWidth, TargetHeight

End Sub

Private Sub CreateVirtualDCs(ByVal TargetDC As Long, ByVal BorderWidth As Long, ByVal TargetWidth As Long, ByVal TargetHeight As Long)

'*************************************************************************
'* creates virtual DCs and corresponding virtual bitmaps that contain    *
'* the bitmap/gradient graphics for the control background and segments. *
'*************************************************************************

'  create the main value horizontal segment gradient virtual DC.
   CreateVirtualDC TargetDC, VirtualDC_SegH, _
                   mMemoryBitmap_SegH, mOriginalBitmap_SegH, _
                   TargetWidth + 1, BorderWidth

'  create the main value vertical segment gradient virtual DC.
   CreateVirtualDC TargetDC, VirtualDC_SegV, _
                   mMemoryBitmap_SegV, mOriginalBitmap_SegV, _
                   BorderWidth, TargetHeight

End Sub

Private Sub CreateVirtualDC(TargetDC As Long, vDC As Long, mMB As Long, mOB As Long, ByVal vWidth As Long, ByVal vHeight As Long)

'*************************************************************************
'* creates virtual bitmaps for background and cells.                     *
'*************************************************************************

   If IsCreated(vDC) Then
      DestroyVirtualDC vDC, mMB, mOB
   End If

'  create a memory device context to use.
   vDC = CreateCompatibleDC(TargetDC)

'  define it as a bitmap so that drawing can be performed to the virtual DC.
   mMB = CreateCompatibleBitmap(TargetDC, vWidth, vHeight)
   mOB = SelectObject(vDC, mMB)

End Sub

Private Function IsCreated(ByVal vDC As Long) As Boolean

'*************************************************************************
'* checks the handle of a virtual DC and returns if it exists.           *
'*************************************************************************

   IsCreated = (vDC <> 0)

End Function

Private Sub DestroyVirtualDC(ByRef vDC As Long, ByVal mMB As Long, ByVal mOB As Long)

'*************************************************************************
'* eliminates a virtual dc bitmap on control's termination.              *
'*************************************************************************

   If Not IsCreated(vDC) Then
      Exit Sub
   End If

   Call SelectObject(vDC, mOB)
   Call DeleteObject(mMB)
   Call DeleteDC(vDC)
   vDC = 0

End Sub

Private Sub CalculateGradients(ByVal BorderWidth As Long, ByVal TargetWidth As Long, ByVal TargetHeight As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal bMOut As Boolean)

'*************************************************************************
'* master routine for calculating various control gradients.             *
'*************************************************************************

'  calculate the primary horizontal segment gradient.
   CalculateGradient TargetWidth, BorderWidth + 1, TranslateColor(Color1), TranslateColor(Color2), _
                     90, bMOut, SegH1uBIH, SegH1lBits()

'  if gradients are not middle-out, calculate the secondary horizontal segment gradient.
   If Not bMOut Then
      CalculateGradient TargetWidth, BorderWidth + 1, TranslateColor(Color2), TranslateColor(Color1), _
                        90, bMOut, SegH2uBIH, SegH2lBits()
   End If

'  calculate the primary vertical segment gradient.
   CalculateGradient BorderWidth + 1, TargetHeight, TranslateColor(Color1), TranslateColor(Color2), _
                     180, bMOut, SegV1uBIH, SegV1lBits()

'  if gradients are not middle-out, calculate the secondary vertical segment gradient.
   If Not bMOut Then
      CalculateGradient BorderWidth + 1, TargetHeight, TranslateColor(Color2), TranslateColor(Color1), _
                        180, bMOut, SegV2uBIH, SegV2lBits()
   End If

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* converts color long COLORREF for api coloring purposes.               *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If

End Function

Private Sub PaintHorizontalGradient(ByVal BorderWidth As Long, ByVal TargetWidth As Long, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

'*************************************************************************
'* paints appropriate horizontal gradient to horizontal virtual bitmap.  *
'*************************************************************************

   Call StretchDIBits(VirtualDC_SegH, _
                      0, 0, _
                      TargetWidth, BorderWidth, _
                      0, 1, _
                      TargetWidth, BorderWidth - 1, _
                      lBits(0), uBIH, _
                      DIB_RGB_COLORS, _
                      vbSrcCopy)

End Sub

Private Sub PaintVerticalGradient(ByVal BorderWidth As Long, ByVal TargetHeight, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

'*************************************************************************
'* paints appropriate vertical gradient to vertical virtual bitmap.      *
'*************************************************************************

   Call StretchDIBits(VirtualDC_SegV, _
                      0, 0, _
                      BorderWidth, TargetHeight, _
                      1, 0, _
                      BorderWidth - 1, TargetHeight, _
                      lBits(0), uBIH, _
                      DIB_RGB_COLORS, _
                      vbSrcCopy)

End Sub

Private Sub CalculateGradient(Width As Long, Height As Long, _
                              ByVal Color1 As Long, ByVal Color2 As Long, _
                              ByVal Angle As Single, ByVal bMOut As Boolean, _
                              ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

'*************************************************************************
'* Carles P.V.'s routine, modified by Matthew R. Usner for middle-out    *
'* gradient capability.  Also modified to just calculate the gradient,   *
'* not draw it.  Original submission at PSC, txtCodeID=60580.            *
'*************************************************************************

   Dim lGrad()   As Long, lGrad2() As Long

   Dim lClr      As Long
   Dim R1        As Long, G1 As Long, b1 As Long
   Dim R2        As Long, G2 As Long, b2 As Long
   Dim dR        As Long, dG As Long, dB As Long

   Dim Scan      As Long
   Dim i         As Long, j As Long, k As Long
   Dim jIn       As Long
   Dim iEnd      As Long, jEnd As Long
   Dim Offset    As Long

   Dim lQuad     As Long
   Dim AngleDiag As Single
   Dim AngleComp As Single

   Dim g         As Long
   Dim luSin     As Long, luCos As Long
 
   If (Width > 0 And Height > 0) Then

'     when angle is >= 91 and <= 270, the colors
'     invert in MiddleOut mode.  This corrects that.
      If bMOut And Angle >= 91 And Angle <= 270 Then
         g = Color1
         Color1 = Color2
         Color2 = g
      End If

'     -- Right-hand [+] (ox=0º)
      Angle = -Angle + 90

'     -- Normalize to [0º;360º]
      Angle = Angle Mod 360
      If (Angle < 0) Then
         Angle = 360 + Angle
      End If

'     -- Get quadrant (0 - 3)
      lQuad = Angle \ 90

'     -- Normalize to [0º;90º]
        Angle = Angle Mod 90

'     -- Calc. gradient length ('distance')
      If (lQuad Mod 2 = 0) Then
         AngleDiag = Atn(Width / Height) * TO_DEG
      Else
         AngleDiag = Atn(Height / Width) * TO_DEG
      End If
      AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
      Angle = Angle * TO_RAD
      g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem

'     -- Decompose colors
      If (lQuad > 1) Then
         lClr = Color1
         Color1 = Color2
         Color2 = lClr
      End If
      R1 = (Color1 And &HFF&)
      G1 = (Color1 And &HFF00&) \ 256
      b1 = (Color1 And &HFF0000) \ 65536
      R2 = (Color2 And &HFF&)
      G2 = (Color2 And &HFF00&) \ 256
      b2 = (Color2 And &HFF0000) \ 65536

'     -- Get color distances
      dR = R2 - R1
      dG = G2 - G1
      dB = b2 - b1

'     -- Size gradient-colors array
      ReDim lGrad(0 To g - 1)
      ReDim lGrad2(0 To g - 1)

'     -- Calculate gradient-colors
      iEnd = g - 1
      If (iEnd = 0) Then
'        -- Special case (1-pixel wide gradient)
         lGrad2(0) = (b1 \ 2 + b2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
         For i = 0 To iEnd
            lGrad2(i) = b1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
         Next i
      End If

'     'if' block added by Matthew R. Usner - accounts for possible MiddleOut gradient draw.
      If bMOut Then
         k = 0
         For i = 0 To iEnd Step 2
            lGrad(k) = lGrad2(i)
            k = k + 1
         Next i
         For i = iEnd - 1 To 1 Step -2
            lGrad(k) = lGrad2(i)
            k = k + 1
         Next i
      Else
         For i = 0 To iEnd
            lGrad(i) = lGrad2(i)
         Next i
      End If

'     -- Size DIB array
      ReDim lBits(Width * Height - 1) As Long
      iEnd = Width - 1
      jEnd = Height - 1
      Scan = Width

'     -- Render gradient DIB
      Select Case lQuad

         Case 0, 2
            luSin = Sin(Angle) * INT_ROT
            luCos = Cos(Angle) * INT_ROT
            Offset = 0
            jIn = 0
            For j = 0 To jEnd
               For i = 0 To iEnd
                  lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
               Next i
               jIn = jIn + luCos
               Offset = Offset + Scan
            Next j

         Case 1, 3
            luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
            luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
            Offset = jEnd * Scan
            jIn = 0
            For j = 0 To jEnd
               For i = 0 To iEnd
                  lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
               Next i
               jIn = jIn + luCos
               Offset = Offset - Scan
            Next j

      End Select

'     -- Define DIB header
      With uBIH
         .biSize = 40
         .biPlanes = 1
         .biBitCount = 32
         .biWidth = Width
         .biHeight = Height
      End With

   End If

End Sub

Private Sub CreateBorderSegments(ByVal BorderWidth As Long, ByVal TargetWidth As Long, ByVal TargetHeight As Long)

'*************************************************************************
'* creates the vertical and horizontal trapezoidal border segments.      *
'*************************************************************************

   DeleteBorderSegmentObjects

   BorderSegment(TOP_SEGMENT) = CreateDiagRectRegion(TargetWidth, BorderWidth, 1, 1)
   BorderSegment(BOTTOM_SEGMENT) = CreateDiagRectRegion(TargetWidth, BorderWidth, -1, -1)
   BorderSegment(RIGHT_SEGMENT) = CreateDiagRectRegion(BorderWidth, TargetHeight, -1, -1)
   BorderSegment(LEFT_SEGMENT) = CreateDiagRectRegion(BorderWidth, TargetHeight, 1, 1)

End Sub

Private Sub DeleteBorderSegmentObjects()

'*************************************************************************
'* destroys the border segment objects if they exist, to save memory.    *
'*************************************************************************

   If BorderSegment(TOP_SEGMENT) Then
      DeleteObject BorderSegment(TOP_SEGMENT)
   End If

   If BorderSegment(RIGHT_SEGMENT) Then
      DeleteObject BorderSegment(RIGHT_SEGMENT)
   End If

   If BorderSegment(BOTTOM_SEGMENT) Then
      DeleteObject BorderSegment(BOTTOM_SEGMENT)
   End If

   If BorderSegment(LEFT_SEGMENT) Then
      DeleteObject BorderSegment(LEFT_SEGMENT)
   End If

End Sub

Private Function CreateDiagRectRegion(ByVal cx As Long, ByVal cy As Long, SideAStyle As Integer, SideBStyle As Integer) As Long

'**************************************************************************
'* Author: LaVolpe                                                        *
'* the cx & cy parameters are the respective width & height of the region *
'* the passed values may be modified which coder can use for other purp-  *
'* oses like drawing borders or calculating the client/clipping region.   *
'* SideAStyle is -1, 0 or 1 depending on horizontal/vertical shape,       *
'*            reflects the left or top side of the region                 *
'*            -1 draws left/top edge like /                               *
'*            0 draws left/top edge like  |                               *
'*            1 draws left/top edge like  \                               *
'* SideBStyle is -1, 0 or 1 depending on horizontal/vertical shape,       *
'*            reflects the right or bottom side of the region             *
'*            -1 draws right/bottom edge like \                           *
'*            0 draws right/bottom edge like  |                           *
'*            1 draws right/bottom edge like  /                           *
'**************************************************************************

   Dim tpts(0 To 4) As POINTAPI    ' holds polygonal region vertices.

   If cx > cy Then ' horizontal

'     absolute minimum width & height of a trapezoid
      If Abs(SideAStyle + SideBStyle) = 2 Then ' has 2 opposing slanted sides
         If cx < cy * 2 Then cy = cx \ 2
      End If

      If SideAStyle < 0 Then
         tpts(0).X = cy - 1
         tpts(1).X = -1
      ElseIf SideAStyle > 0 Then
         tpts(1).X = cy
      End If
      tpts(1).Y = cy

      tpts(2).X = cx + Abs(SideBStyle < 0)
      If SideBStyle > 0 Then tpts(2).X = tpts(2).X - cy
      tpts(2).Y = cy

      tpts(3).X = cx + Abs(SideBStyle < 0)
      If SideBStyle < 0 Then tpts(3).X = tpts(3).X - cy

   Else

'     absolute minimum width & height of a trapezoid
      If Abs(SideAStyle + SideBStyle) = 2 Then ' has 2 opposing slanted sides
         If cy < cx * 2 Then cx = cy \ 2
      End If

      If SideAStyle < 0 Then
         tpts(0).Y = cx - 1
         tpts(3).Y = -1
      ElseIf SideAStyle > 0 Then
         tpts(3).Y = cx - 1
         tpts(0).Y = -1
      End If

      tpts(1).Y = cy
      If SideBStyle < 0 Then tpts(1).Y = tpts(1).Y - cx
      tpts(2).X = cx

      tpts(2).Y = cy
      If SideBStyle > 0 Then tpts(2).Y = tpts(2).Y - cx
      tpts(3).X = cx

   End If

   tpts(4) = tpts(0)

   CreateDiagRectRegion = CreatePolygonRgn(tpts(0), UBound(tpts) + 1, 2)

End Function

Private Sub Class_Terminate()

'*************************************************************************
'* destroys all active objects and regions prior to class termination.   *
'*************************************************************************

'  delete border segment region objects.
   DeleteBorderSegmentObjects

'  destroy the virtual DC's used to store segment gradients.
   DestroyVirtualDC VirtualDC_SegH, mMemoryBitmap_SegH, mOriginalBitmap_SegH
   DestroyVirtualDC VirtualDC_SegV, mMemoryBitmap_SegV, mOriginalBitmap_SegV

End Sub
