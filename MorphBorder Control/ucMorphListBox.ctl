VERSION 5.00
Begin VB.UserControl MorphListBox 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   122
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "ucMorphListBox.ctx":0000
End
Attribute VB_Name = "MorphListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*************************************************************************
'* MorphListBox v1.15 - VB6 ownerdrawn listbox control replacement.      *
'* Written by Matthew R. Usner, September, 2005.                         *
'* Copyright ©2005 - 2006, Matthew R. Usner.  All rights reserved.       *
'* Last: 02/25/06 - Added listitem icon display capability.              *
'*************************************************************************
'* A graphical replacement for the boring VB listbox control.  Standard  *
'* listbox behavior is consistently emulated, with the exception of a    *
'* couple intricacies of list item selection in .MultiSelect Extended    *
'* mode which I thought were a tad overwrought and unnecessary.  Loads   *
'* lists much faster than standard listbox, depending on system. Control *
'* features an integrated graphical vertical scrollbar.  Background can  *
'* be a gradient or bitmap.  Small bitmaps can be tiled or stretched.    *
'* Icons or small bitmaps can be assigned and displayed next to desired  *
'* list items.  >32767 listbox items, limited only by system resources.  *
'* Eight gradient color schemes can be selected via .Theme property.     *
'* Unicode character display supported.  Basic drag-and-drop capability  *
'* is included. Sort order can be maintained numerically for lists of    *
'* numbers, although numbers are still stored as strings.  Any portion   *
'* of list can be displayed via .DisplayFrom method.  Also has a         *
'* .MouseOverIndex method that allows determination of the list item     *
'* (index or item) the mouse cursor is hovering over.                    *
'*************************************************************************
'* Miscellaneous Usage Notes:                                            *
'*   1) Due to size of control (>300 procedures), this is best used as a *
'*      compiled .OCX.                                                   *
'*   2) When filling the MorphListBox with a large number of list items, *
'*      set the .RedrawFlag property to False prior to the loop that     *
'*      fills the list.  Afterwards, set .RedrawFlag to True.  This is a *
'*      big timesaver because list won't redraw after adding each item.  *
'*   3) Since this uses subclassing, use Unload Me, not End, in projects.*
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
'* Carles P.V., for the gradient, bitmap tiling, and corner rounding     *
'* routines.                                                             *
'* Paul Caton, for the self-subclassing usercontrol code.                *
'* Phillipe Lord, for his array handling routines.  His original module  *
'* can be found at PSC, txtCodeId=24546.                                 *
'* Richard Mewett, for the Unicode support routines.                     *
'* Paul Turcksin, for spending hours checking this before I submitted.   *
'* Jeff Mayes, for the .SortAsNumeric idea.                              *
'* Redbird77, for fixing a glitch with the background gradient draw and  *
'* reorganizing and optimizing the DisplayListBoxItem routine.           *
'*************************************************************************

Option Explicit

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const MOUSEEVENTF_LEFTDOWN = &H2        ' for generating a mousedown event to replace double click.
Private m_hBrush As Long                        ' pattern brush for bitmap tiling.

'==================================================================================================
' Subclasser declarations.
' windows messages to be intercepted by subclassing.
Private Const WM_MOUSEMOVE            As Long = &H200
Private Const WM_MOUSELEAVE           As Long = &H2A3
Private Const WM_SETFOCUS             As Long = &H7
Private Const WM_KILLFOCUS            As Long = &H8
Private Const WM_MOUSEWHEEL           As Long = &H20A

Private Enum TRACKMOUSEEVENT_FLAGS
   TME_HOVER = &H1&
   TME_LEAVE = &H2&
   TME_QUERY = &H40000000
   TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
   cbSize                             As Long
   dwFlags                            As TRACKMOUSEEVENT_FLAGS
   hwndTrack                          As Long
   dwHoverTime                        As Long
End Type

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private bInCtrl                       As Boolean          ' flag that indicates if mouse is in control.

Private Enum eMsgWhen
   MSG_AFTER = 1                                          'Message calls back after the original (previous) WndProc
   MSG_BEFORE = 2                                         'Message calls back before the original (previous) WndProc
   MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE         'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES            As Long = -1        'All messages added or deleted
Private Const GMEM_FIXED              As Long = 0         'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC             As Long = -4        'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04                As Long = 88        'Table B (before) address patch offset
Private Const PATCH_05                As Long = 93        'Table B (before) entry count patch offset
Private Const PATCH_08                As Long = 132       'Table A (after) address patch offset
Private Const PATCH_09                As Long = 137       'Table A (after) entry count patch offset

Private Type tSubData                                     'Subclass data type
   hwnd                               As Long             'Handle of the window being subclassed
   nAddrSub                           As Long             'The address of our new WndProc (allocated memory).
   nAddrOrig                          As Long             'The address of the pre-existing WndProc
   nMsgCntA                           As Long             'Msg after table entry count
   nMsgCntB                           As Long             'Msg before table entry count
   aMsgTblA()                         As Long             'Msg after table array
   aMsgTblB()                         As Long             'Msg Before table array
End Type
Private sc_aSubData()                 As tSubData         'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'==================================================================================================

'  declares for Unicode support.
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
   dwOSVersionInfoSize                As Long
   dwMajorVersion                     As Long
   dwMinorVersion                     As Long
   dwBuildNumber                      As Long
   dwPlatformId                       As Long
   szCSDVersion                       As String * 128     '  Maintenance string for PSS usage
End Type
Private mWindowsNT                    As Boolean
Private Const DT_SINGLELINE           As Long = &H20      ' strip cr/lf from string before draw.
Private Const DT_NOPREFIX             As Long = &H800     ' ignore access key ampersand.
Private Const DT_LEFT                 As Long = &H0       ' draw from left edge of rectangle.

' declares for gradient painting and bitmap tiling.
Private Type BITMAPINFOHEADER
   biSize          As Long
   biWidth         As Long
   biHeight        As Long
   biPlanes        As Integer
   biBitCount      As Integer
   biCompression   As Long
   biSizeImage     As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed       As Long
   biClrImportant  As Long
End Type

Private Type BITMAP
   bmType       As Long
   bmWidth      As Long
   bmHeight     As Long
   bmWidthBytes As Long
   bmPlanes     As Integer
   bmBitsPixel  As Integer
   bmBits       As Long
End Type

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0                  ' also used in gradient generation.
Private Const OBJ_BITMAP     As Long = 7                  ' used to determine if picture is a bitmap.

'  used to define various graphics areas and listbox component locations.
Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

' enum tied to the .MultiSelect property.
Public Enum SelectionOptions
   [None] = 0
   [Simple] = 1
   [Extended] = 2
End Enum

' enum tied to the .Style property.
Public Enum ListItemOptions
   [Standard] = 0
   [CheckBox] = 1
End Enum

' enum tied to .Theme property.
Public Enum ThemeOptions
   [None] = 0
   [Cyan Eyed] = 1
   [Gunmetal Grey] = 2
   [Blue Moon] = 3
   [Red Rum] = 4
   [Green With Envy] = 5
   [Purple People Eater] = 6
   [Golden Goose] = 7
   [Penny Wise] = 8
End Enum

'  enum tied to the .DoubleClickBehavior property.
Public Enum DblClickBehaviorOptions
   [Double Click] = 0
   [Two Single Clicks] = 1
End Enum

'  enum tied to .PictureMode property.
Public Enum PictureModeOptions
   [Normal] = 0
   [Stretch] = 1
   [Tiled] = 2
End Enum

'  enum tied to .CheckStyle property.
Public Enum CheckStyleOptions
   [Arrow] = 0
   [Tick] = 1
   [X] = 2
End Enum

'  list display declares.
'  type used to hold indices of the first and last displayable list items, based on current listbox height.
Private Type DisplayRangeType
   FirstListItem As Long                                  ' the index of the first displayed list item.
   LastListItem As Long                                   ' the index of the last displayed list item.
End Type
Private DisplayRange               As DisplayRangeType    ' indices of first and last displayed listbox items.

Private ListItemHeight             As Long                ' height of text in current font.
Private Const Y_CLEARANCE          As Long = 5            ' pixel offset to start and stop displaying list items
Private YCoords(0 To 99)           As Long                ' y coordinate of each displayable list item.
Private MaxDisplayItems            As Long                ' max displayable based on height, font, borderwidth.
Private TextClearance              As Long                ' # pixels from left edge to start drawing text.
Private Const MIN_FONT_HEIGHT      As Long = 13           ' helps adjust item spacing when using very small fonts.
Private ChangingPicture            As Boolean             ' if true, re-blits new picture to virtual DC.
Private Const NO_IMAGE             As Long = -1           ' no image assigned to a particular listitem.
Private Const EQUAL_TO_TEXTHEIGHT  As Long = 0            ' item image will be height/width of text height.

'  'active' variables - these are the values actually used in displaying the control.
'  The Enabled or Disabled color property sets are transferred into these variables.
Private m_ActiveArrowDownColor     As OLE_COLOR
Private m_ActiveArrowUpColor       As OLE_COLOR
Private m_ActiveBackColor1         As OLE_COLOR
Private m_ActiveBackColor2         As OLE_COLOR
Private m_ActiveBorderColor        As OLE_COLOR
Private m_ActiveButtonColor1       As OLE_COLOR
Private m_ActiveButtonColor2       As OLE_COLOR
Private m_ActiveCheckboxArrowColor As OLE_COLOR
Private m_ActiveCheckBoxColor      As OLE_COLOR
Private m_ActiveFocusRectColor     As OLE_COLOR
Private m_ActivePicture            As StdPicture
Private m_ActivePictureMode        As PictureModeOptions
Private m_ActiveSelColor1          As OLE_COLOR
Private m_ActiveSelColor2          As OLE_COLOR
Private m_ActiveSelTextColor       As OLE_COLOR
Private m_ActiveTextColor          As OLE_COLOR
Private m_ActiveThumbBorderColor   As OLE_COLOR
Private m_ActiveThumbColor1        As OLE_COLOR
Private m_ActiveThumbColor2        As OLE_COLOR
Private m_ActiveTrackBarColor1     As OLE_COLOR
Private m_ActiveTrackBarColor2     As OLE_COLOR

'  property variables.
Private m_ItemImageSize As Long
Private m_ShowItemImages        As Boolean
Private m_TopIndex              As Long
Private m_DisArrowDownColor     As OLE_COLOR
Private m_DisArrowUpColor       As OLE_COLOR
Private m_DisBackColor1         As OLE_COLOR
Private m_DisBackColor2         As OLE_COLOR
Private m_DisBorderColor        As OLE_COLOR
Private m_DisButtonColor1       As OLE_COLOR
Private m_DisButtonColor2       As OLE_COLOR
Private m_DisCheckboxArrowColor As OLE_COLOR
Private m_DisCheckboxColor      As OLE_COLOR
Private m_DisFocusRectColor     As OLE_COLOR
Private m_DisPicture            As Picture
Private m_DisPictureMode        As PictureModeOptions
Private m_DisSelColor1          As OLE_COLOR
Private m_DisSelColor2          As OLE_COLOR
Private m_DisSelTextColor       As OLE_COLOR
Private m_DisTextColor          As OLE_COLOR
Private m_DisThumbBorderColor   As OLE_COLOR
Private m_DisThumbColor1        As OLE_COLOR
Private m_DisThumbColor2        As OLE_COLOR
Private m_DisTrackbarColor1     As OLE_COLOR
Private m_DisTrackbarColor2     As OLE_COLOR
Private m_SortAsNumeric         As Boolean                   ' is list sorted as string or numbers?
Private m_TrackClickColor1      As OLE_COLOR                 ' track portion clicked first color.
Private m_TrackClickColor2      As OLE_COLOR                 ' track portion clicked second color.
Private m_DragEnabled           As Boolean                   ' boolean that allows drag and drop.
Private m_CheckStyle            As CheckStyleOptions         ' arrow and checkmark style options.
Private m_PictureMode           As PictureModeOptions        ' normal, stretched or tiled picture display.
Private m_NewIndex              As Long                      ' index of most recently added item.
Private m_DblClickBehavior      As DblClickBehaviorOptions   ' double click or two rapid single clicks?
Private m_FocusRectColor        As OLE_COLOR                 ' custom focus rectangle color.
Private m_CheckBoxColor         As OLE_COLOR                 ' checkbox border color.
Private m_CheckboxArrowColor    As OLE_COLOR                 ' selection checkbox arrow color.
Private m_Theme                 As ThemeOptions              ' color scheme to use.
Private m_ArrowUpColor          As OLE_COLOR                 ' scroll arrow color when button is up.
Private m_ArrowDownColor        As OLE_COLOR                 ' scroll arrow color when button is down.
Private m_ThumbBorderColor      As OLE_COLOR                 ' border color for scroll thumb.
Private m_ThumbColor1           As OLE_COLOR                 ' first scrollbar thumb gradient color.
Private m_ThumbColor2           As OLE_COLOR                 ' second scrollbar thumb gradient color.
Private m_ButtonColor1          As OLE_COLOR                 ' first scrollbar button gradient color.
Private m_ButtonColor2          As OLE_COLOR                 ' second scrollbar button gradient color.
Private m_TrackBarColor1        As OLE_COLOR                 ' first trackbar gradient color.
Private m_TrackBarColor2        As OLE_COLOR                 ' second trackbar gradient color.
Private m_SelCount              As Long                      ' read-only selected item counter.
Private m_Style                 As ListItemOptions           ' standard or checkbox listbox style.
Private m_MultiSelect           As SelectionOptions          ' non/simple/extended list item selection.
Private m_SelColor1             As OLE_COLOR                 ' first selection bar gradient color.
Private m_SelColor2             As OLE_COLOR                 ' second selection bar gradient color.
Private m_SelTextColor          As OLE_COLOR                 ' color to draw selected list item text.
Private m_TextColor             As OLE_COLOR                 ' color to draw unselected list item text.
Private m_RedrawFlag            As Boolean                   ' internal property for redraw yes/no.
Private m_Sorted                As Boolean                   ' if True, new items put in proper order.
Private m_ListFont              As Font                      ' the font to display listbox items with.
Private m_ListIndex             As Long                      ' index of currently selected item; -1 if none selected.
Private m_ListCount             As Long                      ' the number of items in the list.
Private m_Picture               As Picture                   ' the picture to use in lieu of gradient background.
Private m_CurveTopLeft          As Long                      ' the curvature of the top left corner.
Private m_CurveTopRight         As Long                      ' the curvature of the top right corner.
Private m_CurveBottomLeft       As Long                      ' the curvature of the bottom left corner.
Private m_CurveBottomRight      As Long                      ' the curvature of the bottom right corner.
Private m_BackMiddleOut         As Boolean                   ' flag for container background middle-out gradient.
Private m_Enabled               As Boolean                   ' enabled/disabled flag.
Private m_BackAngle             As Single                    ' background gradient display angle
Private m_BackColor1            As OLE_COLOR                 ' the first gradient color of the background.
Private m_BackColor2            As OLE_COLOR                 ' the second gradient color of the background.
Private m_BorderWidth           As Integer                   ' width, in pixels, of border.
Private m_BorderColor           As OLE_COLOR                 ' color of border.

'  default property constants.
Private Const m_def_ItemImageSize = 0
Private Const m_def_ShowItemImages = False
Private Const m_def_TopIndex = 0
Private Const m_def_DisArrowDownColor = &HC0C0C0
Private Const m_def_DisArrowUpColor = &HC0C0C0
Private Const m_def_DisBackColor1 = &H808080
Private Const m_def_DisBackColor2 = &HC0C0C0
Private Const m_def_DisBorderColor = &H0
Private Const m_def_DisButtonColor1 = &H404040
Private Const m_def_DisButtonColor2 = &H808080
Private Const m_def_DisCheckboxArrowColor = &H0
Private Const m_def_DisCheckboxColor = &H0
Private Const m_def_DisFocusRectColor = &H808080
Private Const m_def_DisPictureMode = 0
Private Const m_def_DisSelColor1 = &H808080
Private Const m_def_DisSelColor2 = &HC0C0C0
Private Const m_def_DisSelTextColor = &H808080
Private Const m_def_DisTextColor = &H404040
Private Const m_def_DisThumbBorderColor = &H808080
Private Const m_def_DisThumbColor1 = &H404040
Private Const m_def_DisThumbColor2 = &H808080
Private Const m_def_DisTrackbarColor1 = &H808080
Private Const m_def_DisTrackbarColor2 = &HC0C0C0
Private Const m_def_SortAsNumeric = False                 ' default sort as string, not numeric.
Private Const m_def_TrackClickColor1 = &H0                ' default track portion clicked first color.
Private Const m_def_TrackClickColor2 = &HE0E0E0           ' default track portion clicked second color.
Private Const m_def_DragEnabled = False                   ' no drag and drop is default.
Private Const m_def_CheckStyle = 1                        ' checkmark check style default.
Private Const m_def_PictureMode = 0                       ' normal picture display (not stretched/tiled).
Private Const m_def_NewIndex = -1                         ' indicating list is empty (no new added).
Private Const m_def_DblClickBehavior = 1                  ' default is two rapid single clicks.
Private Const m_def_FocusRectColor = &H0                  ' black focus rectangle.
Private Const m_def_CheckBoxColor = &H0                   ' black checkbox border color.
Private Const m_def_CheckboxArrowColor = &H0              ' black check arrow color.
Private Const m_def_Theme = 2                             ' gunmetal grey default color scheme.
Private Const m_def_ArrowUpColor = &HE0E0E0               ' light grey arrow up color.
Private Const m_def_ArrowDownColor = &H0                  ' black arrow down color.
Private Const m_def_ThumbBorderColor = &HE0E0E0           ' light grey thumb border color.
Private Const m_def_ThumbColor1 = &H0                     ' black first thumb color.
Private Const m_def_ThumbColor2 = &H909090                ' medium grey second thumb colot.
Private Const m_def_ButtonColor1 = &H0                    ' black start color.
Private Const m_def_ButtonColor2 = &HC0C0C0               ' grey end color.
Private Const m_def_TrackBarColor1 = &H606060             ' darker grey start color.
Private Const m_def_TrackBarColor2 = &HE0E0E0             ' lighter grey end color.
Private Const m_def_SelCount = 0                          ' read-only selected item counter.
Private Const m_def_Style = vbListBoxStandard             ' no checkbox.
Private Const m_def_MultiSelect = vbMultiSelectNone       ' one selection at a time.
Private Const m_def_SelColor1 = &HC0FFFF                  ' lighter amber selection bar first gradient color.
Private Const m_def_SelColor2 = &HC0FFFF                  ' lighter amber selection bar second gradient color.
Private Const m_def_SelTextColor = &H0                    ' black selected text color.
Private Const m_def_TextColor = &H0                       ' black text color.
Private Const m_def_RedrawFlag = True                     ' internal redraw flag to True.
Private Const m_def_Sorted = True                         ' list sorting by default.
Private Const m_def_ListIndex = -1                        ' no items selected.
Private Const m_def_CurveTopLeft = 0                      ' no top left curvature.
Private Const m_def_CurveTopRight = 0                     ' no top right curvature.
Private Const m_def_CurveBottomLeft = 0                   ' no bottom left curvature.
Private Const m_def_CurveBottomRight = 0                  ' no bottom right curvature.
Private Const m_def_BackMiddleOut = True                  ' middle-out background gradient.
Private Const m_def_Enabled = True                        ' enabled.
Private Const m_def_BackAngle = 45                        ' horizontal background gradient.
Private Const m_def_BackColor1 = &H606060                 ' darker grey start color.
Private Const m_def_BackColor2 = &HE0E0E0                 ' lighter grey end color.
Private Const m_def_BorderWidth = 1                       ' border width 1 pixel.
Private Const m_def_BorderColor = &H0                     ' black listbox border.

'  events.
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Resize()
Public Event MouseEnter()
Public Event MouseLeave()

Private HasFocus As Boolean       ' master 'control has focus' flag.

'  gradient generation constants.
Private Const RGN_DIFF              As Long = 4
Private Const PI                    As Single = 3.14159265358979
Private Const TO_DEG                As Single = 180 / PI
Private Const TO_RAD                As Single = PI / 180
Private Const INT_ROT               As Long = 1000

'  gradient information for background.
Private BGuBIH                      As BITMAPINFOHEADER
Private BGlBits()                   As Long

'  gradient information for list item selection bar.
Private SeluBIH                     As BITMAPINFOHEADER
Private SellBits()                  As Long

'  gradient information for vertical trackbar.
Private VTrackuBIH                  As BITMAPINFOHEADER
Private VTracklBits()               As Long

'  gradient information for clicked portion of vertical trackbar.
Private vClickTrackuBIH             As BITMAPINFOHEADER
Private vClickTracklBits()          As Long

'  gradient information for scrollbar buttons.
Private TrackButtonuBIH             As BITMAPINFOHEADER
Private TrackButtonlBits()          As Long

'  gradient information for vertical scrollbar thumb.
Private vThumbuBIH                  As BITMAPINFOHEADER
Private vThumblBits()               As Long

Private Const ScrollBarButtonHeight As Long = 15          ' the height, in pixels, of scrollbar button.
Private Const ScrollBarButtonWidth  As Long = 15          ' the width, in pixels, of scrollbar button.
Private vScrollTrackHeight          As Long               ' the height, in pixels, of thumb scroll track.
Private Const vScrollMinThumbHeight As Long = 9           ' keep it an odd number so middle isn't between pixels.
Private VerticalScrollBarActive     As Boolean            ' indicates scrollbar is drawn.
Private vThumbHeight                As Long               ' the height in pixels of the vertical thumb.
Private RecalculateThumbHeight      As Boolean            ' prevents unnecessary thumb height recalulations.
Private ThumbYPos                   As Long               ' y coordinate of top of scrollbar thumb.

' structure for containing exact scrollbar component location info for mouseover.
Private Type ScrollBarLocationType
   UpButtonLocation                 As RECT
   DownButtonLocation               As RECT
   ScrollTrackLocation              As RECT
   ScrollThumbLocation              As RECT
End Type
Private vScrollBarLocation As ScrollBarLocationType

' keyboard, mouse, and list item tracking variables.
Private ItemWithFocus                    As Long          ' the item in the list that has the "virtual focus".
Private CtrlKeyDown                      As Boolean       ' global "control key being pressed" flag.
Private ShiftKeyDown                     As Boolean       ' global "shift key is being pressed" flag.
Private LastSelectedItem                 As Long          ' item last clicked or otherwise selected.
Private RightClickFlag                   As Boolean       ' doubleclick bypass (DblClick detects right click too)
Private MouseAction                      As Long          ' set to value of one of below constants.
Private Const MOUSE_NOACTION             As Long = 0      ' mouse button is not down.
Private Const MOUSE_DOWNED_IN_LIST       As Long = 1      ' mouse downed in text portion of listbox.
Private Const MOUSE_DOWNED_IN_UPPERTRACK As Long = 2      ' mouse downed in trackbar above thumb.
Private Const MOUSE_DOWNED_IN_LOWERTRACK As Long = 3      ' mouse downed in trackbar below thumb.
Private Const MOUSE_DOWNED_IN_DOWNBUTTON As Long = 4      ' mouse downed in down scrollbar button.
Private Const MOUSE_DOWNED_IN_UPBUTTON   As Long = 5      ' mouse downed in up scrollbar button.
Private Const MOUSE_DOWNED_IN_THUMB      As Long = 6      ' mouse downed in scrollbar thumb.

Private FirstExtendedSelection      As Long               ' in Extended mode, the original list item clicked on.
Private LastExtendedSelection       As Long               ' in Extended mode, the last list item clicked on.
Private ItemMouseIsIn               As Long               ' to prevent redraws when mouse moves in same list item.
Private ShiftDownStartItem          As Long               ' item with focus when shift key is pressed.

' flags indicating how list items should be drawn.
Private Const DrawAsSelected        As Boolean = True     ' draw list item with selection bar gradient.
Private Const DrawAsUnselected      As Boolean = False    ' draw list item without selection bar gradient.
Private Const FocusRectangleYes     As Boolean = True     ' draw item with focus rectangle.
Private Const FocusRectangleNo      As Boolean = False    ' draw item without focus rectangle.
Private Const KeepSelectionAsIs     As Boolean = True     ' draw item, keeping item's selection status.
Private Const KeepSelectionNo       As Boolean = False    ' draw item, don't keep item's selection status.

' the arrays for the properties .List, .ItemData and .Selected.
Private ListArray()                 As String             ' tied to the .List property.
Private ItemDataArray()             As Long               ' tied to the .ItemData property.
Private SelectedArray()             As Boolean            ' tied to the .Selected property.
Private ImageIndexArray()           As Long               ' tied to the .ImageIndex property.

' array that holds the images that are displayed by listitems.
Private Images()                    As StdPicture
Private ImageCount                  As Long
Private PicX                        As Long               ' x coordinate of listitem image.

Private lBarWid                     As Long
Private SelBarOffset                As Long

' for keeping track of where the mouse is at any given time.
Private Const OVER_BORDER           As Long = 0           ' not used at this time.
Private Const OVER_LIST             As Long = 1           ' mouse cursor is over list portion of control.
Private Const OVER_UPBUTTON         As Long = 2           ' mouse cursor is over vertical scrollbar up button.
Private Const OVER_DOWNBUTTON       As Long = 3           ' mouse cursor is over vertical scrollbar down button.
Private Const OVER_VTRACKBAR        As Long = 4           ' mouse cursor is over vertical scrollbar trackbar.
Private Const OVER_VTHUMB           As Long = 5           ' mouse cursor is over vertical scrollbar thumb.
Private MouseLocation               As Long               ' set to one of the above constants.
Private MouseOverCheckBox           As Boolean            ' set to True if mouse is in checkbox display area.

Private DragFlag                    As Boolean            ' master 'is drag enabled?' flag.

' for scrollbar button display.
Private Const UPBUTTON              As Long = 1           ' display the vertical scrollbar up button.
Private Const DOWNBUTTON            As Long = 2           ' display the vertical scrollbar up button.

' for space bar selection of list items.
Private Const SPACEBAR              As Long = 32          ' space bar is chr(32).

'  the pixel range for the center of the vertical scrollbar
'  thumb as it goes up and down the scroller track.
Private Type vThumbRangeType
   Top                              As Long               ' top pixel position of middle of thumb.
   Bottom                           As Long               ' bottom pixel position of middle of thumb.
End Type
Private vThumbRange                 As vThumbRangeType

' thumb scroll tracking variables and constants.
Private MouseX                      As Single             ' global mouse X position variable.
Private MouseY                      As Single             ' global mouse Y position variable.
Private MouseDownYPos               As Single             ' mouse y position when mouse is clicked down.
Private Const SCROLL_TICKCOUNT      As Long = 50          ' scroll time delay interval.
Private Const INITIAL_SCROLL_DELAY  As Long = 400         ' delay before scrolling commences.
Private DraggingVThumb              As Boolean            ' flag indicating mouse is down on thumb..
Private ThumbScrolling              As Boolean            ' flag indicating thumb scrolling is now in progress.
Private MousePosInVThumb            As Single             ' distance in pixels from top of vertical thumb.
Private ScrollFlag                  As Boolean            ' mouse down on list, then moved above/below list.
Private Const SCROLL_LISTDOWN       As Long = 1           ' display range increment for scrolling list down.
Private Const SCROLL_LISTUP         As Long = -1          ' display range increment for scrolling list up.

' declares for virtual listbox background bitmap.
Private VirtualBackgroundDC         As Long               ' DC handle of the created Device Context
Private mMemoryBitmap               As Long               ' Handle of the created bitmap
Private mOrginalBitmap              As Long               ' Used in Destroy
'Property Variables:
Dim m_ScaleWidth As Single
Dim m_ScaleMode As Integer
Dim m_ScaleHeight As Single
Dim m_AutoRedraw As Boolean
'Default Property Values:
Const m_def_ScaleWidth = 0
Const m_def_ScaleMode = 0
Const m_def_ScaleHeight = 0
Const m_def_AutoRedraw = 0



Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

'*************************************************************************
'* processes the intercepted windows messages.  This subclass handler    *
'* MUST be the first Public routine in this file.  This includes         *
'* public properties also.  Other subclass routines at bottom of code.   *
'*************************************************************************

   Select Case uMsg

      Case WM_MOUSEWHEEL
         If HasFocus Then ' may want to take focus check out - regular vb textbox wheels w/o focus.
             Select Case wParam
             Case Is > False
                ProcessUpButton
             Case Else
                ProcessDownButton
             End Select
         End If

      Case WM_MOUSEMOVE
'        detect when mouse has entered the control.
         If m_Enabled And Not bInCtrl Then
            bInCtrl = True
            Call TrackMouseLeave(lng_hWnd)
            RaiseEvent MouseEnter
         End If

'     detect when mouse has left the control.
      Case WM_MOUSELEAVE
         bInCtrl = False
         RaiseEvent MouseLeave

'     detect when control has gained the focus.
      Case WM_SETFOCUS
         If m_Enabled Then
            HasFocus = True
            If m_Style = [Standard] Then
               DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
            Else
               DisplayListBoxItem ItemWithFocus, DrawAsSelected, FocusRectangleYes
            End If
            UserControl.Refresh
         End If

'     detect when control has lost the focus.
      Case WM_KILLFOCUS
         If m_Enabled Then
            HasFocus = False
            MouseAction = MOUSE_NOACTION
            ShiftKeyDown = False
            CtrlKeyDown = False
'           since listbox has lost the focus, display focused listbox item without the focus rectangle.
            If m_Style = [Standard] Then
               DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
            Else
               DisplayListBoxItem ItemWithFocus, DrawAsSelected, FocusRectangleNo
            End If
            UserControl.Refresh
         End If

   End Select

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Events >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_Initialize()

'*************************************************************************
'* the first event in the control's life cycle.                          *
'*************************************************************************

   Dim OS As OSVERSIONINFO

   ReDim ListArray(0)       ' the array tied to the .List property.
   ReDim ItemDataArray(0)   ' the array tied to the .ItemData property.
   ReDim SelectedArray(0)   ' the array tied to the .Selected property.
   ReDim ImageIndexArray(0) ' the array tied to the .ImageIndex property.

'  initialize property and internal variables.
   m_ListCount = 0
   LastSelectedItem = -1
   m_ListIndex = -1
   ItemWithFocus = 0
   ItemMouseIsIn = -1
   m_hBrush = 0             ' bitmap tiling pattern brush.
   RecalculateThumbHeight = True

'  get the operating system version for text drawing purposes.
   OS.dwOSVersionInfoSize = Len(OS)
   Call GetVersionEx(OS)
   mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub UserControl_Click()

'*************************************************************************
'* return a click event.  A VB listbox only returns a click event when   *
'* the mouse cursor is over the populated portion of list area; this be- *
'* havior is emulated here by utilizing the MouseOverIndex function.     *
'*************************************************************************

   If m_Enabled And Not RightClickFlag And MouseOverIndex(MouseY) <> -1 Then
      RaiseEvent Click
   End If

End Sub

Private Sub UserControl_DblClick()

'*************************************************************************
'* process a double-click event as normal, or as two rapid single click  *
'* events.  Why offer a choice?  You know how a standard VB button       *
'* interprets double-clicks as two rapid button presses?  I like that    *
'* quick responsiveness and incorporate it into controls that I don't    *
'* need doubleclick functionality for.  However, in a listbox I can see  *
'* where double-clicking a list item might be a desirable feature for    *
'* some, so I provided both options via the DblClickBehavior property.   *
'* Even if normal doubleclicks are permitted, they are still treated as  *
'* two rapid single-clicks if mouse is in vertical scrollbar area or the *
'* listbox is in CheckBox mode (to help emulate vb listbox).             *
'*************************************************************************

   If (m_DblClickBehavior = [Two Single Clicks]) Or (Not IsInList(MouseX, MouseY)) Or _
      (m_Style = [CheckBox] And IsInList(MouseX, MouseY)) Then
      If m_Enabled And Not RightClickFlag Then
'        originally I just sent control to the UserControl_MouseDown routine.  But I found that when
'        double-clicking (keeping second click held down), then drag-scrolling the list, the list would
'        not scroll.  This is because MouseMove would not fire when mouse cursor left the list in
'        this scenario. So I generate the actual mousedown event via this API to solve the problem.
         mouse_event MOUSEEVENTF_LEFTDOWN, MouseX, MouseY, 0, 0
      End If
   Else
      RaiseEvent DblClick
   End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'************************************************************************
'* processes a clicked item or vertical scrollbar component.            *
'************************************************************************

   If Not m_Enabled Then
      Exit Sub
   End If

   m_DragEnabled = False

   If Button = vbRightButton Then
      RightClickFlag = True
      RaiseEvent MouseDown(Button, Shift, X, Y)
      Exit Sub
   End If

   RightClickFlag = False

   Select Case MouseLocation

      Case OVER_UPBUTTON
'        process click of vertical scrollbar up button.
         MouseAction = MOUSE_DOWNED_IN_UPBUTTON
         ProcessUpButton

      Case OVER_DOWNBUTTON
'        process click of vertical scrollbar down button.
         MouseAction = MOUSE_DOWNED_IN_DOWNBUTTON
         ProcessDownButton

      Case OVER_VTRACKBAR
'        process a page up or page down based on mouse-click of vertical scrollbar in relation to thumb.
         If MouseCursorIsAboveThumb(Y) Then
            MouseAction = MOUSE_DOWNED_IN_UPPERTRACK
            ProcessPageUp
         Else
            MouseAction = MOUSE_DOWNED_IN_LOWERTRACK
            ProcessPageDown
         End If

      Case OVER_VTHUMB
'        initiate scolling of list via dragging of vertical scrollbar thumb.
         MouseAction = MOUSE_DOWNED_IN_THUMB
         DraggingVThumb = True
         MouseDownYPos = Y
         ProcessVThumbScroll

      Case OVER_LIST
'        make sure mouse pointer is in populated area of listbox before continuing.
         If MouseOverIndex(Y) <> -1 Then
            m_DragEnabled = DragFlag ' set to original user-selected property state.
'           process selection or deselection of a list item.
            MouseAction = MOUSE_DOWNED_IN_LIST
            If m_Style = [CheckBox] Then
               ProcessMouseDown_CheckBoxMode
            Else
               Select Case m_MultiSelect
                  Case vbMultiSelectNone
                     ProcessMouseDown_MultiSelectNone
                  Case vbMultiSelectSimple
                     ProcessMouseDown_MultiSelectSimple
                  Case vbMultiSelectExtended
                     ProcessMouseDown_MultiSelectExtended
               End Select
            End If
         End If

   End Select

   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* processes mouse up event.                                             *
'*************************************************************************

   If m_Enabled Then

      If Button = vbLeftButton Then

         Select Case MouseAction
            Case MOUSE_DOWNED_IN_UPBUTTON
'              if mouse was down on one of the scrollbar buttons, redisplay that button in its 'up' colors.
               MouseAction = MOUSE_NOACTION
               DisplayTrackBarButton UPBUTTON
            Case MOUSE_DOWNED_IN_DOWNBUTTON
               MouseAction = MOUSE_NOACTION
               DisplayTrackBarButton DOWNBUTTON
            Case MOUSE_DOWNED_IN_UPPERTRACK, MOUSE_DOWNED_IN_LOWERTRACK
'              if mouse was down on the scrollbar track, redisplay it so clicked portion is normal color.
               MouseAction = MOUSE_NOACTION
               DisplayVerticalScrollBar
            Case Else
               MouseAction = MOUSE_NOACTION
               m_DragEnabled = False
         End Select

         UserControl.Refresh
         DraggingVThumb = False
         RaiseEvent MouseUp(Button, Shift, X, Y)

      Else

'        raise the mouseup event for right mouse button.
         RaiseEvent MouseUp(Button, Shift, X, Y)

      End If

   End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* processes mouse movement based on .MultiSelect property.              *
'*************************************************************************

   Dim DidIt As Boolean    ' flag that lets routine know an action was performed.

'  set the global cursor coordinate variables.
   MouseX = X
   MouseY = Y

'  determine which component of the listbox the mouse is over (list, scrollbar, thumb, buttons, track).
   DetermineMouseLocation X, Y

'  check for and process possible drag scrolling.
   ProcessMouseDragScrolling DidIt
   If DidIt Then
      RaiseEvent MouseMove(Button, Shift, X, Y)
      Exit Sub
   End If

'  check for and process possible dragging of scrollbar out of range.
   ProcessMouseDragThumbOutOfRange DidIt
   If DidIt Then
      RaiseEvent MouseMove(Button, Shift, X, Y)
      Exit Sub
   End If

'  if user was scrolling with the thumb (started in MouseDown), make
'  sure scrolling can continue even if mouse has moved off the thumb.
   If MouseAction <> MOUSE_NOACTION And DraggingVThumb And Not ThumbScrolling Then
      ProcessVThumbScroll
      RaiseEvent MouseMove(Button, Shift, X, Y)
      Exit Sub
   End If

'  if we're dragging mouse over list portion of control, process
'  it.  Ignore this if drag and drop operation is enabled.
   If MouseAction = MOUSE_DOWNED_IN_LIST And MouseLocation = OVER_LIST And Not m_DragEnabled Then
      ProcessMouseMoveItemSelection
   End If

   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* allows user to navigate listbox via keyboard.                         *
'*************************************************************************

   If m_Enabled Then

'     determine Shift and Ctrl key status.
      ShiftKeyDown = (Shift And vbShiftMask) > 0
      If ShiftKeyDown Then
         ShiftDownStartItem = ItemWithFocus
      End If

      CtrlKeyDown = (Shift And vbCtrlMask) > 0

'     process the appropriate key.
      Select Case KeyCode
         Case vbKeyPageDown
            ProcessPageDownKey
         Case vbKeyPageUp
            ProcessPageUpKey
         Case vbKeyEnd
            ProcessEndKey
         Case vbKeyHome
            ProcessHomeKey
         Case vbKeyUp, vbKeyLeft
            ProcessUpArrowKey
         Case vbKeyDown, vbKeyRight
            ProcessDownArrowKey
         Case SPACEBAR
            ProcessSpaceBar
      End Select

      RaiseEvent KeyDown(KeyCode, Shift)

   End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   
'*************************************************************************
'* processes keypress event.                                             *
'*************************************************************************

   If m_Enabled Then
      RaiseEvent KeyPress(KeyAscii)
   End If

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* processes key up event.                                               *
'*************************************************************************

   If m_Enabled Then
'     determine shift and Ctrl key status.
      ShiftKeyDown = (Shift And vbShiftMask) > 0
      CtrlKeyDown = (Shift And vbCtrlMask) > 0
      RaiseEvent KeyUp(KeyCode, Shift)
   End If

End Sub

Private Sub UserControl_Resize()

'*************************************************************************
'* currently only used in design mode.                                   *
'*************************************************************************

'  when a new listbox is drawn onto the form in design mode, the event sequence is
'  Initialize-Show instead of Initialize-ReadProperties-Show.  This means we have to
'  calculate the gradient info from the defaults, as opposed to reading properties.
   CalculateGradients
   RedrawControl
   RaiseEvent Resize

End Sub

Private Sub UserControl_Terminate()

'*************************************************************************
'* restores memory used by listbox and stops subclassing.                *
'*************************************************************************

   On Error GoTo Catch

'  deallocate property arrays.
   Erase ListArray
   Erase ItemDataArray
   Erase SelectedArray
   Erase ImageIndexArray

'  destroy the virtual DC's used in background storage.
   DestroyVirtualDC
   DestroyPattern

'  halt subclassing.
   Call Subclass_StopAll

Catch:

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<< Mouse Processing Routines >>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ProcessMouseDown_CheckBoxMode()

'*************************************************************************
'* processes a mouse down in the list in Checkbox Style mode.            *
'*************************************************************************

   If LastSelectedItem = -1 Then ' need?
      LastSelectedItem = MouseOverIndex(MouseY)
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
   End If

   If Not MouseOverCheckBox Then

      If MouseOverIndex(MouseY) = LastSelectedItem Then
'        if the mouse is over the list (not a checkbox), and we are clicking on the
'        item that is already the focused item, then reverse its selection status and exit.
         SelectedArray(LastSelectedItem) = Not SelectedArray(LastSelectedItem)
         If SelectedArray(LastSelectedItem) Then
            m_SelCount = m_SelCount + 1
         Else
            m_SelCount = m_SelCount - 1
         End If
         DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
         UserControl.Refresh
         Exit Sub
      Else
'        if the mouse is over the list (not a checkbox), and we are NOT clicking on the
'        item that is already the focused item, then set the focus and selection gradient
'        to the new item and exit.  No selected status changes are made.
         DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo
         LastSelectedItem = MouseOverIndex(MouseY)
         ItemWithFocus = LastSelectedItem
         m_ListIndex = ItemWithFocus
         ItemMouseIsIn = LastSelectedItem
'        redisplay the newly selected item.
         DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
         UserControl.Refresh
         Exit Sub
      End If

   Else

'     if the mouse is clicked in a list item's checkbox, that item's selection status
'     is immediately reversed, and the selection gradient moves to that list item.

'     display the previous list item without selection bar.
      DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo
      LastSelectedItem = MouseOverIndex(MouseY)
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
      ItemMouseIsIn = LastSelectedItem
      SelectedArray(LastSelectedItem) = Not SelectedArray(LastSelectedItem)
      If SelectedArray(LastSelectedItem) Then
         m_SelCount = m_SelCount + 1
      Else
         m_SelCount = m_SelCount - 1
      End If

'     redisplay the newly selected item.
      DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
      UserControl.Refresh

   End If

End Sub

Private Sub ProcessMouseDown_MultiSelectNone()

'*************************************************************************
'* processes a mouse down in the list in MultiSelect None mode.          *
'*************************************************************************

'  repaint previously selected list item without selection gradient / focus rectangle.
   ClearPreviousSelection KeepSelectionNo

'  make the newly clicked item the last selected item.
   LastSelectedItem = MouseOverIndex(MouseY)
   ProcessSelectedItem

End Sub

Private Sub ProcessMouseDown_MultiSelectSimple()

'*************************************************************************
'* processes a mouse down in the list in MultiSelect Simple mode.        *
'*************************************************************************

'  repaint previously selected list item, keeping selection status but not focus rectangle.
   ClearPreviousSelection KeepSelectionAsIs

   LastSelectedItem = MouseOverIndex(MouseY) ' even if we're deselecting with the mouse click?
   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus

'  reverse the selection status of the list item that was clicked.
   If SelectedArray(LastSelectedItem) Then
      SelectedArray(LastSelectedItem) = False
      DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleYes
      m_SelCount = m_SelCount - 1
   Else
      SelectedArray(LastSelectedItem) = True
      DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
      m_SelCount = m_SelCount + 1
   End If

   UserControl.Refresh

End Sub

Private Sub ProcessMouseDown_MultiSelectExtended()

'*************************************************************************
'* processes a mouse down in the list in MultiSelect None mode.          *
'*************************************************************************

'  make the newly clicked item the last selected item.
   LastSelectedItem = MouseOverIndex(MouseY)

'  this item must also have the focus rectangle.
   ItemWithFocus = LastSelectedItem

'  in Extended mode, the .ListIndex property is always the item with the focus.
   m_ListIndex = ItemWithFocus

'  if the Ctrl or Shift keys are not down, set the entire Selected array to False.
   If (Not CtrlKeyDown) And (Not ShiftKeyDown) Then
      SetSelectedArrayRange 0, m_ListCount - 1, False
   End If

'  in MultiSelect Extended mode, a mouse down just selects one item unless shift is pressed.
   FirstExtendedSelection = LastSelectedItem
   LastExtendedSelection = LastSelectedItem

'  make sure item is selected.  If Ctrl key is pressed, flip the selection status.
   If (Not CtrlKeyDown) Then
      SelectedArray(LastSelectedItem) = True
   Else
      SelectedArray(LastSelectedItem) = Not SelectedArray(LastSelectedItem)
   End If

'  if Shift is down, all items from the item that had the focus when
'  shift was pressed to the new item that has the focus are selected.
   If ShiftKeyDown Then
      SetSelectedArrayRange ShiftDownStartItem, ItemWithFocus, True
      CalculateSelCount
   End If

'  display whole list instead of just item - other displayed items' selection status may have changed.
   DisplayList

'  used in MouseMove to detect whether mouse is still in the same list item as it is now.
   ItemMouseIsIn = LastSelectedItem

'  since mouse is being clicked down, SelCount is always 1 if Ctrl or Shift keys not pressed.
   If (Not CtrlKeyDown) And (Not ShiftKeyDown) Then
      m_SelCount = 1
   Else
      If Not ShiftKeyDown Then
         If SelectedArray(LastSelectedItem) And Not ShiftKeyDown Then
            m_SelCount = m_SelCount + 1
         Else
            m_SelCount = m_SelCount - 1
         End If
      End If
   End If

End Sub

Private Sub ProcessDownButton()

'*************************************************************************
'* shifts displayed list items down on scrollbar down arrow button click.*
'*************************************************************************

'  only do this if last list item is not already displayed.
   If Not InDisplayedItemRange(m_ListCount - 1) Then
      DisplayRange.FirstListItem = DisplayRange.FirstListItem + 1
      DisplayRange.LastListItem = DisplayRange.LastListItem + 1
      DisplayList
'     check for and process possible continuous scroll (i.e. mouse button held down).
      ProcessContinuousScroll SCROLL_LISTDOWN
   End If

End Sub

Private Sub ProcessUpButton()

'*************************************************************************
'* shifts displayed list items up on scrollbar up arrow button click.    *
'*************************************************************************

'  only do this if first list item is not already displayed.
   If Not InDisplayedItemRange(0) Then
      DisplayRange.FirstListItem = DisplayRange.FirstListItem - 1
      DisplayRange.LastListItem = DisplayRange.LastListItem - 1
      DisplayList
'     check for and process possible continuous scroll (i.e. mouse arrow button held down).
      ProcessContinuousScroll SCROLL_LISTUP
   End If

End Sub

Private Sub ProcessPageUp()

'*************************************************************************
'* shifts displayed list items up one page when mouse is clicked above   *
'* vertical scroll thumb.                                                *
'*************************************************************************

'  only perform a page up if first page is not already displayed.
   If Not InDisplayedItemRange(0) Then

'     adjust the displayed item range.
      DisplayRange.FirstListItem = DisplayRange.FirstListItem - MaxDisplayItems + 1
      If DisplayRange.FirstListItem < 0 Then
         DisplayRange.FirstListItem = 0
      End If
      If DisplayRange.FirstListItem + MaxDisplayItems - 1 > m_ListCount Then
         DisplayRange.LastListItem = m_ListCount - 1
      Else
         DisplayRange.LastListItem = DisplayRange.FirstListItem + MaxDisplayItems - 1
      End If

      DisplayList

'     check for and process possible continuous scroll (i.e. mouse button held down).
      ProcessContinuousScroll -MaxDisplayItems

   End If

End Sub

Private Sub ProcessPageDown()

'*************************************************************************
'* shifts displayed list items down one page when mouse is clicked below *
'* vertical scroll thumb.                                                *
'*************************************************************************

'  only perform a page down if last page is not already being displayed.
   If Not InDisplayedItemRange(m_ListCount - 1) Then

'     adjust the displayed item range.
      If DisplayRange.LastListItem + MaxDisplayItems - 1 <= m_ListCount - 1 Then
         DisplayRange.FirstListItem = DisplayRange.LastListItem
      Else
         DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems + 1
      End If
      DisplayRange.LastListItem = DisplayRange.FirstListItem + MaxDisplayItems
      If DisplayRange.LastListItem > m_ListCount - 1 Then
         DisplayRange.LastListItem = m_ListCount - 1
      End If

      DisplayList

'     check for and process possible continuous scroll (i.e. mouse button held down).
      ProcessContinuousScroll MaxDisplayItems

   End If

End Sub

Private Sub ClearPreviousSelection(SelectionSaveIndicator As Boolean)

'*************************************************************************
'* clears the listbox of last selected item's gradient (if the Style is  *
'* set to None or Extended) and erases the focus rectangle in whatever   *
'* list item possesses it.  Only happens if list item(s) are displayed.  *
'*************************************************************************

'  repaints the selected item without the focus rectangle (saving gradient selection
'  highlight), or redisplays as unselected without the focus rectangle
'  (depends on SelectionSaveStatus parameter).
   If LastSelectedItem <> -1 And InDisplayedItemRange(LastSelectedItem) Then
      Select Case SelectionSaveIndicator
        Case KeepSelectionAsIs
           DisplayListBoxItem LastSelectedItem, SelectedArray(LastSelectedItem), FocusRectangleNo
        Case KeepSelectionNo
           DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo
      End Select
   End If

'  make sure the item with the focus rectangle is 'de-rectangled'.
'  Selected or deselected appearance of item is unchanged.
   If ItemWithFocus <> LastSelectedItem Then
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
   End If

End Sub

Private Function InDisplayedItemRange(Index As Long) As Boolean

'*************************************************************************
'* returns whether a given list item index is in the displayed range.    *
'*************************************************************************

   If Index >= DisplayRange.FirstListItem And Index <= DisplayRange.LastListItem Then
      InDisplayedItemRange = True
   End If

End Function

Private Sub ProcessMouseDragScrolling(DidIt As Boolean)

'*************************************************************************
'* this allows user to scroll through the list by clicking within the    *
'* list area and then dragging the mouse above or below the listbox,     *
'* like a regular vb listbox.                                            *
'*************************************************************************

   If ScrollFlag And MouseY >= 0 And MouseY <= ScaleHeight Then
      ScrollFlag = False
      DidIt = False
   ElseIf MouseAction <> MOUSE_NOACTION And Not DraggingVThumb And Not ScrollFlag Then
      If MouseY > ScaleHeight Then
         ScrollFlag = True
         ProcessContinuousScroll SCROLL_LISTDOWN
         ScrollFlag = False
         DidIt = True
      ElseIf MouseY < 0 Then
         ScrollFlag = True
         ProcessContinuousScroll SCROLL_LISTUP
         ScrollFlag = False
         DidIt = True
      End If
   End If

End Sub

Private Sub ProcessMouseMoveItemSelection()

'*************************************************************************
'* controls selection of items by mouse drag in all listbox states.      *
'*************************************************************************

'  first, make sure the mouse hasn't just been moved within the same list
'  item it was in the last time the mouse was moved.  This avoids unnecessary
'  processing and graphics redraws.
   If ItemMouseIsIn = MouseOverIndex(MouseY) Then
      Exit Sub
   Else
      ItemMouseIsIn = MouseOverIndex(MouseY)
   End If

   If m_Style = [CheckBox] Then
      ProcessMouseMoveItemSelection_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessMouseMoveItemSelection_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessMouseMoveItemSelection_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessMouseMoveItemSelection_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessMouseMoveItemSelection_CheckBoxMode()

'*************************************************************************
'* processes selecting items by mouse drag in CheckBox mode.             *
'*************************************************************************

'  display the previously selected item as unselected, with no focus rectangle.
   DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo

'  the newly moved-over item is now the last item with the selection bar.
   LastSelectedItem = MouseOverIndex(MouseY)

'  if mouse is moved really fast past the end of a less than 1 page list the last item is
'  not always highlighted because the mousemove didn't fire.  This helps correct that.
   If LastSelectedItem = -1 Or (m_ListCount < MaxDisplayItems And MouseY > ScaleHeight) Then
      LastSelectedItem = m_ListCount - 1
   End If

   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus

'  display the item with the selection bar.
   DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessMouseMoveItemSelection_MultiSelectNone()

'*************************************************************************
'* processes selecting items by mouse drag in MultiSelect None mode.     *
'*************************************************************************

'  display the previously selected item as unselected, with no focus rectangle.
   DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo

'  the newly moved-over item is now the last item selected.
   LastSelectedItem = MouseOverIndex(MouseY)

'  if mouse is moved really fast past the end of a less than 1 page list the last item is
'  not always highlighted because the mousemove didn't fire.  This helps correct that.
   If LastSelectedItem = -1 Or (m_ListCount < MaxDisplayItems And MouseY > ScaleHeight) Then
      LastSelectedItem = m_ListCount - 1
   End If

   ProcessSelectedItem

End Sub

Private Sub ProcessMouseMoveItemSelection_MultiSelectSimple()

'*************************************************************************
'* processes selecting items by mouse drag in MultiSelect Simple mode.   *
'*************************************************************************

'  get rid of the focus rectangle in the previously focused item, keeping selection status.
   DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo

'  set the moused-over item as the new item with focus.
   ItemWithFocus = MouseOverIndex(MouseY)

'  if mouse is moved really fast past the end of a less than 1 page list the last item is
'  not always highlighted because the mousemove didn't fire.  This helps correct that.
   If ItemWithFocus = -1 Or (m_ListCount < MaxDisplayItems And MouseY > ScaleHeight) Then
      ItemWithFocus = m_ListCount - 1
   End If

'  in MultiSelect Simple mode, the .ListIndex property is always the item with the focus.
   m_ListIndex = ItemWithFocus

'  paint the focus rectangle over the moused-over item, keeping selection status as-is.
   DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessMouseMoveItemSelection_MultiSelectExtended()

'*************************************************************************
'* processes selecting items by mouse drag in MultiSelect Extended mode. *
'* when mouse is moved in MultiSelect Extended mode then the item's      *
'* selection status is reversed (unless it's the originally selected     *
'* item, in which case it stays selected.)                               *
'*************************************************************************

'  get rid of the focus rectangle in the previously focused item, keeping selection status.
   DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
   UserControl.Refresh

'  if Shift or Ctrl not pressed, clear all selected items so that redisplay of list is handled correctly.
   If (Not ShiftKeyDown) And (Not CtrlKeyDown) Then
      SetSelectedArrayRange 0, m_ListCount - 1, False
   End If

'  set the moused-over item as the new item with focus.
   LastSelectedItem = MouseOverIndex(MouseY)

'  if mouse is moved really fast past the end of a less than 1 page list the last item is
'  not always highlighted because the mousemove didn't fire.  This helps correct that.
   If LastSelectedItem = -1 Or (m_ListCount < MaxDisplayItems And MouseY > ScaleHeight) Then
      LastSelectedItem = m_ListCount - 1
   End If

'  focus rectangle and .ListIndex property are always set to lat selected item in this mode.
   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus

'  set selection status of all items from first to last to True.
   LastExtendedSelection = ItemWithFocus

'  set the selected range.
   SetSelectedArrayRange FirstExtendedSelection, LastExtendedSelection, True

'  determine the number of selected items.
   If (Not ShiftKeyDown) And (Not CtrlKeyDown) Then
      m_SelCount = Abs(LastExtendedSelection - FirstExtendedSelection) + 1
   Else
      CalculateSelCount
   End If

   DisplayList

End Sub

Private Sub DetermineMouseLocation(X As Single, Y As Single)

'*************************************************************************
'* sets the MouseLocation variable based on which listbox component the  *
'* mouse cursor is at the time of the call to this routine.              *
'*************************************************************************

   If IsInList(X, Y) Then
      MouseLocation = OVER_LIST
   ElseIf IsInVerticalThumb(X, Y) Then
      MouseLocation = OVER_VTHUMB
   ElseIf IsInVerticalTrackbar(X, Y) Then
      MouseLocation = OVER_VTRACKBAR
   ElseIf IsInUpButton(X, Y) Then
      MouseLocation = OVER_UPBUTTON
   ElseIf IsInDownButton(X, Y) Then
      MouseLocation = OVER_DOWNBUTTON
   Else
      MouseLocation = OVER_BORDER
   End If

'  check to see if mouse is over a checkbox if in CheckBox mode.
   If m_Style = [CheckBox] And MouseLocation = OVER_LIST Then
      If X >= m_BorderWidth + 3 And X <= m_BorderWidth + 18 Then
         MouseOverCheckBox = True
      Else
         MouseOverCheckBox = False
      End If
   End If

End Sub

Private Function IsInList(XPos As Single, YPos As Single) As Boolean

'*************************************************************************
'* returns True if mouse cursor is in list display portion of control.   *
'*************************************************************************

   Dim RightListBorder As Long   ' right edge of list area.

'  account for possible scrollbar being displayed.
   If VerticalScrollBarActive Then
      RightListBorder = ScaleWidth - m_BorderWidth - ScrollBarButtonWidth - 1
   Else
      RightListBorder = ScaleWidth - m_BorderWidth - 1
   End If

   If XPos >= m_BorderWidth And _
      XPos <= RightListBorder And _
      YPos >= m_BorderWidth And _
      YPos <= ScaleHeight - m_BorderWidth Then
         IsInList = True
   End If

End Function

Private Function IsInVerticalScrollbar(XPos As Single) As Boolean
      
'*************************************************************************
'* returns True if mouse cursor is in any part of the vertical scrollbar.*
'*************************************************************************

   If VerticalScrollBarActive Then
      If XPos >= vScrollBarLocation.ScrollTrackLocation.Left And _
         XPos <= vScrollBarLocation.ScrollTrackLocation.Right Then
            IsInVerticalScrollbar = True
      End If
   End If

End Function

Private Function IsInVerticalThumb(XPos As Single, YPos As Single) As Boolean

'*************************************************************************
'* returns True if mouse cursor is in vertical scrollbar thumb.          *
'*************************************************************************

   If VerticalScrollBarActive Then
      If YPos >= vScrollBarLocation.ScrollThumbLocation.Top And _
         YPos <= vScrollBarLocation.ScrollThumbLocation.Bottom And _
         XPos >= vScrollBarLocation.ScrollThumbLocation.Left And _
         XPos <= vScrollBarLocation.ScrollThumbLocation.Right Then
            IsInVerticalThumb = True
      End If
   End If

End Function

Private Function IsInVerticalTrackbar(XPos As Single, YPos As Single) As Boolean
      
'*************************************************************************
'* returns True if mouse cursor is in vertical scrollbar trackbar.       *
'*************************************************************************

   If IsInVerticalScrollbar(XPos) Then
      If YPos >= vScrollBarLocation.ScrollTrackLocation.Top And _
         YPos <= vScrollBarLocation.ScrollTrackLocation.Bottom And _
         Not IsInVerticalThumb(XPos, YPos) Then
            IsInVerticalTrackbar = True
      End If
   End If

End Function

Private Function IsInUpButton(XPos As Single, YPos As Single) As Boolean
      
'*************************************************************************
'* returns True if mouse cursor is in vertical scrollbar up button.      *
'*************************************************************************

   If IsInVerticalScrollbar(XPos) Then
      If YPos >= vScrollBarLocation.UpButtonLocation.Top And _
         YPos <= vScrollBarLocation.UpButtonLocation.Bottom Then
            IsInUpButton = True
      End If
   End If

End Function

Private Function IsInDownButton(XPos As Single, YPos As Single) As Boolean
      
'*************************************************************************
'* returns True if mouse cursor is in vertical scrollbar down button.    *
'*************************************************************************

   If IsInVerticalScrollbar(XPos) Then
      If YPos >= vScrollBarLocation.DownButtonLocation.Top And _
         YPos <= vScrollBarLocation.DownButtonLocation.Bottom Then
            IsInDownButton = True
      End If
   End If

End Function

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Keyboard Processing >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ProcessPageUpKey()

'*************************************************************************
'* processes page up key for all listbox states.                         *
'* page down ALWAYS starts from the item with the focus rectangle, even  *
'* if that item is not currently in the display range.  After the page   *
'* up the item that formerly had focus rect will be last displayed entry *
'* (unless said entry was less than "MaxDisplayItems" entries below top  *
'* item in list)                                                         *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessPageUpKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessPageUpKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessPageUpKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessPageUpKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessPageUpKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, PgDn moves selection bar down one page.  Selection  *
'* status of list items is unchanged.                                    *
'*************************************************************************

   CalculatePageUpDisplayRange

   LastSelectedItem = DisplayRange.FirstListItem
   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus
   DisplayList

   End Sub

Private Sub ProcessPageUpKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, PgUp moves selection bar up one page.       *
'*************************************************************************

   CalculatePageUpDisplayRange

   LastSelectedItem = DisplayRange.FirstListItem
   ProcessSelectedItem
   DisplayList

End Sub

Private Sub ProcessPageUpKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, PgUp moves focus rectangle up one page.   *
'*************************************************************************

   CalculatePageUpDisplayRange

   ItemWithFocus = DisplayRange.FirstListItem
   m_ListIndex = ItemWithFocus
   DisplayList

End Sub

Private Sub ProcessPageUpKey_MultiSelectExtended()

'*************************************************************************
'* in MultiSelect Extended mode, PgUp acts just like MultiSelect None    *
'* mode if the shift key is not pressed.  If shift is down, all items    *
'* from the item that had the focus when shift was pressed to the new    *
'* item that has the focus are selected.                                 *
'*************************************************************************

   If Not ShiftKeyDown Then
'     reinitialize the selected array to all False.
      SetSelectedArrayRange 0, m_ListCount - 1, False
   End If

   CalculatePageUpDisplayRange

   LastSelectedItem = DisplayRange.FirstListItem
   ProcessSelected_MultiSelectExtended False ' don't calculate m_SelCount

'  if shift is down, all items from the item that had the focus when
'  shift was pressed to the new item that has the focus are selected.
   If ShiftKeyDown Then
      SetSelectedArrayRange ShiftDownStartItem, ItemWithFocus, True
      CalculateSelCount
   Else
      m_SelCount = 1
   End If

   DisplayList

End Sub

Private Sub CalculatePageUpDisplayRange()

'*************************************************************************
'* determines the first and last list items to display on PgUp keypress. *
'*************************************************************************

   If ItemWithFocus - MaxDisplayItems + 1 >= 0 Then
      DisplayRange.LastListItem = ItemWithFocus
      DisplayRange.FirstListItem = DisplayRange.LastListItem - MaxDisplayItems + 1
   Else
      DisplayRange.FirstListItem = 0
      If m_ListCount >= MaxDisplayItems Then
         DisplayRange.LastListItem = MaxDisplayItems - 1
      Else
         DisplayRange.LastListItem = m_ListCount - 1
      End If
   End If

End Sub

Private Sub ProcessPageDownKey()

'*************************************************************************
'* processes page down key for all listbox states.                       *
'* page down ALWAYS starts from the item with the focus rectangle, even  *
'* if that item is not currently in the display range.  After the page   *
'* down the item that formerly had focus rect will be first displayed    *
'* entry (unless said entry was below first item on last page).          *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Or m_ListCount = 0 Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessPageDownKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessPageDownKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessPageDownKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessPageDownKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessPageDownKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, PgDn moves selection bar down one page.  Selection  *
'* status of list items is unchanged.                                    *
'*************************************************************************

   CalculatePageDownDisplayRange

   LastSelectedItem = DisplayRange.LastListItem
   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus
   DisplayList

End Sub

Private Sub ProcessPageDownKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, PgDn moves selection bar down one page.     *
'*************************************************************************

   CalculatePageDownDisplayRange

   LastSelectedItem = DisplayRange.LastListItem
   ProcessSelectedItem
   DisplayList

End Sub

Private Sub ProcessPageDownKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, PgDn moves focus rectangle down one page. *
'*************************************************************************

   CalculatePageDownDisplayRange

   ItemWithFocus = DisplayRange.LastListItem
   m_ListIndex = ItemWithFocus
   DisplayList

End Sub

Private Sub ProcessPageDownKey_MultiSelectExtended()

'*************************************************************************
'* in MultiSelect Extended mode, PgDn acts just like MultiSelect None    *
'* mode if the shift key is not pressed.  If shift is down, all items    *
'* from the item that had the focus when shift was pressed to the new    *
'* item that has the focus are selected.                                 *
'*************************************************************************

   If Not ShiftKeyDown Then
'     reinitialize the selected array to all False.
      SetSelectedArrayRange 0, m_ListCount - 1, False
   End If

   CalculatePageDownDisplayRange

   LastSelectedItem = DisplayRange.LastListItem
   ProcessSelected_MultiSelectExtended False ' don't calculate m_SelCount

'  if shift is down, all items from the item that had the focus when
'  shift was pressed to the new item that has the focus are selected.
   If ShiftKeyDown Then
      SetSelectedArrayRange ShiftDownStartItem, ItemWithFocus, True
      CalculateSelCount
   Else
      m_SelCount = 1
   End If

   DisplayList

End Sub

Private Sub CalculatePageDownDisplayRange()

'*************************************************************************
'* determines the first and last list items to display on PgDn keypress. *
'*************************************************************************

   If ItemWithFocus + MaxDisplayItems - 1 < m_ListCount Then
      DisplayRange.FirstListItem = ItemWithFocus
      DisplayRange.LastListItem = DisplayRange.FirstListItem + MaxDisplayItems - 1
   Else
      DisplayRange.LastListItem = m_ListCount - 1
      If m_ListCount >= MaxDisplayItems Then
         DisplayRange.FirstListItem = DisplayRange.LastListItem - MaxDisplayItems + 1
      Else
         DisplayRange.FirstListItem = 0
      End If
   End If

End Sub

Private Sub ProcessEndKey()

'*************************************************************************
'* processes end key for all listbox states.                             *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessEndKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessEndKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessEndKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessEndKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessEndKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, selection bar is moved to bottom of list.  The sel- *
'* ected status of listbox items is unchanged.                           *
'*************************************************************************

   If LastSelectedItem <> m_ListCount - 1 Then
      LastSelectedItem = m_ListCount - 1
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
      DetermineLastPageDisplayRange
      DisplayList
   End If

End Sub

Private Sub ProcessEndKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, selection bar is moved to end of list.      *
'*************************************************************************

   If LastSelectedItem <> m_ListCount - 1 Then
      ClearPreviousSelection KeepSelectionNo
      DetermineLastPageDisplayRange
      LastSelectedItem = m_ListCount - 1
      ProcessSelectedItem
      DisplayList
   End If

End Sub

Private Sub ProcessEndKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, focus rectangle is moved to end of list.  *
'*************************************************************************

   If ItemWithFocus <> m_ListCount - 1 Then
'     repaint previously focused list item, maintaining selection status.
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
      DetermineLastPageDisplayRange
      LastSelectedItem = m_ListCount - 1
      ItemWithFocus = m_ListCount - 1
      m_ListIndex = ItemWithFocus
      DisplayList
   End If

End Sub

Private Sub ProcessEndKey_MultiSelectExtended()

'*************************************************************************
'* processes End/Shift-End key in MultiSelect Extended mode.           *
'*************************************************************************

'  make sure any previously selected items are de-selected if shift key is not being pressed.
   If Not ShiftKeyDown Then
      SetSelectedArrayRange 0, m_ListCount - 1, False
      m_SelCount = 1
   Else
'     if the shift key is down, all items from ItemWithFocus to the bottom of the list are selected.
      SetSelectedArrayRange ItemWithFocus, m_ListCount - 1, True
'     for special cases like Extended mode Shift-Home/Shift-End we need to brute-force it.
      CalculateSelCount
   End If

   If LastSelectedItem <> m_ListCount - 1 Then
      ClearPreviousSelection KeepSelectionNo
   End If

   DetermineLastPageDisplayRange
   LastSelectedItem = m_ListCount - 1
   ProcessSelected_MultiSelectExtended False
   DisplayList

End Sub

Private Sub DetermineLastPageDisplayRange()

'*************************************************************************
'* calculates range of items to display at bottom of list.               *
'*************************************************************************

   DisplayRange.LastListItem = m_ListCount - 1
   If m_ListCount >= MaxDisplayItems Then
      DisplayRange.FirstListItem = DisplayRange.LastListItem - MaxDisplayItems + 1
   Else
      DisplayRange.FirstListItem = 0
   End If

End Sub

Private Sub ProcessHomeKey()

'*************************************************************************
'* processes home key for all listbox states.                            *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessHomeKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessHomeKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessHomeKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessHomeKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessHomeKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, selection bar is moved to top of list.  The selec-  *
'* ted status of listbox items is unchanged.                             *
'*************************************************************************

   If LastSelectedItem <> 0 Then
      LastSelectedItem = 0
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
      DetermineFirstPageDisplayRange
      DisplayList
   End If

End Sub

Private Sub ProcessHomeKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, selection bar is moved to top of list.      *
'*************************************************************************

   If LastSelectedItem <> 0 Then
      ClearPreviousSelection KeepSelectionNo
      DetermineFirstPageDisplayRange
      LastSelectedItem = 0
      ProcessSelectedItem
      DisplayList
   End If

End Sub

Private Sub ProcessHomeKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, focus rectangle is moved to top of list.  *
'*************************************************************************

   If ItemWithFocus <> 0 Then
'     repaint previously focused list item, maintaining selection status.
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
      DetermineFirstPageDisplayRange
      ItemWithFocus = 0
      m_ListIndex = ItemWithFocus
      DisplayList
   End If

End Sub

Private Sub ProcessHomeKey_MultiSelectExtended()

'*************************************************************************
'* processes Home/Shift-Home key in MultiSelect Extended mode.           *
'*************************************************************************

'  make sure any previously selected items are de-selected if shift key is not being pressed.
   If Not ShiftKeyDown Then
      SetSelectedArrayRange 0, m_ListCount - 1, False
      m_SelCount = 1
   Else
'     if the shift key is down, all items from ItemWithFocus to the top of the list are selected.
      SetSelectedArrayRange 0, ItemWithFocus, True
'     for special cases like Extended mode Shift-Home/Shift-End we need to brute-force it.
      CalculateSelCount
   End If

   If LastSelectedItem <> 0 Then
      ClearPreviousSelection KeepSelectionNo
   End If

   LastSelectedItem = 0
   ProcessSelected_MultiSelectExtended False

   DetermineFirstPageDisplayRange
   DisplayList

End Sub

Private Sub DetermineFirstPageDisplayRange()

'*************************************************************************
'* calculates range of items to display at top of list.                  *
'*************************************************************************

   DisplayRange.FirstListItem = 0
   If m_ListCount < MaxDisplayItems Then
      DisplayRange.LastListItem = m_ListCount - 1
   Else
      DisplayRange.LastListItem = MaxDisplayItems - 1
   End If

End Sub

Private Sub ProcessSpaceBar()

'*************************************************************************
'* processes list item selection via spacebar for all listbox states.    *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessSpaceBar_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessSpaceBar_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessSpaceBar_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessSpaceBar_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessSpaceBar_CheckBoxMode()

'*************************************************************************
'* toggles selection status of focused item in CheckBox mode.            *
'*************************************************************************

   SelectedArray(LastSelectedItem) = Not SelectedArray(LastSelectedItem)
   If SelectedArray(LastSelectedItem) Then
      m_SelCount = m_SelCount + 1
   Else
      m_SelCount = m_SelCount - 1
   End If

'  redisplay the item to reflect current selection status.
   DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessSpaceBar_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, spacebar selects item, with no toggle.      *
'*************************************************************************

   LastSelectedItem = ItemWithFocus
   ProcessSelectedItem

End Sub

Private Sub ProcessSpaceBar_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, the space bar toggles selection of the    *
'* item with focus rectangle.                                            *
'*************************************************************************

   If SelectedArray(ItemWithFocus) Then
      SelectedArray(ItemWithFocus) = False
      m_SelCount = m_SelCount - 1
      m_ListIndex = ItemWithFocus
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
      UserControl.Refresh
   Else
      LastSelectedItem = ItemWithFocus
      ProcessSelectedItem
   End If

End Sub

Private Sub ProcessSpaceBar_MultiSelectExtended()

'*************************************************************************
'* in MultiSelect Extended mode, space bar selects item contained in the *
'* focus rectangle, deselecting all other selected items if the Shift    *
'* key is not being pressed.                                             *
'*************************************************************************

   Dim NumPreviouslySelected As Long

   NumPreviouslySelected = m_SelCount

'  make sure any previously selected items are de-selected if shift key is not being pressed.
   If Not ShiftKeyDown Then
      SetSelectedArrayRange 0, m_ListCount - 1, False
      m_SelCount = 0
   End If

   LastSelectedItem = ItemWithFocus
   ProcessSelectedItem
   AdjustDisplayRange

   If NumPreviouslySelected > 1 Then
      DisplayList ' instead of just the list item; other items may have to be visually deselected.
   Else
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
      UserControl.Refresh
   End If

End Sub

Private Sub ProcessUpArrowKey()

'*************************************************************************
'* processes up arrow key for all listbox states.                        *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessUpArrowKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessUpArrowKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessUpArrowKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessUpArrowKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessUpArrowKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, up and down arrow keys only move selection gradient *
'* bar.  The selection status of the item is unchanged.                  *
'*************************************************************************

   If LastSelectedItem > 0 Then
      If LastSelectedItem <> DisplayRange.FirstListItem Then ' prevents flicker
         DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo
      End If
      LastSelectedItem = LastSelectedItem - 1
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
      AdjustDisplayRange
      DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
      UserControl.Refresh
   End If

End Sub

Private Sub ProcessUpArrowKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, up arrow key selects the item with the      *
'* focus rectangle if it's not already selected.  Otherwise, it moves    *
'* the selection bar up one list item.                                   *
'*************************************************************************

   If m_SelCount = 0 Then
      LastSelectedItem = ItemWithFocus
      ProcessSelectedItem
      AdjustDisplayRange
   Else
      If LastSelectedItem > 0 Then
'        repaint previously selected list item without selection gradient / focus rectangle.
'        to prevent flicker, don't repaint if first item in display is focused.
         If ItemWithFocus <> DisplayRange.FirstListItem Then
            ClearPreviousSelection KeepSelectionNo
         End If
         LastSelectedItem = LastSelectedItem - 1
         ProcessSelectedItem
         AdjustDisplayRange
      End If
   End If

End Sub

Private Sub ProcessUpArrowKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, the arrow keys move just the focus rect-  *
'* angle.  Selection status of each affected list item is unchanged.     *
'*************************************************************************

   If ItemWithFocus > 0 Then

'     'defocus' previously focused list item, maintaining selection status.
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo

      ItemWithFocus = ItemWithFocus - 1
      m_ListIndex = ItemWithFocus

      AdjustDisplayRange
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
      UserControl.Refresh

   End If

End Sub

Private Sub ProcessUpArrowKey_MultiSelectExtended()

'*************************************************************************
'* controls up arrow processing in MultiSelect Extended mode.            *
'*************************************************************************

   Dim NumPreviouslySelected As Long

   NumPreviouslySelected = m_SelCount

   If Not ShiftKeyDown Then
'     make sure any previously selected items are de-selected if shift key is not being pressed.
      SetSelectedArrayRange 0, m_ListCount - 1, False
   End If

'  repaint previously focused list item without focus rectangle, maintaining selection status.
'  to prevent flicker, don't repaint if item with focus is first in displayed range.
   If ItemWithFocus <> DisplayRange.FirstListItem Then
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
   End If

   If LastSelectedItem > 0 Then
      LastSelectedItem = LastSelectedItem - 1
   End If

   ProcessSelected_MultiSelectExtended False

   If ShiftKeyDown Then
      CalculateSelCount
   Else
      m_SelCount = 1
   End If

   AdjustDisplayRange
   If NumPreviouslySelected > 1 Then
      DisplayList ' instead of just the list item; other items may have to be visually deselected
   Else
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
      UserControl.Refresh
   End If

End Sub

Private Sub ProcessDownArrowKey()

'*************************************************************************
'* processes down arrow key for all listbox states.                      *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessDownArrowKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessDownArrowKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessDownArrowKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessDownArrowKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessDownArrowKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, up and down arrow keys only move selection gradient *
'* bar.  The selection status of the item is unchanged.                  *
'*************************************************************************

   If LastSelectedItem < m_ListCount - 1 Then
      If LastSelectedItem <> DisplayRange.LastListItem Then ' prevents flicker
         DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo
      End If
      LastSelectedItem = LastSelectedItem + 1
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
      AdjustDisplayRange
      DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
      UserControl.Refresh
   End If

End Sub

Private Sub ProcessDownArrowKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, down arrow key selects the item with the    *
'* focus rectangle if it's not already selected.  Otherwise, it moves    *
'* the selection bar down one list item.                                 *
'*************************************************************************

   If m_SelCount = 0 Then
      LastSelectedItem = ItemWithFocus
      ProcessSelectedItem
      AdjustDisplayRange
   Else
      If LastSelectedItem < m_ListCount - 1 Then
'        repaint previously selected list item without selection gradient / focus rectangle.
'        to prevent flicker, don't repaint if last item in display is focused.
         If ItemWithFocus <> DisplayRange.LastListItem Then
            ClearPreviousSelection KeepSelectionNo
         End If
         LastSelectedItem = LastSelectedItem + 1
         ProcessSelectedItem
         AdjustDisplayRange
      End If
   End If

End Sub

Private Sub ProcessDownArrowKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, the arrow keys move just the focus        *
'* rectangle.  Selection status of each affected list item is unchanged. *
'*************************************************************************

   If ItemWithFocus < m_ListCount - 1 Then

'     repaint previously focused list item, maintaining selection status.
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo

      ItemWithFocus = ItemWithFocus + 1
      m_ListIndex = ItemWithFocus

      AdjustDisplayRange
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
      UserControl.Refresh

   End If

End Sub

Private Sub ProcessDownArrowKey_MultiSelectExtended()

'*************************************************************************
'* controls down arrow processing in MultiSelect Extended mode.          *
'*************************************************************************

   Dim NumPreviouslySelected As Long

   NumPreviouslySelected = m_SelCount

   If LastSelectedItem < m_ListCount - 1 Then

'     make sure any previously selected items are de-selected if shift key is not being pressed.
      If Not ShiftKeyDown Then
         SetSelectedArrayRange 0, m_ListCount - 1, False
      End If

'     repaint previously selected list item without selection gradient / focus rectangle.
'     to prevent flicker, don't repaint if item with focus is last in displayed range.
      If ItemWithFocus <> DisplayRange.LastListItem Then
         DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
      End If

      LastSelectedItem = LastSelectedItem + 1
      ProcessSelected_MultiSelectExtended False
      
      If ShiftKeyDown Then
'        other items may selected throughout the list; do it the hard way.
         CalculateSelCount
      Else
         m_SelCount = 1
      End If

      AdjustDisplayRange
      If NumPreviouslySelected > 1 Then
         DisplayList
      Else
         DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
         UserControl.Refresh
      End If

   Else

      If Not ShiftKeyDown Then
         SetSelectedArrayRange 0, m_ListCount - 1, False
         m_SelCount = 1
         ProcessSelected_MultiSelectExtended False
         AdjustDisplayRange
         DisplayList
      End If

   End If

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Graphics >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'******************* bitmap tiling routines by Carles P.V.
' adapted from Carles' class titled "DIB Brush - Easy Image Tiling Using FillRect"
' at Planet Source Code, txtCodeId=40585.

Private Function SetPattern(Picture As StdPicture) As Boolean

'*************************************************************************
'* creates the brush pattern for tiling into the listbox.  By Carles P.V.*
'*************************************************************************

   Dim tBI       As BITMAP
   Dim tBIH      As BITMAPINFOHEADER
   Dim Buff()    As Byte 'Packed DIB

   Dim lhDC      As Long
   Dim lhOldBmp  As Long

   If (GetObjectType(Picture) = OBJ_BITMAP) Then

'     -- Get image info
      GetObject Picture, Len(tBI), tBI

'     -- Prepare DIB header and redim. Buff() array
      With tBIH
         .biSize = Len(tBIH) '40
         .biPlanes = 1
         .biBitCount = 24
         .biWidth = tBI.bmWidth
         .biHeight = tBI.bmHeight
         .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
      End With
      ReDim Buff(1 To Len(tBIH) + tBIH.biSizeImage) '[Header + Bits]

'     -- Create DIB brush
      lhDC = CreateCompatibleDC(0)
      If (lhDC <> 0) Then
         lhOldBmp = SelectObject(lhDC, Picture)

'        -- Build packed DIB:
'        - Merge Header
         CopyMemory Buff(1), tBIH, Len(tBIH)
'        - Get and merge DIB Bits
         GetDIBits lhDC, Picture, 0, tBI.bmHeight, Buff(Len(tBIH) + 1), tBIH, DIB_RGB_COLORS

         SelectObject lhDC, lhOldBmp
         DeleteDC lhDC

'        -- Create brush from packed DIB
         DestroyPattern
         m_hBrush = CreateDIBPatternBrushPt(Buff(1), DIB_RGB_COLORS)
      End If

   End If

   SetPattern = (m_hBrush <> 0)

End Function

Private Sub Tile(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

'*************************************************************************
'* performs the tiling of the bitmap on the control.  By Carles P.V.     *
'*************************************************************************

   Dim TileRect As RECT
   Dim PtOrg    As POINTAPI

   If (m_hBrush <> 0) Then
      SetRect TileRect, x1, y1, x2, y2
      SetBrushOrgEx hdc, x1, y1, PtOrg
'     -- Tile image
      FillRect hdc, TileRect, m_hBrush
   End If

End Sub

Private Sub DestroyPattern()
   
'*************************************************************************
'* destroys the pattern brush used to tile the bitmap.  By Carles P.V.   *
'*************************************************************************
   
   If (m_hBrush <> 0) Then
      DeleteObject m_hBrush
      m_hBrush = 0
   End If

End Sub

'******************* end of bitmap tiling routines by Carles P.V.

Private Sub InitListBoxDisplayCharacteristics()

'*************************************************************************
'* initializes gradients, listitem height, and display coordinates.      *
'*************************************************************************

   Dim i As Long

'  get the height range characters in the current font.
   ListItemHeight = TextHeight("^j")

'  calculate the x coordinate for displaying listitem icons.
   If m_Style = [Standard] Then
      PicX = m_BorderWidth + 2
   Else
      PicX = m_BorderWidth + 20
   End If

'  calculate selection bar offset from left side of control.  [Standard] = 1, [CheckBox] = 21.
'  Account for listitem images possibly being displayed.
   SelBarOffset = 20 * -(m_Style = [CheckBox]) + 1
   If m_ShowItemImages Then
      If m_ItemImageSize = 0 Then
         SelBarOffset = SelBarOffset + ListItemHeight
      Else
         SelBarOffset = SelBarOffset + m_ItemImageSize
      End If
   End If

'  initialize text draw coordinates and boundaries.
   InitTextDisplayCharacteristics

'  create a virtual bitmap that will hold the background gradient or picture.  Portions of
'  this virtual bitmap are blitted to the control background to restore the background
'  gradient/picture when list items are changed.  Saves time over repainting whole control
'  when we're just doing things like adding a listbox item or changing selection gradient.
   CreateVirtualBackgroundDC

'  calculate the various gradients that may be used in the control.
   CalculateGradients

'  place either the picture or gradient background onto the virtual DC.
   If IsPictureThere(m_ActivePicture) Then
      DisplayPicture
      CreateBorder
'     transfer the picture (with border) to the virtual DC bitmap.
      i = BitBlt(VirtualBackgroundDC, 0, 0, ScaleWidth, ScaleHeight, hdc, 0, 0, vbSrcCopy)
   Else
'     paint the gradient onto the virtual DC bitmap.
      Call StretchDIBits(VirtualBackgroundDC, _
                         0, 0, _
                         ScaleWidth, _
                         ScaleHeight, _
                         0, 0, _
                         ScaleWidth, _
                         ScaleHeight, _
                         BGlBits(0), _
                         BGuBIH, _
                         DIB_RGB_COLORS, _
                         vbSrcCopy)
'     transfer the gradient in the virtual bitmap to the usercontrol.
      i = BitBlt(hdc, 0, 0, ScaleWidth, ScaleHeight, VirtualBackgroundDC, 0, 0, vbSrcCopy)
      CreateBorder
   End If

End Sub

Private Sub InitTextDisplayCharacteristics()

'*************************************************************************
'* Calculate text display coordinates and boundaries.  I keep this a     *
'* separate routine for when properties (such as .BorderWidth) that      *
'* affect text display are changed programmatically.  I can then quickly *
'* change text boundaries.                                               *
'*************************************************************************

   Dim AvailableDisplayHeight As Long    ' height, in pixels of displayable listbox area.
   Dim i As Long                         ' loop variable.

'  determine the number of items that can be displayed given listbox height, list
'  item height in the current font, and display style (normal or checkbox).
'  Also account for listitem images possibly being displayed.
   AvailableDisplayHeight = ScaleHeight - (Y_CLEARANCE * 2) - (m_BorderWidth * 2)
   If m_Style = [Standard] Then
      If m_ShowItemImages Then
         If ListItemHeight >= m_ItemImageSize Then
            MaxDisplayItems = Int(AvailableDisplayHeight / ListItemHeight)
         Else
            MaxDisplayItems = Int(AvailableDisplayHeight / m_ItemImageSize)
         End If
      Else
         MaxDisplayItems = Int(AvailableDisplayHeight / ListItemHeight)
      End If
   Else
      If m_ShowItemImages Then
         If ListItemHeight >= m_ItemImageSize Then
            If ListItemHeight < MIN_FONT_HEIGHT Then
               MaxDisplayItems = Int(AvailableDisplayHeight / ((ListItemHeight + 2) + (MIN_FONT_HEIGHT - ListItemHeight)))
            Else
               MaxDisplayItems = Int(AvailableDisplayHeight / (ListItemHeight + 2))
            End If
         Else
            MaxDisplayItems = Int(AvailableDisplayHeight / (m_ItemImageSize + 2))
         End If
      Else
         If ListItemHeight < MIN_FONT_HEIGHT Then
            MaxDisplayItems = Int(AvailableDisplayHeight / ((ListItemHeight + 2) + (MIN_FONT_HEIGHT - ListItemHeight)))
         Else
            MaxDisplayItems = Int(AvailableDisplayHeight / (ListItemHeight + 2))
         End If
      End If

   End If

'  determine the number of pixels of clearance from the border to start drawing text,
'  given the display mode (checkbox or normal) and listitem image display mode.
   If m_Style = [Standard] Then
      TextClearance = m_BorderWidth + 3
   Else
      TextClearance = m_BorderWidth + 23
   End If
   If m_ShowItemImages Then
      If m_ItemImageSize = 0 Then
         TextClearance = TextClearance + ListItemHeight + 1
      Else
         TextClearance = TextClearance + m_ItemImageSize + 1
      End If
   End If

'  initialize the y coordinate array.  Make the spacing between
'  list items a little wider if Checkbox display mode is active.
   YCoords(0) = m_BorderWidth + Y_CLEARANCE
   For i = 1 To MaxDisplayItems - 1
      If m_Style = [Standard] Then
         If m_ShowItemImages Then
            If ListItemHeight >= m_ItemImageSize Then
               YCoords(i) = YCoords(i - 1) + ListItemHeight
            Else
               YCoords(i) = YCoords(i - 1) + m_ItemImageSize
            End If
         Else
            YCoords(i) = YCoords(i - 1) + ListItemHeight
         End If
      Else
'        helps keeps checkboxes from getting squished when displaying in very small fonts.
         If m_ShowItemImages Then
            If ListItemHeight >= m_ItemImageSize Then
               YCoords(i) = YCoords(i - 1) + ListItemHeight + 2
            Else
               YCoords(i) = YCoords(i - 1) + m_ItemImageSize + 2
            End If
         Else
            If ListItemHeight < MIN_FONT_HEIGHT Then
               YCoords(i) = YCoords(i - 1) + ListItemHeight + 2 + (MIN_FONT_HEIGHT - ListItemHeight)
            Else
               YCoords(i) = YCoords(i - 1) + ListItemHeight + 2
            End If
         End If
      End If
   Next i

End Sub

Private Sub DrawRectangle(x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)

'*************************************************************************
'* draws the checkbox, thumb border, and focus rectangles.               *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim hBrush As Long        ' the brush pattern used to 'paint' the border.
   Dim hRgn1  As Long        ' the outer boundary of the border region.
   Dim hRgn2  As Long        ' the inner boundary of the border region.

'  create the outer region.
   hRgn1 = CreateRoundRectRgn(x1, y1, x2, y2, 0, 0)
'  create the inner region.
   hRgn2 = CreateRoundRectRgn(x1 + 1, y1 + 1, x2 - 1, y2 - 1, 0, 0)
   
'  combine the regions into one border region.
   CombineRgn hRgn2, hRgn1, hRgn2, 3

'  create and apply the color brush.
   hBrush = CreateSolidBrush(TranslateColor(lColor))
   FillRgn hdc, hRgn2, hBrush

'  free the memory.
   DeleteObject hRgn2
   DeleteObject hBrush
   DeleteObject hRgn1

ErrHandler:
   Exit Sub

End Sub

Private Sub CalculateGradients()

'*************************************************************************
'* define all gradients (background, scroll track/button, selection bar).*
'* I split them up into different procedures so that when a particular   *
'* gradient property is changed, all gradients don't get regenerated.    *
'*************************************************************************

   CalculateBackGroundGradient
   CalculateHighlightBarGradient
   CalculateVerticalTrackbarGradients
   CalculateScrollbarButtonGradient
   CalculateScrollbarThumbGradient

End Sub

Private Sub CalculateBackGroundGradient()

'*************************************************************************
'* calculate the gradient for the background.  Even if a picture is used *
'* instead of a gradient, this allows control user to switch back and    *
'* forth between those two options in design or runtime modes.           *
'*************************************************************************

   CalculateGradient ScaleWidth, ScaleHeight, TranslateColor(m_ActiveBackColor1), TranslateColor(m_ActiveBackColor2), m_BackAngle, m_BackMiddleOut, BGuBIH, BGlBits()

End Sub

Private Sub CalculateHighlightBarGradient()

'*************************************************************************
'*  calculate the gradient for the selected item highlight bar.          *
'*************************************************************************

   CalculateGradient ScaleWidth - (m_BorderWidth * 2), ListItemHeight, TranslateColor(m_ActiveSelColor1), TranslateColor(m_ActiveSelColor2), 90, True, SeluBIH, SellBits()

End Sub

Private Sub CalculateScrollbarThumbGradient()

'*************************************************************************
'* calculate the vertical scrollbar thumb gradient.                      *
'*************************************************************************

'  sized to scrollbar height at first, sized on the fly by StretchDIBits when drawing thumb.
   CalculateGradient ScaleWidth - ScrollBarButtonWidth, vScrollTrackHeight, TranslateColor(m_ActiveThumbColor1), TranslateColor(m_ActiveThumbColor2), 180, True, vThumbuBIH, vThumblBits()

End Sub

Private Sub CalculateVerticalTrackbarGradients()

'*************************************************************************
'* master routine for generating clicked/unclicked trackbar gradients.   *
'*************************************************************************

'  determine the height of the scrollbar track (area between up and down buttons).
   vScrollTrackHeight = CalculateScrollTrackHeight

   CalculateVerticalTrackbarGradientUnclicked
   CalculateVerticalTrackbarGradientClicked

End Sub

Private Sub CalculateVerticalTrackbarGradientUnclicked()

'*************************************************************************
'* calculate the gradient for the vertical scrollbar trackbar when the   *
'* mouse is not down on the trackbar.                                    *
'*************************************************************************

   CalculateGradient ScaleWidth - ScrollBarButtonWidth, vScrollTrackHeight, TranslateColor(m_ActiveTrackBarColor1), TranslateColor(m_ActiveTrackBarColor2), 180, True, VTrackuBIH, VTracklBits()

End Sub

Private Sub CalculateVerticalTrackbarGradientClicked()

'*************************************************************************
'* calculate the gradient for the vertical scrollbar trackbar when the   *
'* mouse is down on the trackbar.                                        *
'*************************************************************************

   CalculateGradient ScaleWidth - ScrollBarButtonWidth, vScrollTrackHeight, TranslateColor(m_TrackClickColor1), TranslateColor(m_TrackClickColor2), 180, True, vClickTrackuBIH, vClickTracklBits()

End Sub

Private Sub CalculateScrollbarButtonGradient()

'*************************************************************************
'*  calculate the gradient for the vertical scrollbar buttons.           *
'*************************************************************************

   CalculateGradient ScaleWidth - ScrollBarButtonWidth, ScrollBarButtonHeight, TranslateColor(m_ActiveButtonColor1), TranslateColor(m_ActiveButtonColor2), 180, True, TrackButtonuBIH, TrackButtonlBits()

End Sub

Private Sub RedrawControl()

'*************************************************************************
'* master routine for painting of MorphListBox control.                  *
'*************************************************************************

'  if the .RedrawFlag property is false, then don't redraw.  This property is
'  set to False by the programmer before large operations on the listbox (for
'  example, adding or removing a thousand items) and set back to True after the
'  operations are complete.  This saves unnecessary and time-consuming redraws.
   If m_RedrawFlag Then
      SetBackGround
      CreateBorder
      DisplayList
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

Private Sub SetBackGround()

'*************************************************************************
'* displays control's background gradient or picture in initial draw.    *
'*************************************************************************

   If IsPictureThere(m_ActivePicture) Then
'     if the .Picture property has been defined, it takes precedence over gradient.
      DisplayPicture
   Else
'     paint the gradient onto the actual usercontrol DC.  Most subsequent repaints are handled
'     by blitting the appropriate gradient portions from the virtual bitmap's DC to the usercontrol.
'     Thanks to RedBird77 for tweaking this to work correctly with wide borders!
      Call StretchDIBits(hdc, m_BorderWidth, m_BorderWidth, _
                         ScaleWidth - (m_BorderWidth * 2), _
                         ScaleHeight - (m_BorderWidth * 2), _
                         m_BorderWidth, m_BorderWidth, _
                         ScaleWidth - (m_BorderWidth * 2), _
                         ScaleHeight - (m_BorderWidth * 2), _
                         BGlBits(0), BGuBIH, DIB_RGB_COLORS, vbSrcCopy)
   End If

End Sub

Private Sub DisplayPicture()

'*************************************************************************
'* if the .Picture property is defined, paints the picture onto the      *
'* control.  If picture tiling is indicated, that is performed.          *
'*************************************************************************

   Select Case m_ActivePictureMode
      Case [Normal]
         Set UserControl.Picture = m_ActivePicture
      Case [Tiled]
         SetPattern m_ActivePicture
         Tile hdc, m_BorderWidth, m_BorderWidth, ScaleWidth - m_BorderWidth, ScaleHeight - m_BorderWidth
      Case [Stretch]
         StretchPicture
   End Select

End Sub

Private Sub StretchPicture()

'*************************************************************************
'* stretch bitmap to fit listbox background.  Thanks to LaVolpe for the  *
'* suggestion and AllAPI.net / VBCity.com for the learning to do it.     *
'*************************************************************************

   Dim TempBitmap As BITMAP       ' bitmap structure that temporarily holds picture.
   Dim CreateDC As Long           ' used in creating temporary bitmap structure virtual DC.
   Dim TempBitmapDC As Long       ' virtual DC of temporary bitmap structure.
   Dim TempBitmapOld As Long      ' used in destroying temporary bitmap structure virtual DC.
   Dim r As Long                  ' result long for StretchBlt call.

'  create a temporary bitmap and DC to place the picture in.
   GetObjectAPI m_ActivePicture.Handle, Len(TempBitmap), TempBitmap
   CreateDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   TempBitmapDC = CreateCompatibleDC(CreateDC)
   TempBitmapOld = SelectObject(TempBitmapDC, m_ActivePicture.Handle)

'  streeeeeeeetch it...
   r = StretchBlt(hdc, m_BorderWidth, m_BorderWidth, _
                  ScaleWidth - m_BorderWidth * 2, ScaleHeight - m_BorderWidth * 2, _
                  TempBitmapDC, _
                  0, 0, _
                  TempBitmap.bmWidth, _
                  TempBitmap.bmHeight, vbSrcCopy)

'  destroy temporary bitmap DC.
   SelectObject TempBitmapDC, TempBitmapOld
   DeleteDC TempBitmapDC
   DeleteDC CreateDC

End Sub

Private Function IsPictureThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks for existence of a picture.  Thanks to Roger Gilchrist.        *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then
         IsPictureThere = Pic.Width <> 0
      End If
   End If

End Function

Private Sub CreateBorder()

'*************************************************************************
'* draws the border around the control, using appropriate curvatures.    *
'*************************************************************************

   Dim i       As Long   ' return variable for BitBlt.
   Dim hRgn1   As Long   ' the outer region of the border.
   Dim hRgn2   As Long   ' the inner region of the border.
   Dim hBrush  As Long   ' the solid-color brush used to paint the combined border regions.

'  create the outer region.
   hRgn1 = pvGetRoundedRgn(0, 0, _
                           ScaleWidth, _
                           ScaleHeight, _
                           m_CurveTopLeft, _
                           m_CurveTopRight, _
                           m_CurveBottomLeft, _
                           m_CurveBottomRight)
'  create the inner region.
   hRgn2 = pvGetRoundedRgn(m_BorderWidth, _
                           m_BorderWidth, _
                           ScaleWidth - m_BorderWidth, _
                           ScaleHeight - m_BorderWidth, _
                           m_CurveTopLeft, _
                           m_CurveTopRight, _
                           m_CurveBottomLeft, _
                           m_CurveBottomRight)

'  combine the outer and inner regions.
   CombineRgn hRgn2, hRgn1, hRgn2, RGN_DIFF

'  create the solid brush pattern used to color the combined regions.
   hBrush = CreateSolidBrush(TranslateColor(m_ActiveBorderColor))

'  color the combined regions.
   FillRgn hdc, hRgn2, hBrush

'  set the container's visibility region.
   SetWindowRgn hwnd, hRgn1, True

'  delete created objects to restore memory.
   DeleteObject hBrush
   DeleteObject hRgn1
   DeleteObject hRgn2

'  if we are redrawing the control because of a change to the .Picture property,
'  this is the time to re-blit the new picture/border to the virtual DC. I do
'  it here because I blit the entire control surface, including border.
   If ChangingPicture Then
      i = BitBlt(VirtualBackgroundDC, 0, 0, ScaleWidth, ScaleHeight, hdc, 0, 0, vbSrcCopy)
      ChangingPicture = False
   End If

End Sub

Private Function pvGetRoundedRgn(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, _
                                 ByVal TopLeftRadius As Long, ByVal TopRightRadius As Long, _
                                 ByVal BottomLeftRadius As Long, ByVal BottomRightRadius As Long) As Long

'*************************************************************************
'* allows each corner of the container to have its own curvature.        *
'* Code by the Amazing Carles P.V.  Thanks a million (as usual) Carles.  *
'*************************************************************************

   Dim hRgnMain As Long   ' the original "starting point" region.
   Dim hRgnTmp1 As Long   ' the first region that defines a corner's radius.
   Dim hRgnTmp2 As Long   ' the second region that defines a corner's radius.

'  bounding region.
   hRgnMain = CreateRectRgn(x1, y1, x2, y2)

'  top-left corner.
   hRgnTmp1 = CreateRectRgn(x1, y1, x1 + TopLeftRadius, y1 + TopLeftRadius)
   hRgnTmp2 = CreateEllipticRgn(x1, y1, x1 + 2 * TopLeftRadius, y1 + 2 * TopLeftRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

'  top-right corner.
   hRgnTmp1 = CreateRectRgn(x2, y1, x2 - TopRightRadius, y1 + TopRightRadius)
   hRgnTmp2 = CreateEllipticRgn(x2 + 1, y1, x2 + 1 - 2 * TopRightRadius, y1 + 2 * TopRightRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

'  bottom-left corner.
   hRgnTmp1 = CreateRectRgn(x1, y2, x1 + BottomLeftRadius, y2 - BottomLeftRadius)
   hRgnTmp2 = CreateEllipticRgn(x1, y2 + 1, x1 + 2 * BottomLeftRadius, y2 + 1 - 2 * BottomLeftRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

'  bottom-right corner.
   hRgnTmp1 = CreateRectRgn(x2, y2, x2 - BottomRightRadius, y2 - BottomRightRadius)
   hRgnTmp2 = CreateEllipticRgn(x2 + 1, y2 + 1, x2 + 1 - 2 * BottomRightRadius, y2 + 1 - 2 * BottomRightRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

   pvGetRoundedRgn = hRgnMain

End Function

Private Sub CombineRegions(ByVal Region1 As Long, ByVal Region2 As Long, ByVal MainRegion As Long)

'*************************************************************************
'* combines outer/inner rectangular regions for border painting.         *
'*************************************************************************

   Call CombineRgn(Region1, Region1, Region2, RGN_DIFF)
   Call CombineRgn(MainRegion, MainRegion, Region1, RGN_DIFF)
   Call DeleteObject(Region1)
   Call DeleteObject(Region2)

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

Private Sub DisplayList(Optional vThumbYPos As Single = -1)

'*************************************************************************
'* controls the display of all visible list items.                       *
'*************************************************************************

   Dim i As Long     ' loop variable.

   VerticalScrollBarActive = (m_ListCount > MaxDisplayItems)

'  Calculate scroll bar impact on listitem display.  Zero if scroll bar not active.
   lBarWid = ScrollBarButtonWidth * -VerticalScrollBarActive

'  repaint the picture background or gradient.
   SetBackGround

'  safety net.
   If DisplayRange.FirstListItem = -1 Then
      UserControl.Refresh
      Exit Sub
   End If

'  if the entire list will fit in the listbox...
   If m_ListCount <= MaxDisplayItems Then
      DisplayRange.FirstListItem = 0
      DisplayRange.LastListItem = m_ListCount - 1
   Else
'     if not displaying the very end of the list...
      If Not InDisplayedItemRange(m_ListCount - 1) Then
         DisplayRange.LastListItem = DisplayRange.FirstListItem + MaxDisplayItems - 1
      Else
         DisplayRange.LastListItem = m_ListCount - 1
         DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems
      End If
   End If

'  set the .TopIndex property.
   m_TopIndex = DisplayRange.FirstListItem

'  display the appropriate listbox items, as selected or unselected.
   For i = DisplayRange.FirstListItem To DisplayRange.LastListItem
      DisplayListBoxItem i, SelectedArray(i), FocusRectangleNo
   Next i

'  if there's a list entry with a focus rectangle visible, redraw it.
   If m_Style = [Standard] Then
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
   Else
      If LastSelectedItem = -1 Then
         LastSelectedItem = 0
         ItemWithFocus = 0
         m_ListIndex = 0
      End If
      DisplayListBoxItem ItemWithFocus, DrawAsSelected, FocusRectangleYes
   End If

'  draw scrollbar if called for.
   If VerticalScrollBarActive Then
      DisplayVerticalScrollBar vThumbYPos
   End If

'  if there is a picture background instead of a gradient, or any control corners have
'  curvature, the border needs to be redrawn.  Thanks to Light Templer for catching this bug.
   If IsPictureThere(m_ActivePicture) Or (m_CurveTopLeft + m_CurveTopRight + m_CurveBottomLeft + m_CurveBottomRight > 0) Then
      CreateBorder
   End If

   UserControl.Refresh

End Sub

Private Sub DisplayListBoxItem(ByVal Index As Long, ByVal ItemSelected As Boolean, ByVal FocusRectFlag As Boolean)

'*************************************************************************
'* displays one (selected or unselected) listbox entry, using the spec-  *
'* ified ListArray index, in appropriate style (CheckBox or Standard).   *
'* Thanks to Redbird77 for optimizing the hell out of this routine!      *
'*************************************************************************

   Dim r           As RECT    ' the listitem text display rectangle.
   Dim lRet        As Long    ' bitblt function return.
   Dim nDisp       As Long    ' the index in the viewable list area of the listitem.
   Dim yStart      As Long
   Dim SelY        As Long

   If Not InDisplayedItemRange(Index) Or (m_ListFont Is Nothing) Then
      Exit Sub
   End If

   nDisp = GetDisplayIndexFromArrayIndex(Index)

   If nDisp < 0 Then
      nDisp = 0
   End If

   If m_ShowItemImages Then
      If ListItemHeight >= m_ItemImageSize Then
         SelY = YCoords(nDisp)
      Else
         SelY = YCoords(nDisp) + (m_ItemImageSize \ 4)
      End If
   Else
      SelY = YCoords(nDisp)
   End If

'  Draw selection bar background.  However, don't draw it if in checkbox mode
'  and item is selected but not the focused item (Index<> LastSelectedItem).  This is
'  because in CheckBox mode, only the item with focus has the selection bar background.
   If (m_Style = [Standard] And ItemSelected) Or (m_Style = [CheckBox] And ItemSelected And Index = LastSelectedItem) Then

'     if in checkbox mode, repaint gradient under checkbox so that checkbox is
'     "unchecked" when drawing checkbox (if the list item is now unselected).
'     Account for listitem images possibly being displayed.
      If m_Style = [CheckBox] Then
'        center the checkbox.
         If m_ShowItemImages Then
            If ListItemHeight >= m_ItemImageSize Then
               yStart = YCoords(nDisp) + (ListItemHeight - MIN_FONT_HEIGHT) / 2
            Else
               yStart = YCoords(nDisp) + (m_ItemImageSize \ 4)
            End If
         Else
            If ListItemHeight > MIN_FONT_HEIGHT Then
               yStart = YCoords(nDisp) + (ListItemHeight - MIN_FONT_HEIGHT) / 2
            Else
               yStart = YCoords(nDisp)
            End If
         End If
         lRet = BitBlt(hdc, m_BorderWidth, yStart, _
                       m_BorderWidth + 16, 13, _
                       VirtualBackgroundDC, _
                       m_BorderWidth, yStart, vbSrcCopy)
      End If

      UserControl.ForeColor = TranslateColor(m_ActiveSelTextColor)
'     draw the selection highlight gradient.
      Call StretchDIBits(hdc, m_BorderWidth + SelBarOffset, SelY, _
                         ScaleWidth - (m_BorderWidth * 2) - lBarWid - SelBarOffset, ListItemHeight, _
                         0, 0, _
                         ScaleWidth - (m_BorderWidth * 2), ListItemHeight, _
                         SellBits(0), SeluBIH, _
                         DIB_RGB_COLORS, vbSrcCopy)

   Else

'     item is not highlighted; draw regular item background.
      UserControl.ForeColor = TranslateColor(m_ActiveTextColor)
      lRet = BitBlt(hdc, m_BorderWidth, SelY, _
                    ScaleWidth - (m_BorderWidth * 2) - lBarWid, _
                    ListItemHeight + 1, _
                    VirtualBackgroundDC, _
                    m_BorderWidth, SelY - 1, vbSrcCopy)

   End If

'  display listitem image if necessary.
   If m_ShowItemImages Then
      If ImageIndexArray(Index) <> -1 Then
         If m_ItemImageSize = 0 Then
'           if the ItemImageSize property is zero, that means we paint icon
'           in same width/height dimensions as listitem text height.
            UserControl.PaintPicture Images(ImageIndexArray(Index)), PicX, YCoords(nDisp), ListItemHeight - 1, ListItemHeight - 1
         Else
'           otherwise, set the width and height of the icon to ItemImageSize.
'           Determine the Y coordinate based on list item text height.
            If m_ItemImageSize < ListItemHeight Then
               UserControl.PaintPicture Images(ImageIndexArray(Index)), _
                                        PicX, YCoords(nDisp) + (ListItemHeight - m_ItemImageSize) \ 2, _
                                        m_ItemImageSize, m_ItemImageSize
            Else
               UserControl.PaintPicture Images(ImageIndexArray(Index)), _
                                        PicX, YCoords(nDisp), _
                                        m_ItemImageSize, m_ItemImageSize
            End If
         End If
      End If
   End If

'  calculate text rectangle size and position.
   With r
      .Left = TextClearance
      If m_ShowItemImages Then
         If ListItemHeight >= m_ItemImageSize Then
            .Top = YCoords(nDisp)
         Else
            .Top = YCoords(nDisp) + m_ItemImageSize \ 4
         End If
      Else
         .Top = YCoords(nDisp)
      End If
      .Bottom = .Top + ListItemHeight
      .Right = ScaleWidth - m_BorderWidth - lBarWid
   End With

'  display the text using DrawText api.
   Call DrawText(UserControl.hdc, ListArray(Index), -1, r, DT_LEFT Or DT_NOPREFIX Or DT_SINGLELINE)

'  display the checkbox, if the .Style property is set to CheckBox.
   If m_Style = [CheckBox] Then
      Call DisplayCheckBox(nDisp, SelectedArray(Index))
   End If

'  if the control and item both have focus, display the item's focus rectangle.
   If FocusRectFlag And HasFocus Then
      Call DisplayFocusRectangle(nDisp)
   End If

End Sub

Private Sub DrawText(ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long)

'*************************************************************************
'* draws the text with Unicode support based on OS version.              *
'* Thanks to Richard Mewett.                                             *
'*************************************************************************

   If mWindowsNT Then
      DrawTextW hdc, StrPtr(lpString), nCount, lpRect, wFormat
   Else
      DrawTextA hdc, lpString, nCount, lpRect, wFormat
   End If

End Sub

Private Sub DisplayCheckBox(ByVal Index As Long, ByVal SelectedStatus As Boolean)

'*************************************************************************
'* draws an item-centered, one-pixel wide checkbox next to a list item.  *
'* If item is selected, draws a checkmark in one of three styles.        *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim YCoordStart As Long   ' starting y position of box based on font size.

'  center the checkbox, if the font is higher then the checkbox.
   If m_ShowItemImages Then
      If ListItemHeight >= m_ItemImageSize Then
         YCoordStart = YCoords(Index) + (ListItemHeight - MIN_FONT_HEIGHT) / 2
      Else
         YCoordStart = YCoords(Index) + (m_ItemImageSize \ 4)
      End If
   Else
      If ListItemHeight > MIN_FONT_HEIGHT Then
         YCoordStart = YCoords(Index) + (ListItemHeight - MIN_FONT_HEIGHT) / 2
      Else
         YCoordStart = YCoords(Index)
      End If
   End If

'  display the checkbox.
   DrawRectangle m_BorderWidth + 4, YCoordStart, m_BorderWidth + 18, YCoordStart + 14, m_CheckBoxColor

   If SelectedStatus Then
      Select Case m_CheckStyle
         Case [Arrow]
            DisplayCheckBoxArrow YCoordStart
         Case [Tick]
            DisplayCheckBoxTickMark YCoordStart
         Case [X]
            DisplayCheckBoxX YCoordStart
      End Select
   End If

ErrHandler:
   Exit Sub

End Sub

Private Sub DisplayCheckBoxArrow(ByVal Index As Long)

'*************************************************************************
'* draws an arrow in the checkbox of a selected list item when the       *
'* .Style property is set to Checkbox mode.                              *
'*************************************************************************

   Dim hPO           As Long       ' selected pen object.
   Dim hPN           As Long       ' pen object for drawing checkmark.
   Dim r             As Long       ' loop and result variable for api calls.
   Dim x1            As Long       ' the x coordinate of the start of the checkmark.
   Dim y1            As Long       ' the y coordinate of the start of the checkmark vertical line.
   Dim y2            As Long       ' the y coordinate of the end of the checkmark vertical line.

'  determine x coordinate of first part of check arrow to draw.
   x1 = m_BorderWidth + 9

   y1 = Index + 2
   y2 = 10
   hPN = CreatePen(0, 2, m_CheckboxArrowColor)
   hPO = SelectObject(hdc, hPN)

'  draw the checkmark.
   MoveTo hdc, x1, y1, ByVal 0&
   For r = 1 To 6
      LineTo hdc, x1, y1 + y2
      x1 = x1 + 1
      y1 = y1 + 1
      y2 = y2 - 2
      MoveTo hdc, x1, y1, ByVal 0&
   Next r

'  delete the pen object.
   r = SelectObject(hdc, hPO)
   r = DeleteObject(hPN)

End Sub

Private Sub DisplayCheckBoxX(ByVal Index As Long)

'*************************************************************************
'* draws an X in the checkbox of a selected list item when the .Style    *
'* property is set to Checkbox mode.                                     *
'*************************************************************************

   Dim i As Long

   For i = 1 To 2: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 4, m_CheckboxArrowColor: Next i
   For i = 6 To 7: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 4, m_CheckboxArrowColor: Next i
   For i = 1 To 3: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 5, m_CheckboxArrowColor: Next i
   For i = 5 To 7: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 5, m_CheckboxArrowColor: Next i
   For i = 2 To 6: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 6, m_CheckboxArrowColor: Next i
   For i = 3 To 5: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 7, m_CheckboxArrowColor: Next i
   For i = 2 To 6: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 8, m_CheckboxArrowColor: Next i
   For i = 1 To 3: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 9, m_CheckboxArrowColor: Next i
   For i = 5 To 7: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 9, m_CheckboxArrowColor: Next i
   For i = 1 To 2: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 10, m_CheckboxArrowColor: Next i
   For i = 6 To 7: SetPixelV hdc, m_BorderWidth + 7 + i, Index + 10, m_CheckboxArrowColor: Next i

End Sub

Private Sub DisplayCheckBoxTickMark(ByVal Index As Long)

'*************************************************************************
'* draws a tick mark in the checkbox of a selected list item when the    *
'* .Style property is set to Checkbox mode.                              *
'*************************************************************************

   Dim i As Long

   For i = 9 To 12: SetPixelV hdc, m_BorderWidth + 5 + i, Index + 3, m_CheckboxArrowColor: Next i
   For i = 8 To 11: SetPixelV hdc, m_BorderWidth + 5 + i, Index + 4, m_CheckboxArrowColor: Next i
   For i = 7 To 10: SetPixelV hdc, m_BorderWidth + 5 + i, Index + 5, m_CheckboxArrowColor: Next i
   For i = 1 To 2: SetPixelV hdc, m_BorderWidth + 5 + i, Index + 6, m_CheckboxArrowColor: Next i
   For i = 6 To 9: SetPixelV hdc, m_BorderWidth + 5 + i, Index + 6, m_CheckboxArrowColor: Next i
   For i = 1 To 3: SetPixelV hdc, m_BorderWidth + 5 + i, Index + 7, m_CheckboxArrowColor: Next i
   For i = 5 To 8: SetPixelV hdc, m_BorderWidth + 5 + i, Index + 7, m_CheckboxArrowColor: Next i
   For i = 1 To 7: SetPixelV hdc, m_BorderWidth + 5 + i, Index + 8, m_CheckboxArrowColor: Next i
   For i = 2 To 6: SetPixelV hdc, m_BorderWidth + 5 + i, Index + 9, m_CheckboxArrowColor: Next i
   For i = 3 To 5: SetPixelV hdc, m_BorderWidth + 5 + i, Index + 10, m_CheckboxArrowColor: Next i
   SetPixelV hdc, m_BorderWidth + 5 + 4, Index + 11, m_CheckboxArrowColor

End Sub

Private Sub DisplayFocusRectangle(ByVal DispIndex As Long)

'*************************************************************************
'* draws a custom focus rectangle around the specified listbox entry.    *
'* Originally I used the DrawFocusRect API, but found that the default   *
'* dotted focus rectangle was often hard to see against darker back-     *
'* grounds.  So I did this to give the user complete control over color. *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim x1     As Long        ' first x coordinate for focus rectangle.
   Dim y1     As Long        ' first y coordinate for focus rectangle.
   Dim x2     As Long        ' second x coordinate for focus rectangle.
   Dim y2     As Long        ' second y coordinate for focus rectange.

'  calculate the x and y coordinates of the focus rectangle.
   If m_Style = [Standard] Then
      x1 = m_BorderWidth + 1
   Else
      x1 = m_BorderWidth + 21
   End If

   If m_ShowItemImages Then
      If m_ItemImageSize = 0 Then
         x1 = x1 + ListItemHeight
      Else
         x1 = x1 + m_ItemImageSize
      End If
   End If

'  define the top and bottom y coordinates for the focus rectangle.
   If m_ShowItemImages Then
      If ListItemHeight >= m_ItemImageSize Then
         y1 = YCoords(DispIndex)
      Else
         y1 = YCoords(DispIndex) + (m_ItemImageSize \ 4)
      End If
   Else
      y1 = YCoords(DispIndex)
   End If
   y2 = y1 + ListItemHeight

'  define right edge of focus rectangle, accounting for possible scrollbar.
   If Not VerticalScrollBarActive Then
      x2 = ScaleWidth - m_BorderWidth
   Else
      x2 = ScaleWidth - m_BorderWidth - ScrollBarButtonWidth
   End If

   DrawRectangle x1, y1, x2, y2, m_FocusRectColor

ErrHandler:
   Exit Sub

End Sub

Private Sub CreateVirtualBackgroundDC()

'*************************************************************************
'* creates a virtual bitmap, with its own DC, that will hold a copy of   *
'* the control's background gradient (or picture).  This is used by      *
'* BitBlt to update just the part of the control's background that is    *
'* changed when the selected status or text of a listbox entry has       *
'* changed.  This allows for lightning-quick updates of the background   *
'* and display of individual list items without having to repaint the    *
'* whole control after a list item is added or selection status changed. *
'*************************************************************************

'  safety net that makes sure virtal DC doesn't already exist.
   If IsCreated Then
      DestroyVirtualDC
   End If

'  Create a memory device context to use
   VirtualBackgroundDC = CreateCompatibleDC(hdc)

'  define it as a bitmap so that drawing can be performed to the virtual DC.
   mMemoryBitmap = CreateCompatibleBitmap(hdc, ScaleWidth, ScaleHeight)
   mOrginalBitmap = SelectObject(VirtualBackgroundDC, mMemoryBitmap)

End Sub

Private Function IsCreated() As Boolean

'*************************************************************************
'* checks the handle of the created DC and returns if it exists.         *
'*************************************************************************

   IsCreated = (VirtualBackgroundDC <> 0)

End Function

Private Sub DestroyVirtualDC()

'*************************************************************************
'* eliminates the virtual background dc bitmap on control's termination. *
'*************************************************************************

   If Not IsCreated Then
      Exit Sub
   End If

   Call SelectObject(VirtualBackgroundDC, mOrginalBitmap)
   Call DeleteObject(mMemoryBitmap)
   Call DeleteDC(VirtualBackgroundDC)
   VirtualBackgroundDC = -1

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<< Public Methods and Method Helper Routines >>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Sub DisplayFrom(ByVal ItemIndex As Long)

'*************************************************************************
'* .DisplayFrom method.  Displays list items from ItemIndex to maximum   *
'* number of displayable list items.  If ItemIndex is anywhere within    *
'* last displayable page of the list, the entire last page is displayed. *
'*************************************************************************

   If m_Enabled Then

'     if the entire list can fit in the display area, just exit.
      If m_ListCount <= MaxDisplayItems Then
         Exit Sub
      End If

'     make sure ItemIndex actually points to an existing list item before continuing.
      If ItemIndex < 0 Or ItemIndex > m_ListCount - 1 Then
         Exit Sub
      End If

'     calculate the display range.
      DisplayRange.LastListItem = ItemIndex + MaxDisplayItems - 1
      If DisplayRange.LastListItem > m_ListCount - 1 Then
         DisplayRange.LastListItem = m_ListCount - 1
         DisplayRange.FirstListItem = DisplayRange.LastListItem - MaxDisplayItems + 1
      Else
         DisplayRange.FirstListItem = ItemIndex
      End If

'     redisplay the new range of list items.
      DisplayList

   End If

End Sub

Public Function FindIndex(ByVal sStringToMatch As String, Optional CaseSensitive As Boolean = False) As Long

'*************************************************************************
'* .FindIndex method.  Returns the .List() index for the supplied string *
'* or -1 if the string is not found in the list.  Uses a binary search   *
'* algorithm if the .Sorted property is true; otherwise has to rely on a *
'* much slower sequential search.  If the optional CaseSensitive boolean *
'* parameter is set to True, case of supplied string must match the case *
'* of the intended target in the .List array for match to be successful. *
'*************************************************************************

   Dim i      As Long      ' loop variable.
   Dim tmpStr As String    ' string that holds current .List array item for equality comparison.
   Dim iLBound As Long     ' lower bound of list array portion currently being searched.
   Dim iUBound As Long     ' upper bound of list array portion currently being searched.
   Dim iMiddle As Long     ' middle of list array portion currently being searched.

   If m_Enabled Then

'     if we don't care about case sensitivity, make the source and target strings lower case.
      If Not CaseSensitive Then
         sStringToMatch = LCase(sStringToMatch)
      End If

      If m_Sorted Then

'        if list is sorted, a binary search can be used.
         iLBound = 0
         iUBound = m_ListCount - 1
         Do
            iMiddle = (iLBound + iUBound) \ 2
            tmpStr = ListArray(iMiddle)
            If Not CaseSensitive Then
               tmpStr = LCase(tmpStr)
            End If
            If tmpStr = sStringToMatch Then
               FindIndex = iMiddle
               Exit Function
            ElseIf tmpStr < sStringToMatch Then
               iLBound = iMiddle + 1
            Else
               iUBound = iMiddle - 1
            End If
         Loop Until iLBound > iUBound

      Else

'        if list is not sorted, a sequential search must be performed.
         For i = 0 To m_ListCount - 1
            tmpStr = ListArray(i)
            If Not CaseSensitive Then
               tmpStr = LCase(tmpStr)
            End If
            If tmpStr = sStringToMatch Then
               FindIndex = i
               Exit Function
            End If
         Next i

      End If

'     if we get here a match has not been found.
      FindIndex = -1

   Else

'     control is disabled; return -1.
      FindIndex = -1

   End If

End Function

Public Function MouseOverIndex(ByVal YPos As Single) As Long

'*************************************************************************
'* .MouseOverIndex method.  Returns the .List() index of the item the    *
'* mouse pointer is over, based on the mouse y-coordinate and the first  *
'* displayed item's index.  Is also used internally by other usercontrol *
'* routines.  Returns -1 if mouse cursor is not over populated part of   *
'* the list or is not in list portion of control (e.g. over scrollbar).  *
'*************************************************************************

   Dim DisplayIndex As Long     ' display position in listbox.

   If m_Enabled Then

'     determine the display order index based on mouse Y coordinate.
      DisplayIndex = GetDisplayOrderIndex(YPos)

'     add that index to the index of the first displayed value.
      MouseOverIndex = DisplayRange.FirstListItem + DisplayIndex

'     safety net for below last item in list, no items at all, and mouse not in list portion.
'     the "If Not ScrollFlag" ensures that a -1 is not returned when drag scrolling.
      If Not ScrollFlag And (MouseOverIndex > m_ListCount - 1 Or Not IsInList(MouseX, MouseY)) Then
         MouseOverIndex = -1
      End If

   Else

'     control is disabled; return -1.
      MouseOverIndex = -1

   End If

End Function

Private Function GetDisplayOrderIndex(ByVal YPos As Single) As Long

'*************************************************************************
'* determines the display (YCoords array) index of the desired displayed *
'* list item, given the mouse Y coordinate.  Helper function for the     *
'* MouseOverIndex method function.                                       *
'*************************************************************************

   Dim iLBound As Long      ' lower bound of list array portion currently being searched.
   Dim iUBound As Long      ' upper bound of list array portion currently being searched.
   Dim iMiddle As Long      ' middle of list array portion currently being searched.
   Dim Done    As Boolean   ' while loop finished flag.

   iLBound = LBound(YCoords)
   iUBound = MaxDisplayItems - 2

   Done = False
   While Not Done
      iMiddle = (iLBound + iUBound) / 2
      If YPos >= YCoords(iMiddle) And YPos < YCoords(iMiddle + 1) Then
         GetDisplayOrderIndex = iMiddle
         Done = True
      ElseIf iLBound > iUBound Then
         GetDisplayOrderIndex = iLBound
         Done = True
      Else
         If YCoords(iMiddle) < YPos Then
            iLBound = iMiddle + 1
         Else
            iUBound = iMiddle - 1
         End If
      End If
   Wend

End Function

Public Sub Refresh()

'*************************************************************************
'* allows user to refresh the graphics of the control if ever necessary. *
'*************************************************************************

   If m_Enabled Then
      UserControl.Refresh
   End If

End Sub

Public Sub AddImage(ByVal ImagePath As String)

'*************************************************************************
'* .AddImage method.  Allows user to add an image to the list of images  *
'* that can be displayed next to listitems.  Adapted from a routine in   *
'* Jim Jose's "McImageList" submission at PSC, txtCodeId=62417.  Thanks  *
'* to Jim.  Note:  It is up to the programmer to keep track of the image *
'* order in project code so that it is known which image is which.       *
'*************************************************************************

   Dim mArray() As StdPicture    ' temporary image array.
   Dim NewImage As StdPicture    ' the image to add, loaded using ImagePath parameter.
   Dim i        As Long          ' loop variable.

   Set NewImage = LoadPicture(ImagePath)

   If ImageCount = 0 Then
      ReDim Images(0)
      Set Images(0) = NewImage
      ImageCount = 1
   Else
      mArray = Images
      Erase Images
      ImageCount = ImageCount + 1
      ReDim Images(0 To ImageCount - 1)
      For i = 0 To ImageCount - 2
         Set Images(i) = mArray(i)
      Next i
      Set Images(ImageCount - 1) = NewImage
   End If

End Sub

Public Sub AddItem(ByVal ItemToAdd As String, Optional ByVal Index As Long = -1)

'*************************************************************************
'* .AddItem method - adds an item/ItemData item to the list, optionally  *
'* to the given index.  If the Sorted property is True, adds to the list *
'* in the appropriate spot.  If Index parameter is supplied, this takes  *
'* precedence over Sorted property (as is the case with the intrinsic VB *
'* textbox).  If Sorted is False and no index is supplied, appends item  *
'* to end of list.                                                       *
'*************************************************************************

   If Not m_Enabled Then
      Exit Sub
   End If

   m_ListCount = m_ListCount + 1
   RecalculateThumbHeight = True    ' we need to recalculate thumb height since size of list has changed.

'  add the item to the list and set the .NewItem property.
   If Index = -1 Then
'     if the index is -1 (i.e. not supplied), and .Sorted = False, just append the item.
      If Not m_Sorted Then
         AddToList ItemToAdd, m_ListCount - 1
         AddToLongPropertyArray ItemDataArray(), m_ListCount - 1, 0
         AddToSelected m_ListCount - 1
         If m_ShowItemImages Then
            AddToLongPropertyArray ImageIndexArray(), m_ListCount - 1, -1
         End If
         m_NewIndex = m_ListCount - 1
      Else
'        if the index is -1 (i.e. not supplied), and .Sorted = True, insert item alphabetically.
'        if .SortAsNumeric property is .True, treat strings as numbers.
         If Not m_SortAsNumeric Then
            Index = AddToSortedList(ItemToAdd)
         Else
            Index = AddToSortedListAsNumeric(ItemToAdd)
         End If
         AddToLongPropertyArray ItemDataArray(), Index, 0
         AddToSelected Index
         If m_ShowItemImages Then
            AddToLongPropertyArray ImageIndexArray(), Index, -1
         End If
         m_NewIndex = Index
      End If
   Else
'     index has been supplied; insert the item at the indicated position.
'     NOTE: A supplied index overrides the .Sorted property (by design).
'     Therefore it is the programmer's responsibility to remember this fact and
'     to realize proper sort order will be lost if an index is supplied when .Sorted
'     is True.  Search using the .FindIndex method is also adversely affected.
      AddToList ItemToAdd, Index
      AddToLongPropertyArray ItemDataArray, Index, 0
      AddToSelected Index
      If m_ShowItemImages Then
         AddToLongPropertyArray ImageIndexArray(), Index, -1
      End If
      m_NewIndex = Index
   End If

'  if the item is not just being appended to the list, we may have to adjust the following
'  variables if they point to items that come on or after the supplied index of added item.
'  Ignored if .ListIndex = -1 or 0 and no item clicked on (i.e. list is newly initialized).
   If Not (m_ListIndex <= 0 And m_SelCount = 0) Then
      If m_ListIndex >= Index Then
         m_ListIndex = m_ListIndex + 1
      End If
      If ItemWithFocus >= Index Then
         ItemWithFocus = ItemWithFocus + 1
      End If
      If LastSelectedItem >= Index Then
         LastSelectedItem = LastSelectedItem + 1
      End If
   End If

'  since there's at least one item in the list now, activate the first display
'  item index.  The last display item index is calculated in the DrawText routine.
   If DisplayRange.FirstListItem = -1 Then
      DisplayRange.FirstListItem = 0
   End If

'  the .RedrawFlag property is used to postpone redrawing if large numbers
'  of items are added at one time.  .RemoveItem method also uses this property.
   If m_RedrawFlag Then
      DisplayList
   End If

End Sub

Private Function AddToSortedList(ByVal sToAdd As String) As Long

'*************************************************************************
'* helper routine for the .AddItem method.  Places a new list entry      *
'* into the proper place in an already-sorted listbox array.             *
'*************************************************************************

   Dim Lower As Long     ' lower bound of list array portion currently being searched.
   Dim Middle As Long    ' middle of list array portion currently being searched.
   Dim Upper As Long     ' upper bound of list array portion currently being searched.

   If m_ListCount = 1 Then ' already incremented m_listcount in AddItem method so this means empty list.
      AddToSortedList = 0
      AddToList sToAdd, 0
      Exit Function
   End If

   Lower = LBound(ListArray)
   Upper = UBound(ListArray) - 1

'  find the appropriate index to place new list item into.
   While (True)
      Middle = (Lower + Upper) / 2
      If ListArray(Middle) = sToAdd Then
         AddToSortedList = Middle
         AddToList sToAdd, Middle
         Exit Function
      ElseIf Lower > Upper Then
         AddToSortedList = Lower
         AddToList sToAdd, Lower
         Exit Function
      Else
         If ListArray(Middle) < sToAdd Then
            Lower = Middle + 1
         Else
            Upper = Middle - 1
         End If
      End If
   Wend

End Function

Private Function AddToSortedListAsNumeric(ByVal sToAdd As String) As Long

'*************************************************************************
'* helper routine for the .AddItem method.  Places a new list entry      *
'* into the proper place in an already-sorted listbox array as a numeric *
'* sort when .SortAsNumeric property is True.  Suggestion by Jeff Mayes. *
'* Note:  This is considerably slower than normal string comparison.     *
'* However, it should still load much faster than a standard VB listbox. *
'*************************************************************************

   Dim Lower    As Long     ' lower bound of list array portion currently being searched.
   Dim Middle   As Long     ' middle of list array portion currently being searched.
   Dim Upper    As Long     ' upper bound of list array portion currently being searched.
   Dim nToAdd   As Double   ' must account for large or non-whole numbers.
   Dim nCompare As Double   ' treat values already in list as doubles also.

   nToAdd = Val(sToAdd)

   If m_ListCount = 1 Then ' already incremented m_listcount in AddItem method so this means empty list.
      AddToSortedListAsNumeric = 0
      AddToList sToAdd, 0
      Exit Function
   End If

   Lower = LBound(ListArray)
   Upper = UBound(ListArray) - 1

'  find the appropriate index to place new list item into.
   While (True)
      Middle = (Lower + Upper) / 2
      nCompare = Val(ListArray(Middle))
      If nCompare = nToAdd Then
         AddToSortedListAsNumeric = Middle
         AddToList sToAdd, Middle
         Exit Function
      ElseIf Lower > Upper Then
         AddToSortedListAsNumeric = Lower
         AddToList sToAdd, Lower
         Exit Function
      Else
         If nCompare < nToAdd Then
            Lower = Middle + 1
         Else
            Upper = Middle - 1
         End If
      End If
   Wend

End Function

Private Sub AddToList(ByVal sStringToAdd As String, Optional ByVal iPos As Long = -1)

'*************************************************************************
'* helper routine for the .AddItem method.  Places a new list entry      *
'* into the specified position in a listbox array. If no index is spec-  *
'* ified, item is appended to the end of the list.  Modification of a    *
'* routine by Philippe Lord.                                             *
'*************************************************************************

   Dim iUBound As Long    ' upper bound of the List array.
   Dim iTemp   As Long    ' don't really know :)

   iUBound = UBound(ListArray)

'  if array is empty.
   If iUBound = -1 Then
      ReDim ListArray(0)
      ListArray(0) = sStringToAdd
      Exit Sub
   End If

'  if adding at the end.
   If (iPos > iUBound) Or (iPos = -1) Then
      ReDim Preserve ListArray(iUBound + 1)
      ListArray(iUBound + 1) = sStringToAdd
      Exit Sub
   End If

'  in case a negative less than -1 is erroneously passed.
   If iPos < 0 Then
      iPos = 0
   End If

'  increase size of array by one element.
   iUBound = iUBound + 1
   ReDim Preserve ListArray(iUBound)

   CopyMemory ByVal VarPtr(ListArray(iPos + 1)), ByVal VarPtr(ListArray(iPos)), (iUBound - iPos) * 4

   iTemp = 0 ' view this as String(4, Chr(0)) or a NULL value
   CopyMemory ByVal VarPtr(ListArray(iPos)), iTemp, 4

   ListArray(iPos) = sStringToAdd

End Sub

Private Sub AddToLongPropertyArray(PropArray() As Long, ByVal iPos As Long, ByVal InitialValue As Long)

'*************************************************************************
'* helper routine for the .AddItem method.  places a new ItemData entry  *
'* into the specified position in an ItemData or ImageIndex long array.  *
'* Modification of a routine by Philippe Lord.                           *
'*************************************************************************

   Dim iUBound As Long    ' upper bound of the ItemData or ImageIndex array.

   iUBound = UBound(PropArray)

'  if array is empty.
   If iUBound = -1 Then
      ReDim PropArray(0)
      PropArray(0) = InitialValue
      Exit Sub
   End If

'  if adding at the end.
   If (iPos > iUBound) Or (iPos = -1) Then
      ReDim Preserve PropArray(iUBound + 1)
      PropArray(iUBound + 1) = InitialValue
      Exit Sub
   End If

'  in case a negative index is erroneously passed.
   If iPos < 0 Then
      iPos = 0
   End If

'  increase size of array by one element.
   iUBound = iUBound + 1
   ReDim Preserve PropArray(iUBound)

   CopyMemory PropArray(iPos + 1), PropArray(iPos), (iUBound - LBound(PropArray) - iPos) * Len(PropArray(iPos))
   PropArray(iPos) = InitialValue

End Sub

Private Sub AddToSelected(ByVal iPos As Long)

'*************************************************************************
'* helper routine for the .AddItem method.  Adds a new .Selected() entry *
'* into the specified position in a Selected boolean array. Modification *
'* of a routine by Philippe Lord.                                        *
'*************************************************************************

   Dim iUBound As Long    ' upper bound of the Selected array.

   iUBound = UBound(SelectedArray)

   If iUBound = -1 Then
      ReDim SelectedArray(0)
      SelectedArray(0) = False
      Exit Sub
   End If

' if adding at the end.
   If (iPos > iUBound) Or (iPos = -1) Then
      ReDim Preserve SelectedArray(iUBound + 1)
      SelectedArray(iUBound + 1) = False
      Exit Sub
   End If

'  in case a negative is erroneously passed.
   If iPos < 0 Then
      iPos = 0
   End If

'  increase size of array by one element.
   iUBound = iUBound + 1
   ReDim Preserve SelectedArray(iUBound)

   CopyMemory SelectedArray(iPos + 1), SelectedArray(iPos), (iUBound - LBound(SelectedArray) - iPos) * Len(SelectedArray(iPos))
   SelectedArray(iPos) = False

End Sub

Public Sub RemoveItem(ByVal Index As Long)

'*************************************************************************
'* .RemoveItem method - removes the specified item from the List,        *
'* ItemData, Selected, and ImageIndex property arrays.                   *
'*************************************************************************

   If Not m_Enabled Then
      Exit Sub
   End If

   If m_ListCount > 0 Then

      If Index >= LBound(ListArray) And Index <= UBound(ListArray) Then

'        reduce the .ListCount property variable.
         m_ListCount = m_ListCount - 1
         RecalculateThumbHeight = True   ' we need to recalculate thumb height since size of list has changed.

'        if the item to remove is selected, decrease the m_SelCount property variable.
         If SelectedArray(Index) Then
            m_SelCount = m_SelCount - 1
         End If

'        remove from the property arrays.
         RemoveFromListArray Index
         RemoveFromLongPropertyArray ItemDataArray(), Index ' RemoveFromItemDataArray Index
         RemoveFromSelectedArray Index
         RemoveFromLongPropertyArray ImageIndexArray(), Index ' RemoveFromImageIndexArray Index

'        adjust the .ListIndex property variable.
'        if .ListIndex points to the item that was just removed, set it to cleared status per mode.
         If m_ListIndex = Index Then
            If m_Style = [CheckBox] Or m_MultiSelect <> vbMultiSelectNone Then
               m_ListIndex = 0
            Else
               m_ListIndex = -1
            End If
         Else
'           if .ListIndex comes after deleted item we must decrement
'           it to reflect the ListIndex item's new array position.
            If m_ListIndex > Index Then
               m_ListIndex = m_ListIndex - 1
            End If
         End If

'        adjust the LastSelectedItem internal variable.
'        if it points to the item that was just removed, clear it.
         If LastSelectedItem = Index Then
            LastSelectedItem = -1
         Else
'           if it comes after deleted item we must decrement
'           it to reflect the item's new array position.
            If LastSelectedItem > Index Then
               LastSelectedItem = LastSelectedItem - 1
            End If
         End If

'        adjust the ItemWithFocus internal variable.
'        if it points to the item that was just removed, clear it.
         If ItemWithFocus = Index Then
            ItemWithFocus = 0
         Else
'           if it comes after deleted item we must decrement
'           it to reflect the item's new array position.
            If ItemWithFocus > Index Then
               ItemWithFocus = ItemWithFocus - 1
            End If
         End If

'        since an item has been removed, we must set the .NewIndex property to -1.
         m_NewIndex = -1

'        as with the .AddItem method, the .RedrawFlag property can be used to postpone
'        redrawing the control if a large number of items are being removed from the list.
         If m_RedrawFlag Then
            DisplayList
         End If

      End If

   End If

End Sub

Private Sub RemoveFromListArray(ByVal iPos As Long)

'*************************************************************************
'* helper routine for the RemoveItem method - removes the specified item *
'* from the List array.  Modification of a routine by Philippe Lord.     *
'*************************************************************************

   Dim iLBound As Long     ' lower bound of List array.
   Dim iUBound As Long     ' upper bound of List array.
   Dim iTemp   As Long     ' pointer to address of List array element.

   iLBound = LBound(ListArray)
   iUBound = UBound(ListArray)

'  if we only have one element in array.
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then
      Erase ListArray
      ReDim ListArray(0)
      Exit Sub
   End If

'  if invalid iPos - might not need 1st two checks now.
   If (iPos > iUBound) Or (iPos = -1) Then
      iPos = iUBound
   End If
   If iPos < iLBound Then
      iPos = iLBound
   End If
   If iPos = iUBound Then
      ReDim Preserve ListArray(iUBound - 1)
      Exit Sub
   End If

   iTemp = StrPtr(ListArray(iPos))
   CopyMemory ByVal VarPtr(ListArray(iPos)), ByVal VarPtr(ListArray(iPos + 1)), (iUBound - iPos) * 4

'  do this to have VB deallocate the string; avoids memory leaks.
   CopyMemory ByVal VarPtr(ListArray(iUBound)), iTemp, 4

   ReDim Preserve ListArray(iUBound - 1)

End Sub

Private Sub RemoveFromLongPropertyArray(PropArray() As Long, ByVal iPos As Long)

'*************************************************************************
'* helper routine for the .RemoveItem method - removes specified item    *
'* from the ItemData or ImageIndex arrays.  Modification of a routine by *
'* Philippe Lord.                                                        *
'*************************************************************************

   Dim iLBound As Long   ' lower bound of ItemData or ImageIndex array.
   Dim iUBound As Long   ' upper bound of ItemData or ImageIndex array.

   iLBound = LBound(PropArray)
   iUBound = UBound(PropArray)

'  if we only have one element in array.
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then
      Erase PropArray
      ReDim PropArray(0)
      Exit Sub
   End If

'  if invalid iPos.
   If (iPos > iUBound) Or (iPos = -1) Then
      iPos = iUBound
   End If
   If iPos < iLBound Then
      iPos = iLBound
   End If
   If iPos = iUBound Then
      ReDim Preserve PropArray(iUBound - 1)
      Exit Sub
   End If

   CopyMemory PropArray(iPos), PropArray(iPos + 1), (iUBound - iLBound - iPos) * Len(PropArray(iPos))

   ReDim Preserve PropArray(iUBound - 1)

End Sub

Private Sub RemoveFromSelectedArray(ByVal iPos As Long)

'*************************************************************************
'* helper routine for the .RemoveItem method - removes specified item    *
'* from the Selected array.  Modification of a routine by Philippe Lord. *
'*************************************************************************

   Dim iLBound As Long    ' lower bound of .Selected() array.
   Dim iUBound As Long    ' upper bound of .Selected() array.

   iLBound = LBound(SelectedArray)
   iUBound = UBound(SelectedArray)

'  if we only have one element in array.
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then
      Erase SelectedArray
      ReDim SelectedArray(0)
      Exit Sub
   End If

'  if invalid iPos.
   If (iPos > iUBound) Or (iPos = -1) Then
      iPos = iUBound
   End If
   If iPos < iLBound Then
      iPos = iLBound
   End If
   If iPos = iUBound Then
      ReDim Preserve SelectedArray(iUBound - 1)
      Exit Sub
   End If

   CopyMemory SelectedArray(iPos), SelectedArray(iPos + 1), (iUBound - iLBound - iPos) * Len(SelectedArray(iPos))

   ReDim Preserve SelectedArray(iUBound - 1)

End Sub

Public Sub Clear()

'*************************************************************************
'* the .Clear method for the listbox; removes all entries.               *
'*************************************************************************

   If Not m_Enabled Then
      Exit Sub
   End If

'  re-initialize the four property arrays.
   ReDim ListArray(0)
   ReDim ItemDataArray(0)
   ReDim SelectedArray(0)
   ReDim ImageIndexArray(0)

'  set the appropriate property and internal values to initialized state.
   m_SelCount = 0
   m_ListCount = 0
   LastSelectedItem = -1
   ItemWithFocus = 0
   DisplayRange.FirstListItem = -1
   DisplayRange.LastListItem = -1
   m_RedrawFlag = True
   VerticalScrollBarActive = False
   RecalculateThumbHeight = True    ' we need to recalculate thumb height since size of list has changed.

'  in CheckBox, MultiSelect Simple and MultiSelect Extended modes, the
'  ListIndex property is 0 in a cleared list but -1 in MultiSelect None mode.
   If m_Style = [CheckBox] Or m_MultiSelect <> vbMultiSelectNone Then
      m_ListIndex = 0
   Else
      m_ListIndex = -1
   End If

'  since the listbox has been cleared, we must set the .NewIndex property to -1.
   m_NewIndex = -1

'  redraw the background (and border, if picture) onto the usercontrol DC.
   SetBackGround
   If IsPictureThere(m_ActivePicture) Then
      CreateBorder
   End If
   UserControl.Refresh

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Miscellaneous ListBox Helper Functions >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ProcessSelectedItem()

'*************************************************************************
'* handles pointers and painting of a newly selected listbox item        *
'* according to the .MultiSelect property (Single, Simple, Extended).    *
'*************************************************************************

   Select Case m_MultiSelect

      Case vbMultiSelectNone
         ProcessSelected_MultiSelectNone

      Case vbMultiSelectSimple
         ProcessSelected_MultiSelectSimple

      Case vbMultiSelectExtended
         ProcessSelected_MultiSelectExtended

   End Select

End Sub

Private Sub ProcessSelected_MultiSelectNone()

'*************************************************************************
'* performs operations to select an item in MultiSelect None mode.       *
'*************************************************************************

'  reinitialize the Selected array to all False. (Remember, no multiselect here.)
   SetSelectedArrayRange 0, m_ListCount - 1, False

'  set the .ListIndex property to the index of the selected item.
   m_ListIndex = LastSelectedItem

'  set the clicked item's selected status to True.
   SelectedArray(LastSelectedItem) = True
   ItemWithFocus = LastSelectedItem
   m_SelCount = 1

   DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessSelected_MultiSelectSimple()

'*************************************************************************
'* performs operations to select an item in MultiSelect Simple mode.     *
'*************************************************************************

'  set the clicked item's selected status to True.
   SelectedArray(LastSelectedItem) = True
   ItemWithFocus = LastSelectedItem ' can take out?

'  in Simple mode, the .ListIndex property is ALWAYS the index of the focused item, selected or not.
   m_ListIndex = ItemWithFocus
   m_SelCount = m_SelCount + 1

   DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessSelected_MultiSelectExtended(Optional ByVal SelCountFlag As Boolean = True)

'*************************************************************************
'* performs operations to select an item in MultiSelect Extended mode.   *
'*************************************************************************

'  set the .ListIndex property to the index of the selected item.
   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus

'  set the clicked item's selected status to True.
   SelectedArray(LastSelectedItem) = True
   If SelCountFlag Then
      m_SelCount = m_SelCount + 1
   End If

   DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessContinuousScroll(ScrollAmount As Long)

'*************************************************************************
'* performs a continuous vertical scroll, scrolling by specified amount. *
'* this routine handles all vertical scrollbar continuous scrolling,     *
'* (up arrow, down arrow, mouse above/below listbox, and trackbar).      *
'*************************************************************************

   Dim OriginalTickCount As Long   ' comparison tick count for calculating elapsed time.
   Dim CurrentTickCount  As Long   ' current time (tick count).
   Dim LastItemIndex     As Long   ' last item for down scroll, first item for up scroll.

'  based on scroll direction, determine the end item to check for during scroll.
   LastItemIndex = IIf(ScrollAmount > 0, m_ListCount - 1, 0)

'  create a preliminary delay before list starts scrolling, to give the user time to unclick the
'  mouse button and prevent scrolling.  If user clicked in list area and is scrolling by dragging
'  the mouse above or below the listbox, ScrollFlag will be True and this initial delay is ignored.
   If Not ScrollFlag Then
      OriginalTickCount = GetTickCount
      CurrentTickCount = OriginalTickCount
      While MouseAction <> MOUSE_NOACTION And CurrentTickCount - OriginalTickCount < INITIAL_SCROLL_DELAY
         CurrentTickCount = GetTickCount
         DoEvents
      Wend
   End If

'  if the DoEvents from the above delay loop did not reveal a MouseUp event, the display-scrolling
'  loop below can execute (MouseAction will be <> MOUSE_NOACTION).  Loop until the mouse button is
'  unclicked or the top or bottom of the list has been reached, depending on scroll direction.
'  Scroll interval (SCROLL_TICKCOUNT) is 50 milliseconds.
   OriginalTickCount = GetTickCount
   While MouseAction <> MOUSE_NOACTION And (Not InDisplayedItemRange(LastItemIndex))

'     allow an opportunity for a MouseUp event to stop the scrolling.
      DoEvents
'     keeps cpu usage from maxing out.  Thanks to Mike Douglas for the tip.
      Sleep 25

'     this 'If' statement allows control to process trackbar scrolling like the VB listbox -
'     scrolling will continue until thumb is under mouse cursor, at which point scrolling stops.
      If ScrollFlag Or MouseLocation = OVER_VTRACKBAR Or MouseLocation = OVER_UPBUTTON Or MouseLocation = OVER_DOWNBUTTON Then

'        get the current time.
         CurrentTickCount = GetTickCount

'        perform a scroll if at least SCROLL_TICKCOUNT milliseconds have passed.
         If CurrentTickCount - OriginalTickCount >= SCROLL_TICKCOUNT Then

'           adjust the display range up or down according to scroll direction.
            DisplayRange.FirstListItem = DisplayRange.FirstListItem + ScrollAmount
            If DisplayRange.FirstListItem < 0 Then
               DisplayRange.FirstListItem = 0
            End If
            DisplayRange.LastListItem = DisplayRange.LastListItem + ScrollAmount
            If DisplayRange.LastListItem > m_ListCount - 1 Then
               DisplayRange.LastListItem = m_ListCount - 1
            End If

'           if we're page scrolling (i.e. scrolling using trackbar), make sure end of list is
'           displayed correctly with the last list item at the physical bottom of the control.
            If Abs(ScrollAmount) > 1 And DisplayRange.FirstListItem + MaxDisplayItems - 1 > m_ListCount Then
               DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems + 1
            End If

'           if page scrolling via trackbar, and mouse pointer moves out of section of trackbar
'           that determined scroll direction (e.g. mouse is moved below thumb when scrolling up
'           due to mouse being originally clicked in track above thumb) stop scrolling.
            If Abs(ScrollAmount) > 1 Then
               If MouseCursorIsAboveThumb(MouseY) And MouseAction = MOUSE_DOWNED_IN_LOWERTRACK Then
                  Exit Sub
               End If
               If Not MouseCursorIsAboveThumb(MouseY) And MouseAction = MOUSE_DOWNED_IN_UPPERTRACK Then
                  Exit Sub
               End If
            End If

'           if we're scrolling by clicking on the list and then dragging the mouse above
'           or below the listbox, the top or bottom displayed list item (depending on
'           scrolling direction) always has the highlight gradient (unless in MultiSelect
'           Simple mode, where just the focus rectangle is used).
            If ScrollFlag Then
               If ScrollAmount > 0 Then   ' if scrolling down (mouse dragged below listbox)...
                  If m_MultiSelect <> vbMultiSelectSimple Then
                     LastSelectedItem = DisplayRange.LastListItem
                  End If
               Else                       ' if scrolling up (mouse dragged above listbox)...
                  If m_MultiSelect <> vbMultiSelectSimple Then
                     LastSelectedItem = DisplayRange.FirstListItem
                  End If
               End If
               ProcessMouseMoveItemSelection
            End If

'           time to redisplay the list after display range adjustments.
            DisplayList

'           store the current time and wait for the next SCROLL_TICKCOUNT milliseconds.
            OriginalTickCount = GetTickCount

         End If

      End If

   Wend

End Sub

Private Sub SetSelectedArrayRange(ByVal FirstValue As Long, ByVal LastValue As Long, ByVal bSelectedStatus As Boolean)

'*************************************************************************
'* this procedure sets the given range of elements in the SelectedArray  *
'* array to either True or False using the FillMemory API.  It is used   *
'* when .MultiSelect is Extended and extremely large ranges of list      *
'* items' Selected status must be set very quickly (for example, when    *
'* 25000 list items must all be selected at once due clicking the first  *
'* item and then pressing Shift-End).  It is also used to instantly set  *
'* the entire array to unselected (False).                               *
'*************************************************************************

   Dim Temp As Long     ' swap temporary variable.

'  the way I do things in this control, FirstValue might be greater than LastValue.
'  For example, when doing a Shift-PageUp, this will be the case.  To correctly use
'  the FillMemory API,  we must swap FirstValue and LastValue in these circumstances.
   If FirstValue > LastValue Then
      Temp = FirstValue
      FirstValue = LastValue
      LastValue = Temp
   End If

   FillMemory SelectedArray(FirstValue), 2 * (LastValue - FirstValue + 1), bSelectedStatus

End Sub

Private Sub AdjustDisplayRange()

'*************************************************************************
'* the following code helps emulate the vb listbox when the focused item *
'* is out of displayed range (such as when scrollbar is used to navigate *
'* up or down the list), and an arrow key, page key, etc. is pressed.    *
'* MorphListBox display range is adjusted according to vb listbox rules. *
'*************************************************************************

   If ItemWithFocus < DisplayRange.FirstListItem Then
'     if the list item with the focus is above the first displayed item, the display
'     will adjust so that the focused item is at the top of the displayed range.
      DisplayRange.FirstListItem = ItemWithFocus
      If DisplayRange.FirstListItem + MaxDisplayItems - 1 <= m_ListCount Then
         DisplayRange.LastListItem = ItemWithFocus + MaxDisplayItems - 1
      Else
         DisplayRange.LastListItem = m_ListCount - 1
      End If
      DisplayList
   Else
      If ItemWithFocus > DisplayRange.LastListItem Then
'        if the list item with the focus is below the last displayed item, the display
'        will adjust so that the focused item is at the bottom of the displayed range.
         DisplayRange.LastListItem = ItemWithFocus
         If DisplayRange.LastListItem - MaxDisplayItems + 1 >= 0 Then
            DisplayRange.FirstListItem = DisplayRange.LastListItem - MaxDisplayItems + 1
         Else
            DisplayRange.FirstListItem = 0
         End If
         DisplayList
      End If
   End If

End Sub

Private Function GetDisplayIndexFromArrayIndex(ArrayIndex As Long) As Long

'*************************************************************************
'* given the item's array index, returns the display (YCoords) index,    *
'* or returns -1 if the item is not in the display range.                *
'*************************************************************************

   Dim iLBound      As Long        ' lower bound of display range currently being searched.
   Dim iUBound      As Long        ' upper bound of display range currently being searched.
   Dim iMiddle      As Long        ' middle of display range currently being searched.
   Dim CompareIndex As Long        ' current display range index currently being examined.

   iLBound = 0
   iUBound = DisplayRange.LastListItem - DisplayRange.FirstListItem + 1

   Do
      iMiddle = (iLBound + iUBound) \ 2
      CompareIndex = DisplayRange.FirstListItem + iMiddle - 1
      If CompareIndex = ArrayIndex Then
         GetDisplayIndexFromArrayIndex = iMiddle - 1
         Exit Function
      ElseIf CompareIndex < ArrayIndex Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If
   Loop Until iLBound > iUBound

   GetDisplayIndexFromArrayIndex = -1

End Function

Private Sub CalculateSelCount()

'*************************************************************************
'* I don't like recalculating m_SelCount with this brute force method.   *
'* But (for example) in MultiSelect Extended mode using the Shift-Home   *
'* or Shift-End  keys, with other items or groups of items possibly hav- *
'* ing been selected in other parts of the list (or even the part of the *
'* list affected by a Shift-Home/Shift-End), this is the most logical    *
'* way.  Still very fast though.                                         *
'*************************************************************************

   Dim i      As Long    ' loop variable.
   Dim EndVal As Long    ' last element of the .Selected array.

   m_SelCount = 0
   EndVal = m_ListCount - 1

   For i = 0 To EndVal
'     remember that True = -1 and False = 0.  So the absolute value
'     of the sum of the elements of the SelectedArray can be used.
      m_SelCount = m_SelCount + SelectedArray(i)
   Next i

   m_SelCount = Abs(m_SelCount)

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Vertical ScrollBar >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub DisplayVerticalScrollBar(Optional vThumbYPos As Single = -1)

'**************************************************************************
'* displays the vertical scrollbar.                                       *
'**************************************************************************

   DisplayVerticalTrackBar
   DisplayTrackBarButton UPBUTTON
   DisplayTrackBarButton DOWNBUTTON
   DisplayVerticalThumb vThumbYPos

'  get the information so that mouse cursor location in scrollbar can be determined.
   GetVScrollbarLocationInfo

End Sub

Private Sub GetVScrollbarLocationInfo()

'*************************************************************************
'* gets position info for all parts of vertical scrollbar except the     *
'* thumb, which is calculated on-the-fly when the thumb is drawn.        *
'*************************************************************************

   vScrollBarLocation.UpButtonLocation.Top = m_BorderWidth
   vScrollBarLocation.UpButtonLocation.Left = ScaleWidth - ScrollBarButtonWidth
   vScrollBarLocation.UpButtonLocation.Bottom = vScrollBarLocation.UpButtonLocation.Top + ScrollBarButtonHeight - 1
   vScrollBarLocation.UpButtonLocation.Right = vScrollBarLocation.UpButtonLocation.Left + ScrollBarButtonWidth - 1

   vScrollBarLocation.DownButtonLocation.Top = ScaleHeight - m_BorderWidth - ScrollBarButtonHeight
   vScrollBarLocation.DownButtonLocation.Left = ScaleWidth - ScrollBarButtonWidth
   vScrollBarLocation.DownButtonLocation.Bottom = vScrollBarLocation.DownButtonLocation.Top + ScrollBarButtonHeight - 1
   vScrollBarLocation.DownButtonLocation.Right = vScrollBarLocation.DownButtonLocation.Left + ScrollBarButtonWidth - 1

   vScrollBarLocation.ScrollTrackLocation.Top = m_BorderWidth + ScrollBarButtonHeight
   vScrollBarLocation.ScrollTrackLocation.Left = ScaleWidth - ScrollBarButtonWidth - m_BorderWidth
   vScrollBarLocation.ScrollTrackLocation.Bottom = ScaleHeight - m_BorderWidth - ScrollBarButtonHeight - 1
   vScrollBarLocation.ScrollTrackLocation.Right = vScrollBarLocation.ScrollTrackLocation.Left + ScrollBarButtonWidth - 1

End Sub

Private Function MouseCursorIsAboveThumb(YPos As Single) As Boolean

'*************************************************************************
'* when mouse is clicked on vertical scroll bar trackbar, need to deter- *
'* mine if the mouse is above or below the scroll thumb so a page up or  *
'* page down can be performed.                                           *
'*************************************************************************

   If YPos < vScrollBarLocation.ScrollThumbLocation.Top Then
      MouseCursorIsAboveThumb = True
   End If

End Function

Private Sub DisplayVerticalThumb(Optional Y As Single = -1)

'*************************************************************************
'* displays the vertical scrollbar's thumb scroller.                     *
'*************************************************************************

   Dim YPos As Long ' the y position for the top of the thumb.

'  obtain the height of the scroller thumb.  Save processing
'  time by only recalculating when list size has changed.
   If RecalculateThumbHeight Then
      vThumbHeight = CalculateVScrollThumbHeight
   End If

'  get the top y pos of the vertical scroller thumb.
   If Y = -1 Then
'     if no cursor position in thumb is passed, then just calculate
'     the thumb's top y coordinate based on where we are in the list.
      YPos = ThumbYPos ' calculated in DisplayVerticalTrackbar.
   Else
'     otherwise, calculate based on the current mouse pos and where
'     it is in the thumb.  This is only done when dragging the thumb.
      YPos = MouseY - Y
'     these 'if' statements ensure the thumb stays between the up and down buttons.
      If YPos + vThumbHeight > vScrollBarLocation.DownButtonLocation.Top Then
         YPos = vScrollBarLocation.DownButtonLocation.Top - vThumbHeight
      End If
      If YPos < vScrollBarLocation.UpButtonLocation.Bottom Then
         YPos = vScrollBarLocation.UpButtonLocation.Bottom + 1
      End If
   End If

'  define the vertical thumb's rectangle coordinates.
   vScrollBarLocation.ScrollThumbLocation.Top = YPos
   vScrollBarLocation.ScrollThumbLocation.Left = ScaleWidth - ScrollBarButtonWidth - m_BorderWidth
   vScrollBarLocation.ScrollThumbLocation.Bottom = YPos + vThumbHeight - 1
   vScrollBarLocation.ScrollThumbLocation.Right = vScrollBarLocation.ScrollThumbLocation.Left + ScrollBarButtonWidth - 1

'  display the thumb.
   Call StretchDIBits(hdc, _
                      vScrollBarLocation.ScrollThumbLocation.Left, YPos, _
                      ScrollBarButtonWidth, vThumbHeight, _
                      0, 0, _
                      ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, vThumbHeight, _
                      vThumblBits(0), vThumbuBIH, _
                      DIB_RGB_COLORS, _
                      vbSrcCopy)

'  draw the thumb's border.
   DisplayVerticalThumbBorder

End Sub

Private Sub DisplayVerticalThumbBorder()

'*************************************************************************
'* draws a one-pixel wide border around the vertical scrollbar thumb.    *
'*************************************************************************

   DrawRectangle vScrollBarLocation.ScrollThumbLocation.Left + 1, _
                 vScrollBarLocation.ScrollThumbLocation.Top, _
                 vScrollBarLocation.ScrollThumbLocation.Right, _
                 vScrollBarLocation.ScrollThumbLocation.Bottom, _
                 m_ActiveThumbBorderColor

End Sub

Private Function VerticalThumbY() As Long

'*************************************************************************
'* determines the y coordinate of the top of the vertical scrollbar's    *
'* thumb, so thumb will be displayed in the right place in the track.    *
'*************************************************************************

   Dim PixelsPerScroll       As Single  ' how many pixels involved in scrolling one list item.
   Dim NumClicks             As Long    ' basically, how many clicks it takes to get to list end.

'  calculate the thumb middle pixel's range of motion in the track.
   vThumbRange.Top = m_BorderWidth + ScrollBarButtonHeight + (vThumbHeight / 2) + 1
   vThumbRange.Bottom = ScaleHeight - m_BorderWidth - ScrollBarButtonHeight - (vThumbHeight / 2) + 1

'  if we're at the very top or bottom of the list, this is easy.
   If DisplayRange.FirstListItem = 0 Then
      VerticalThumbY = m_BorderWidth + ScrollBarButtonHeight
      Exit Function
   End If
   If DisplayRange.LastListItem = m_ListCount - 1 Then
      VerticalThumbY = ScaleHeight - m_BorderWidth - vThumbHeight - ScrollBarButtonHeight
      Exit Function
   End If

'  determine how by many items the list can be scrolled.
   NumClicks = m_ListCount - MaxDisplayItems

'  how many pixels of scroll thumb motion per list item scroll?
   PixelsPerScroll = (vThumbRange.Bottom - vThumbRange.Top) / NumClicks
'  how many pixels into the thumb's middle pixel motion range are we?
   PixelsPerScroll = PixelsPerScroll * DisplayRange.FirstListItem

   VerticalThumbY = (vThumbRange.Top + PixelsPerScroll) - vThumbHeight / 2

End Function

Private Sub DisplayVerticalTrackBar()

'*************************************************************************
'* displays the vertical scrollbar trackbar between up and down buttons. *
'*************************************************************************

'  determine the y position of the top of the vertical scrollbar thumb.
'  This is also used in the DisplayVerticalThumb routine. It is calculated
'  here so that if the scroll track under the thumb is being clicked down,
'  the correct height of trackbar under the thumb is highlighted.
   ThumbYPos = VerticalThumbY

'  display the trackbar.
   Call StretchDIBits(hdc, _
                      ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, _
                      m_BorderWidth + ScrollBarButtonHeight, _
                      ScrollBarButtonWidth, _
                      vScrollTrackHeight, _
                      0, 0, _
                      ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, _
                      vScrollTrackHeight, _
                      VTracklBits(0), _
                      VTrackuBIH, _
                      DIB_RGB_COLORS, _
                      vbSrcCopy)

'  if the mouse is currently clicked down in the track, repaint that portion accordingly.
   If MouseAction = MOUSE_DOWNED_IN_UPPERTRACK Then

      Call StretchDIBits(hdc, _
                         ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, _
                         m_BorderWidth + ScrollBarButtonHeight, _
                         ScrollBarButtonWidth, _
                         vScrollBarLocation.ScrollThumbLocation.Top - vThumbHeight - 7, _
                         0, 0, _
                         ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, _
                         vScrollTrackHeight, _
                         vClickTracklBits(0), _
                         vClickTrackuBIH, _
                         DIB_RGB_COLORS, _
                         vbSrcCopy)

   ElseIf MouseAction = MOUSE_DOWNED_IN_LOWERTRACK Then

      Call StretchDIBits(hdc, _
                         ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, _
                         ThumbYPos + vThumbHeight, _
                         ScrollBarButtonWidth, _
                         ScaleHeight - ThumbYPos - vThumbHeight - ScrollBarButtonHeight, _
                         0, 0, _
                         ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, _
                         vScrollTrackHeight, _
                         vClickTracklBits(0), _
                         vClickTrackuBIH, _
                         DIB_RGB_COLORS, _
                         vbSrcCopy)

   End If

End Sub

Private Sub DisplayTrackBarButton(ByVal WhichButton As Long)

'*************************************************************************
'* displays the appropriate scrollbar button.                            *
'*************************************************************************

   Select Case WhichButton

      Case UPBUTTON
         Call StretchDIBits(hdc, _
                            ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, _
                            m_BorderWidth, _
                            ScrollBarButtonWidth, _
                            ScrollBarButtonHeight, _
                            0, 0, _
                            ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, _
                            ScrollBarButtonHeight, _
                            TrackButtonlBits(0), _
                            TrackButtonuBIH, _
                            DIB_RGB_COLORS, _
                            vbSrcCopy)
         DrawScrollButtonArrow UPBUTTON
   
      Case DOWNBUTTON
         Call StretchDIBits(hdc, _
                            ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, _
                            ScaleHeight - m_BorderWidth - ScrollBarButtonHeight, _
                            ScrollBarButtonWidth, _
                            ScrollBarButtonHeight, _
                            0, 0, _
                            ScaleWidth - ScrollBarButtonWidth - m_BorderWidth, _
                            ScrollBarButtonHeight, _
                            TrackButtonlBits(0), _
                            TrackButtonuBIH, _
                            DIB_RGB_COLORS, _
                            vbSrcCopy)
         DrawScrollButtonArrow DOWNBUTTON

   End Select

End Sub

Private Sub DrawScrollButtonArrow(ByVal WhichButton As Long)

'*************************************************************************
'* draws both the up and down arrows on vertical scrollbar buttons.      *
'*************************************************************************

   Dim hPO           As Long       ' selected pen object.
   Dim hPN           As Long       ' pen object for drawing checkmark.
   Dim r             As Long       ' loop and result variable for api calls.
   Dim x1            As Long       ' the x coordinate of the start of the checkmark.
   Dim y1            As Long       ' the y coordinate of the start of the checkmark vertical line.
   Dim x2            As Long       ' the y coordinate of the end of the checkmark vertical line.
   Dim DrawDirection As Long       ' draw from left to right or right to left?
   Dim ArrowColor    As Long       ' up color or down color?

   x1 = ScaleWidth - ScrollBarButtonWidth - m_BorderWidth + 3
'  determine x coordinate of first part of check arrow to draw and the direction to draw.
   If WhichButton = UPBUTTON Then
      y1 = m_BorderWidth + ScrollBarButtonHeight - 6
      DrawDirection = -1
   Else
      y1 = ScaleHeight - m_BorderWidth - ScrollBarButtonHeight + 6
      DrawDirection = 1
   End If

'  select the correct button color.
   If WhichButton = UPBUTTON Then
      If MouseAction = MOUSE_DOWNED_IN_UPBUTTON Then
         ArrowColor = m_ActiveArrowDownColor
      Else
         ArrowColor = m_ActiveArrowUpColor
      End If
   Else
      If MouseAction = MOUSE_DOWNED_IN_DOWNBUTTON Then
         ArrowColor = m_ActiveArrowDownColor
      Else
         ArrowColor = m_ActiveArrowUpColor
      End If
   End If

'  draw the arrow.
   x2 = 9
   hPN = CreatePen(0, 1, ArrowColor)
   hPO = SelectObject(hdc, hPN)
   MoveTo hdc, x1, y1, ByVal 0&
   For r = 1 To 5
      LineTo hdc, x1 + x2, y1
      x1 = x1 + 1
      y1 = y1 + DrawDirection
      x2 = x2 - 2
      MoveTo hdc, x1, y1, ByVal 0&
   Next r

'  delete the pen object.
   r = SelectObject(hdc, hPO)
   r = DeleteObject(hPN)

End Sub

Private Function CalculateScrollTrackHeight() As Long

'**************************************************************************
'* calculates the vertical scrollbar's track height (the distance in      *
'* pixels between the bottom of the top arrow button and the top of the   *
'* bottom arrow button).  Borders are accounted for.                      *
'**************************************************************************

  CalculateScrollTrackHeight = ScaleHeight - (2 * ScrollBarButtonHeight) - (2 * m_BorderWidth)

End Function

Private Function CalculateVScrollThumbHeight() As Long

'*************************************************************************
'* returns the proper height of the vertical scrollbar thumb (in pixels) *
'* based on the number of items in the listbox, the number displayable   *
'* at one time, and the minimum allowable height of the thumb.  Makes    *
'* sure thumb is an odd number of pixels in height, so that the middle   *
'* of the thumb doesn't fall between two pixels.                         *
'*************************************************************************

   Dim VisiblePercentage As Single    ' percentage of list that can fit in display area.
   Dim THeight           As Long      ' preliminary thumb height.

'  no vertical scrollbar needed if the entire list can fit in the display area.  Just a safety net.
   If m_ListCount <= MaxDisplayItems Then
      CalculateVScrollThumbHeight = -1
      Exit Function
   End If

'  calculate the percentage of the list that can be displayed at one time.
   VisiblePercentage = MaxDisplayItems / m_ListCount

'  calculate the corresponding VisiblePercentage of the vertical scrollbar track height.
   THeight = Int(vScrollTrackHeight * VisiblePercentage)

'  if the thumb height is under the defined minimum, change it to the minimum.
   If THeight < vScrollMinThumbHeight Then
      CalculateVScrollThumbHeight = vScrollMinThumbHeight
   Else
'     make sure the height is an odd number of pixels so middle of thumb isn't between two pixels.
      If THeight Mod 2 = 0 Then
         THeight = THeight - 1
      End If
      CalculateVScrollThumbHeight = THeight
   End If

'  reset the flag so scrollbar thumb hieght doesn't get calculated unnecessarily.
   RecalculateThumbHeight = False

End Function

Private Sub ProcessVThumbScroll()

'*************************************************************************
'* allows the user to scroll through the list by dragging the vertical   *
'* scrollbar thumb up or down.                                           *
'*************************************************************************

   Dim ThumbRange     As Long        ' number of pixels middle of thumb can move in scroll track.
   Dim YPosPct        As Single      ' the percentage of the total list to move.
   Dim MouseMovement  As Long        ' how many pixels thumb was dragged from original position.
   Dim NumItemsToMove As Long        ' number of items to move list display by.

   ThumbScrolling = True
   MousePosInVThumb = MouseY - vScrollBarLocation.ScrollThumbLocation.Top

   While MouseAction <> MOUSE_NOACTION And DraggingVThumb

'     detect MouseMove for thumb positioning, and MouseUp to change value of MouseAction when it happens.
      DoEvents
'     keeps cpu usage from maxing out.  Thanks to Mike Douglas for the tip.
      Sleep 25

'     if mouse has moved below permissible scrolling range, display bottom of list and exit.
      If MouseY > ScaleHeight - m_BorderWidth - ScrollBarButtonHeight Then
         DisplayRange.LastListItem = m_ListCount - 1
         DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems + 1
         DisplayList
         ThumbScrolling = False ' used in MouseMove to see if we can resume thumb scrolling.
         Exit Sub
      End If

'     if mouse has moved above permissible scrolling range, display top of list and exit.
      If MouseY <= m_BorderWidth + ScrollBarButtonHeight Then
         DisplayRange.FirstListItem = 0
         DisplayRange.LastListItem = MaxDisplayItems - 1
         DisplayList
         ThumbScrolling = False ' used in MouseMove to see if we can resume thumb scrolling.
         Exit Sub
      End If

'     calculate how far the mouse has moved from the original mousedown y position.
      MouseMovement = MouseY - MouseDownYPos ' could be pos. or neg., depends on direction mouse moves.

'     determine the thumb's middle-pixel movement range.
      ThumbRange = vThumbRange.Bottom - vThumbRange.Top + 1 '# of pixels in the range.
'     determine how far, percentagewise, the list should be moved.
      YPosPct = (Abs(MouseMovement) / ThumbRange)
'     how many items is that?
      NumItemsToMove = Int(m_ListCount * YPosPct)

'     change it to negative if necessary.
      If MouseMovement < 0 Then
         NumItemsToMove = 0 - NumItemsToMove
      End If

'     only do the scroll if the range to display has changed.
      If NumItemsToMove <> 0 Then
         DisplayRange.FirstListItem = DisplayRange.FirstListItem + NumItemsToMove
         If DisplayRange.FirstListItem < 0 Then
            DisplayRange.FirstListItem = 0
         End If
         If DisplayRange.FirstListItem + MaxDisplayItems - 1 > m_ListCount Then
            DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems + 1
         End If
         DisplayRange.LastListItem = DisplayRange.FirstListItem + MaxDisplayItems - 1
         DisplayList MousePosInVThumb
         MousePosInVThumb = MouseY - vScrollBarLocation.ScrollThumbLocation.Top
         MouseDownYPos = MouseY
      End If

   Wend

End Sub

Private Sub ProcessMouseDragThumbOutOfRange(DidIt As Boolean)

'*************************************************************************
'* this code accounts for when mouse drags thumb above or below maximum  *
'* vertical scroll range.  The top or bottom of the list is then         *
'* automatically displayed.                                              *
'*************************************************************************

   DidIt = False

   If DraggingVThumb Then
      If MouseY <= m_BorderWidth + ScrollBarButtonHeight Then
         DisplayRange.FirstListItem = 0
         DisplayRange.LastListItem = MaxDisplayItems - 1
         DisplayList
         DidIt = True
      ElseIf MouseY >= vScrollBarLocation.DownButtonLocation.Top Then
         DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems + 1
         DisplayRange.LastListItem = m_ListCount - 1
         DisplayList
         DidIt = True
      End If
   End If

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Properties >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_InitProperties()

'*************************************************************************
'* initialize properties to the default constants.                       *
'*************************************************************************

   Set m_Picture = LoadPicture("")
   Set m_DisPicture = LoadPicture("")
   Set m_ListFont = Ambient.Font
   m_BackAngle = m_def_BackAngle
   m_BackColor2 = m_def_BackColor2
   m_BackColor1 = m_def_BackColor1
   m_BorderColor = m_def_BorderColor
   m_Enabled = m_def_Enabled
   m_BorderWidth = m_def_BorderWidth
   m_BackMiddleOut = m_def_BackMiddleOut
   m_CurveTopLeft = m_def_CurveTopLeft
   m_CurveTopRight = m_def_CurveTopRight
   m_CurveBottomLeft = m_def_CurveBottomLeft
   m_CurveBottomRight = m_def_CurveBottomRight
   m_ListIndex = m_def_ListIndex
   m_Sorted = m_def_Sorted
   m_RedrawFlag = m_def_RedrawFlag
   m_TextColor = m_def_TextColor
   m_SelColor1 = m_def_SelColor1
   m_SelColor2 = m_def_SelColor2
   m_SelTextColor = m_def_SelTextColor
   m_MultiSelect = m_def_MultiSelect
   m_Style = m_def_Style
   m_SelCount = m_def_SelCount
   m_TrackBarColor1 = m_def_TrackBarColor1
   m_TrackBarColor2 = m_def_TrackBarColor2
   m_ButtonColor1 = m_def_ButtonColor1
   m_ButtonColor2 = m_def_ButtonColor2
   m_ThumbColor1 = m_def_ThumbColor1
   m_ThumbColor2 = m_def_ThumbColor2
   m_ThumbBorderColor = m_def_ThumbBorderColor
   m_ArrowUpColor = m_def_ArrowUpColor
   m_ArrowDownColor = m_def_ArrowDownColor
   m_Theme = m_def_Theme
   m_CheckboxArrowColor = m_def_CheckboxArrowColor
   m_CheckBoxColor = m_def_CheckBoxColor
   m_FocusRectColor = m_def_FocusRectColor
   m_DblClickBehavior = m_def_DblClickBehavior
   m_NewIndex = m_def_NewIndex
   m_PictureMode = m_def_PictureMode
   m_CheckStyle = m_def_CheckStyle
   m_DragEnabled = m_def_DragEnabled
   m_TrackClickColor1 = m_def_TrackClickColor1
   m_TrackClickColor2 = m_def_TrackClickColor2
   m_SortAsNumeric = m_def_SortAsNumeric
   m_DisArrowDownColor = m_def_DisArrowDownColor
   m_DisArrowUpColor = m_def_DisArrowUpColor
   m_DisBackColor1 = m_def_DisBackColor1
   m_DisBackColor2 = m_def_DisBackColor2
   m_DisBorderColor = m_def_DisBorderColor
   m_DisButtonColor1 = m_def_DisButtonColor1
   m_DisButtonColor2 = m_def_DisButtonColor2
   m_DisCheckboxArrowColor = m_def_DisCheckboxArrowColor
   m_DisCheckboxColor = m_def_DisCheckboxColor
   m_DisFocusRectColor = m_def_DisFocusRectColor
   m_DisPictureMode = m_def_DisPictureMode
   m_DisSelColor1 = m_def_DisSelColor1
   m_DisSelColor2 = m_def_DisSelColor2
   m_DisSelTextColor = m_def_DisSelTextColor
   m_DisTextColor = m_def_DisTextColor
   m_DisThumbBorderColor = m_def_DisThumbBorderColor
   m_DisThumbColor1 = m_def_DisThumbColor1
   m_DisThumbColor2 = m_def_DisThumbColor2
   m_DisTrackbarColor1 = m_def_DisTrackbarColor1
   m_DisTrackbarColor2 = m_def_DisTrackbarColor2
   m_TopIndex = m_def_TopIndex
   m_ShowItemImages = m_def_ShowItemImages
   m_ItemImageSize = m_def_ItemImageSize
   m_ScaleWidth = UserControl.ScaleWidth 'm_def_ScaleWidth
   m_ScaleMode = m_def_ScaleMode
'   m_ScaleHeight = m_def_ScaleHeight
'   m_AutoRedraw = m_def_AutoRedraw

'  initialize appropriate display colors.
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*************************************************************************
'* read properties in the property bag.                                  *
'*************************************************************************

   With PropBag
      Set m_ListFont = .ReadProperty("ListFont", Ambient.Font)
      Set UserControl.Font = m_ListFont
      Set m_Picture = .ReadProperty("Picture", Nothing)
      Set m_DisPicture = PropBag.ReadProperty("DisPicture", Nothing)
      m_BackAngle = .ReadProperty("BackAngle", m_def_BackAngle)
      m_BackColor2 = .ReadProperty("BackColor2", m_def_BackColor2)
      m_BackColor1 = .ReadProperty("BackColor1", m_def_BackColor1)
      m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
      m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
      m_BorderWidth = .ReadProperty("BorderWidth", m_def_BorderWidth)
      m_BackMiddleOut = .ReadProperty("BackMiddleOut", m_def_BackMiddleOut)
      m_CurveTopLeft = .ReadProperty("CurveTopLeft", m_def_CurveTopLeft)
      m_CurveTopRight = .ReadProperty("CurveTopRight", m_def_CurveTopRight)
      m_CurveBottomLeft = .ReadProperty("CurveBottomLeft", m_def_CurveBottomLeft)
      m_CurveBottomRight = .ReadProperty("CurveBottomRight", m_def_CurveBottomRight)
      m_Sorted = .ReadProperty("Sorted", m_def_Sorted)
      m_RedrawFlag = .ReadProperty("RedrawFlag", m_def_RedrawFlag)
      m_TextColor = .ReadProperty("TextColor", m_def_TextColor)
      m_SelColor1 = .ReadProperty("SelColor1", m_def_SelColor1)
      m_SelColor2 = .ReadProperty("SelColor2", m_def_SelColor2)
      m_SelTextColor = .ReadProperty("SelTextColor", m_def_SelTextColor)
      m_MultiSelect = .ReadProperty("MultiSelect", m_def_MultiSelect)
      m_Style = .ReadProperty("Style", m_def_Style)
      m_SelCount = .ReadProperty("SelCount", m_def_SelCount)
      m_TrackBarColor1 = .ReadProperty("TrackBarColor1", m_def_TrackBarColor1)
      m_TrackBarColor2 = .ReadProperty("TrackBarColor2", m_def_TrackBarColor2)
      m_ButtonColor1 = .ReadProperty("ButtonColor1", m_def_ButtonColor1)
      m_ButtonColor2 = .ReadProperty("ButtonColor2", m_def_ButtonColor2)
      m_ThumbColor1 = .ReadProperty("ThumbColor1", m_def_ThumbColor1)
      m_ThumbColor2 = .ReadProperty("ThumbColor2", m_def_ThumbColor2)
      m_ThumbBorderColor = .ReadProperty("ThumbBorderColor", m_def_ThumbBorderColor)
      m_ArrowUpColor = .ReadProperty("ArrowUpColor", m_def_ArrowUpColor)
      m_ArrowDownColor = .ReadProperty("ArrowDownColor", m_def_ArrowDownColor)
      m_Theme = .ReadProperty("Theme", m_def_Theme)
      m_CheckboxArrowColor = .ReadProperty("CheckboxArrowColor", m_def_CheckboxArrowColor)
      m_CheckBoxColor = .ReadProperty("CheckBoxColor", m_def_CheckBoxColor)
      m_FocusRectColor = .ReadProperty("FocusRectColor", m_def_FocusRectColor)
      m_DblClickBehavior = .ReadProperty("DblClickBehavior", m_def_DblClickBehavior)
      m_NewIndex = .ReadProperty("NewIndex", m_def_NewIndex)
      m_PictureMode = .ReadProperty("PictureMode", m_def_PictureMode)
      m_CheckStyle = .ReadProperty("CheckStyle", m_def_CheckStyle)
      m_DragEnabled = .ReadProperty("DragEnabled", m_def_DragEnabled)
      m_TrackClickColor1 = .ReadProperty("TrackClickColor1", m_def_TrackClickColor1)
      m_TrackClickColor2 = .ReadProperty("TrackClickColor2", m_def_TrackClickColor2)
      m_SortAsNumeric = .ReadProperty("SortAsNumeric", m_def_SortAsNumeric)
      m_DisArrowDownColor = .ReadProperty("DisArrowDownColor", m_def_DisArrowDownColor)
      m_DisArrowUpColor = .ReadProperty("DisArrowUpColor", m_def_DisArrowUpColor)
      m_DisBackColor1 = .ReadProperty("DisBackColor1", m_def_DisBackColor1)
      m_DisBackColor2 = .ReadProperty("DisBackColor2", m_def_DisBackColor2)
      m_DisBorderColor = .ReadProperty("DisBorderColor", m_def_DisBorderColor)
      m_DisButtonColor1 = .ReadProperty("DisButtonColor1", m_def_DisButtonColor1)
      m_DisButtonColor2 = .ReadProperty("DisButtonColor2", m_def_DisButtonColor2)
      m_DisCheckboxArrowColor = .ReadProperty("DisCheckboxArrowColor", m_def_DisCheckboxArrowColor)
      m_DisCheckboxColor = .ReadProperty("DisCheckboxColor", m_def_DisCheckboxColor)
      m_DisFocusRectColor = .ReadProperty("DisFocusRectColor", m_def_DisFocusRectColor)
      m_DisPictureMode = .ReadProperty("DisPictureMode", m_def_DisPictureMode)
      m_DisSelColor1 = .ReadProperty("DisSelColor1", m_def_DisSelColor1)
      m_DisSelColor2 = .ReadProperty("DisSelColor2", m_def_DisSelColor2)
      m_DisSelTextColor = .ReadProperty("DisSelTextColor", m_def_DisSelTextColor)
      m_DisTextColor = .ReadProperty("DisTextColor", m_def_DisTextColor)
      m_DisThumbBorderColor = .ReadProperty("DisThumbBorderColor", m_def_DisThumbBorderColor)
      m_DisThumbColor1 = .ReadProperty("DisThumbColor1", m_def_DisThumbColor1)
      m_DisThumbColor2 = .ReadProperty("DisThumbColor2", m_def_DisThumbColor2)
      m_DisTrackbarColor1 = .ReadProperty("DisTrackbarColor1", m_def_DisTrackbarColor1)
      m_DisTrackbarColor2 = .ReadProperty("DisTrackbarColor2", m_def_DisTrackbarColor2)
      m_TopIndex = .ReadProperty("TopIndex", m_def_TopIndex)
      m_ShowItemImages = .ReadProperty("ShowItemImages", m_def_ShowItemImages)
      m_ItemImageSize = .ReadProperty("ItemImageSize", m_def_ItemImageSize)
   m_ScaleWidth = UserControl.ScaleWidth ' PropBag.ReadProperty("ScaleWidth", m_def_ScaleWidth)
   m_ScaleMode = .ReadProperty("ScaleMode", m_def_ScaleMode)
   'm_ScaleHeight = PropBag.ReadProperty("ScaleHeight", m_def_ScaleHeight)
   'm_AutoRedraw = PropBag.ReadProperty("AutoRedraw", m_def_AutoRedraw)
   End With

   DragFlag = m_DragEnabled

'  initially, the .ListIndex property is determined by the value of the .MultiSelect property.
   m_ListIndex = IIf(m_MultiSelect = vbMultiSelectNone, -1, 0)

'  initialize appropriate display colors.
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If

'  initialize the list item display range structure.
   DisplayRange.FirstListItem = -1
   DisplayRange.LastListItem = -1

'  initialize gradients, list item height, display coordinates.
   InitListBoxDisplayCharacteristics

'  start up the subclassing.
   StartSubclassing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'*************************************************************************
'* write the properties in the property bag.                             *
'*************************************************************************

   With PropBag
      .WriteProperty "Picture", m_Picture, Nothing
      .WriteProperty "DisPicture", m_DisPicture, Nothing
      .WriteProperty "BackAngle", m_BackAngle, m_def_BackAngle
      .WriteProperty "BackColor2", m_BackColor2, m_def_BackColor2
      .WriteProperty "BackColor1", m_BackColor1, m_def_BackColor1
      .WriteProperty "BorderColor", m_BorderColor, m_def_BorderColor
      .WriteProperty "Enabled", m_Enabled, m_def_Enabled
      .WriteProperty "BorderWidth", m_BorderWidth, m_def_BorderWidth
      .WriteProperty "BackMiddleOut", m_BackMiddleOut, m_def_BackMiddleOut
      .WriteProperty "CurveTopLeft", m_CurveTopLeft, m_def_CurveTopLeft
      .WriteProperty "CurveTopRight", m_CurveTopRight, m_def_CurveTopRight
      .WriteProperty "CurveBottomLeft", m_CurveBottomLeft, m_def_CurveBottomLeft
      .WriteProperty "CurveBottomRight", m_CurveBottomRight, m_def_CurveBottomRight
      .WriteProperty "ListIndex", m_ListIndex, m_def_ListIndex
      .WriteProperty "ListFont", m_ListFont, Ambient.Font
      .WriteProperty "Sorted", m_Sorted, m_def_Sorted
      .WriteProperty "RedrawFlag", m_RedrawFlag, m_def_RedrawFlag
      .WriteProperty "TextColor", m_TextColor, m_def_TextColor
      .WriteProperty "SelColor1", m_SelColor1, m_def_SelColor1
      .WriteProperty "SelColor2", m_SelColor2, m_def_SelColor2
      .WriteProperty "SelTextColor", m_SelTextColor, m_def_SelTextColor
      .WriteProperty "MultiSelect", m_MultiSelect, m_def_MultiSelect
      .WriteProperty "Style", m_Style, m_def_Style
      .WriteProperty "SelCount", m_SelCount, m_def_SelCount
      .WriteProperty "TrackBarColor1", m_TrackBarColor1, m_def_TrackBarColor1
      .WriteProperty "TrackBarColor2", m_TrackBarColor2, m_def_TrackBarColor2
      .WriteProperty "ButtonColor1", m_ButtonColor1, m_def_ButtonColor1
      .WriteProperty "ButtonColor2", m_ButtonColor2, m_def_ButtonColor2
      .WriteProperty "ThumbColor1", m_ThumbColor1, m_def_ThumbColor1
      .WriteProperty "ThumbColor2", m_ThumbColor2, m_def_ThumbColor2
      .WriteProperty "ThumbBorderColor", m_ThumbBorderColor, m_def_ThumbBorderColor
      .WriteProperty "ArrowUpColor", m_ArrowUpColor, m_def_ArrowUpColor
      .WriteProperty "ArrowDownColor", m_ArrowDownColor, m_def_ArrowDownColor
      .WriteProperty "Theme", m_Theme, m_def_Theme
      .WriteProperty "CheckboxArrowColor", m_CheckboxArrowColor, m_def_CheckboxArrowColor
      .WriteProperty "CheckBoxColor", m_CheckBoxColor, m_def_CheckBoxColor
      .WriteProperty "FocusRectColor", m_FocusRectColor, m_def_FocusRectColor
      .WriteProperty "DblClickBehavior", m_DblClickBehavior, m_def_DblClickBehavior
      .WriteProperty "NewIndex", m_NewIndex, m_def_NewIndex
      .WriteProperty "PictureMode", m_PictureMode, m_def_PictureMode
      .WriteProperty "CheckStyle", m_CheckStyle, m_def_CheckStyle
      .WriteProperty "DragEnabled", m_DragEnabled, m_def_DragEnabled
      .WriteProperty "TrackClickColor1", m_TrackClickColor1, m_def_TrackClickColor1
      .WriteProperty "TrackClickColor2", m_TrackClickColor2, m_def_TrackClickColor2
      .WriteProperty "SortAsNumeric", m_SortAsNumeric, m_def_SortAsNumeric
      .WriteProperty "DisArrowDownColor", m_DisArrowDownColor, m_def_DisArrowDownColor
      .WriteProperty "DisArrowUpColor", m_DisArrowUpColor, m_def_DisArrowUpColor
      .WriteProperty "DisBackColor1", m_DisBackColor1, m_def_DisBackColor1
      .WriteProperty "DisBackColor2", m_DisBackColor2, m_def_DisBackColor2
      .WriteProperty "DisBorderColor", m_DisBorderColor, m_def_DisBorderColor
      .WriteProperty "DisButtonColor1", m_DisButtonColor1, m_def_DisButtonColor1
      .WriteProperty "DisButtonColor2", m_DisButtonColor2, m_def_DisButtonColor2
      .WriteProperty "DisCheckboxArrowColor", m_DisCheckboxArrowColor, m_def_DisCheckboxArrowColor
      .WriteProperty "DisCheckboxColor", m_DisCheckboxColor, m_def_DisCheckboxColor
      .WriteProperty "DisFocusRectColor", m_DisFocusRectColor, m_def_DisFocusRectColor
      .WriteProperty "DisPictureMode", m_DisPictureMode, m_def_DisPictureMode
      .WriteProperty "DisSelColor1", m_DisSelColor1, m_def_DisSelColor1
      .WriteProperty "DisSelColor2", m_DisSelColor2, m_def_DisSelColor2
      .WriteProperty "DisSelTextColor", m_DisSelTextColor, m_def_DisSelTextColor
      .WriteProperty "DisTextColor", m_DisTextColor, m_def_DisTextColor
      .WriteProperty "DisThumbBorderColor", m_DisThumbBorderColor, m_def_DisThumbBorderColor
      .WriteProperty "DisThumbColor1", m_DisThumbColor1, m_def_DisThumbColor1
      .WriteProperty "DisThumbColor2", m_DisThumbColor2, m_def_DisThumbColor2
      .WriteProperty "DisTrackbarColor1", m_DisTrackbarColor1, m_def_DisTrackbarColor1
      .WriteProperty "DisTrackbarColor2", m_DisTrackbarColor2, m_def_DisTrackbarColor2
      .WriteProperty "TopIndex", m_TopIndex, m_def_TopIndex
      .WriteProperty "ShowItemImages", m_ShowItemImages, m_def_ShowItemImages
      .WriteProperty "ItemImageSize", m_ItemImageSize, m_def_ItemImageSize
   End With

End Sub

Public Property Get BackAngle() As Single
Attribute BackAngle.VB_Description = "The angle of the listbox background gradient."
Attribute BackAngle.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BackAngle = m_BackAngle
End Property

Public Property Let BackAngle(ByVal New_BackAngle As Single)
'  do some bounds checking.
   If New_BackAngle > 360 Then
      New_BackAngle = 360
   ElseIf New_BackAngle < 0 Then
      New_BackAngle = 0
   End If
   m_BackAngle = New_BackAngle
   PropertyChanged "BackAngle"
   CalculateBackGroundGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BackColor1() As OLE_COLOR
Attribute BackColor1.VB_Description = "The first color of the listbox background gradient."
Attribute BackColor1.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
   m_BackColor1 = New_BackColor1
   PropertyChanged "BackColor1"
   CalculateBackGroundGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "The second color of the listbox background gradient."
Attribute BackColor2.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
   m_BackColor2 = New_BackColor2
   PropertyChanged "BackColor2"
   CalculateBackGroundGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BackMiddleOut() As Boolean
Attribute BackMiddleOut.VB_Description = "Allows the background gradient to be middle-out (Color 1> Color 2 > Color 1)."
Attribute BackMiddleOut.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BackMiddleOut = m_BackMiddleOut
End Property

Public Property Let BackMiddleOut(ByVal New_BackMiddleOut As Boolean)
   m_BackMiddleOut = New_BackMiddleOut
   PropertyChanged "BackMiddleOut"
   CalculateBackGroundGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "The color of the ListBox border."
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   m_BorderColor = New_BorderColor
   PropertyChanged "BorderColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "The width, in pixels, of the ListBox border."
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
   m_BorderWidth = New_BorderWidth
   PropertyChanged "BorderWidth"
'  since the border width has been changed, we have to recalculate text display boundaries.
   InitTextDisplayCharacteristics
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CurveBottomLeft() As Long
Attribute CurveBottomLeft.VB_Description = "The amount of curve of the bottom left corner of the ListBox."
Attribute CurveBottomLeft.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveBottomLeft = m_CurveBottomLeft
End Property

Public Property Let CurveBottomLeft(ByVal New_CurveBottomLeft As Long)
   m_CurveBottomLeft = New_CurveBottomLeft
   PropertyChanged "CurveBottomLeft"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CurveBottomRight() As Long
Attribute CurveBottomRight.VB_Description = "The amount of curve of the bottom right corner of the ListBox."
Attribute CurveBottomRight.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveBottomRight = m_CurveBottomRight
End Property

Public Property Let CurveBottomRight(ByVal New_CurveBottomRight As Long)
   m_CurveBottomRight = New_CurveBottomRight
   PropertyChanged "CurveBottomRight"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CurveTopLeft() As Long
Attribute CurveTopLeft.VB_Description = "The amount of curve of the top left corner of the ListBox."
Attribute CurveTopLeft.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveTopLeft = m_CurveTopLeft
End Property

Public Property Let CurveTopLeft(ByVal New_CurveTopLeft As Long)
   m_CurveTopLeft = New_CurveTopLeft
   PropertyChanged "CurveTopLeft"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CurveTopRight() As Long
Attribute CurveTopRight.VB_Description = "The amount of curve of the top right corner of the ListBox."
Attribute CurveTopRight.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveTopRight = m_CurveTopRight
End Property

Public Property Let CurveTopRight(ByVal New_CurveTopRight As Long)
   m_CurveTopRight = New_CurveTopRight
   PropertyChanged "CurveTopRight"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get DisArrowDownColor() As OLE_COLOR
Attribute DisArrowDownColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisArrowDownColor = m_DisArrowDownColor
End Property

Public Property Let DisArrowDownColor(ByVal New_DisArrowDownColor As OLE_COLOR)
   m_DisArrowDownColor = New_DisArrowDownColor
   PropertyChanged "DisArrowDownColor"
End Property

Public Property Get DisArrowUpColor() As OLE_COLOR
Attribute DisArrowUpColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisArrowUpColor = m_DisArrowUpColor
End Property

Public Property Let DisArrowUpColor(ByVal New_DisArrowUpColor As OLE_COLOR)
   m_DisArrowUpColor = New_DisArrowUpColor
   PropertyChanged "DisArrowUpColor"
End Property

Public Property Get DisBackColor1() As OLE_COLOR
Attribute DisBackColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisBackColor1 = m_DisBackColor1
End Property

Public Property Let DisBackColor1(ByVal New_DisBackColor1 As OLE_COLOR)
   m_DisBackColor1 = New_DisBackColor1
   PropertyChanged "DisBackColor1"
End Property

Public Property Get DisBackColor2() As OLE_COLOR
Attribute DisBackColor2.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisBackColor2 = m_DisBackColor2
End Property

Public Property Let DisBackColor2(ByVal New_DisBackColor2 As OLE_COLOR)
   m_DisBackColor2 = New_DisBackColor2
   PropertyChanged "DisBackColor2"
End Property

Public Property Get DisBorderColor() As OLE_COLOR
Attribute DisBorderColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisBorderColor = m_DisBorderColor
End Property

Public Property Let DisBorderColor(ByVal New_DisBorderColor As OLE_COLOR)
   m_DisBorderColor = New_DisBorderColor
   PropertyChanged "DisBorderColor"
End Property

Public Property Get DisButtonColor1() As OLE_COLOR
Attribute DisButtonColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisButtonColor1 = m_DisButtonColor1
End Property

Public Property Let DisButtonColor1(ByVal New_DisButtonColor1 As OLE_COLOR)
   m_DisButtonColor1 = New_DisButtonColor1
   PropertyChanged "DisButtonColor1"
End Property

Public Property Get DisButtonColor2() As OLE_COLOR
Attribute DisButtonColor2.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisButtonColor2 = m_DisButtonColor2
End Property

Public Property Let DisButtonColor2(ByVal New_DisButtonColor2 As OLE_COLOR)
   m_DisButtonColor2 = New_DisButtonColor2
   PropertyChanged "DisButtonColor2"
End Property

Public Property Get DisCheckboxArrowColor() As OLE_COLOR
Attribute DisCheckboxArrowColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisCheckboxArrowColor = m_DisCheckboxArrowColor
End Property

Public Property Let DisCheckboxArrowColor(ByVal New_DisCheckboxArrowColor As OLE_COLOR)
   m_DisCheckboxArrowColor = New_DisCheckboxArrowColor
   PropertyChanged "DisCheckboxArrowColor"
End Property

Public Property Get DisCheckboxColor() As OLE_COLOR
Attribute DisCheckboxColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisCheckboxColor = m_DisCheckboxColor
End Property

Public Property Let DisCheckboxColor(ByVal New_DisCheckboxColor As OLE_COLOR)
   m_DisCheckboxColor = New_DisCheckboxColor
   PropertyChanged "DisCheckboxColor"
End Property

Public Property Get DisFocusRectColor() As OLE_COLOR
Attribute DisFocusRectColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisFocusRectColor = m_DisFocusRectColor
End Property

Public Property Let DisFocusRectColor(ByVal New_DisFocusRectColor As OLE_COLOR)
   m_DisFocusRectColor = New_DisFocusRectColor
   PropertyChanged "DisFocusRectColor"
End Property

Public Property Get DisPicture() As Picture
Attribute DisPicture.VB_ProcData.VB_Invoke_Property = ";Disabled"
   Set DisPicture = m_DisPicture
End Property

Public Property Set DisPicture(ByVal New_DisPicture As Picture)
   Set m_DisPicture = New_DisPicture
   PropertyChanged "DisPicture"
'  this flag tells Redraw to re-blit the new background to the virtual DC.
   ChangingPicture = True
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get DisPictureMode() As PictureModeOptions
Attribute DisPictureMode.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisPictureMode = m_DisPictureMode
End Property

Public Property Let DisPictureMode(ByVal New_DisPictureMode As PictureModeOptions)
   m_DisPictureMode = New_DisPictureMode
   PropertyChanged "DisPictureMode"
End Property

Public Property Get DisSelColor1() As OLE_COLOR
Attribute DisSelColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisSelColor1 = m_DisSelColor1
End Property

Public Property Let DisSelColor1(ByVal New_DisSelColor1 As OLE_COLOR)
   m_DisSelColor1 = New_DisSelColor1
   PropertyChanged "DisSelColor1"
End Property

Public Property Get DisSelColor2() As OLE_COLOR
Attribute DisSelColor2.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisSelColor2 = m_DisSelColor2
End Property

Public Property Let DisSelColor2(ByVal New_DisSelColor2 As OLE_COLOR)
   m_DisSelColor2 = New_DisSelColor2
   PropertyChanged "DisSelColor2"
End Property

Public Property Get DisSelTextColor() As OLE_COLOR
Attribute DisSelTextColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisSelTextColor = m_DisSelTextColor
End Property

Public Property Let DisSelTextColor(ByVal New_DisSelTextColor As OLE_COLOR)
   m_DisSelTextColor = New_DisSelTextColor
   PropertyChanged "DisSelTextColor"
End Property

Public Property Get DisTextColor() As OLE_COLOR
Attribute DisTextColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisTextColor = m_DisTextColor
End Property

Public Property Let DisTextColor(ByVal New_DisTextColor As OLE_COLOR)
   m_DisTextColor = New_DisTextColor
   PropertyChanged "DisTextColor"
End Property

Public Property Get DisThumbBorderColor() As OLE_COLOR
Attribute DisThumbBorderColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisThumbBorderColor = m_DisThumbBorderColor
End Property

Public Property Let DisThumbBorderColor(ByVal New_DisThumbBorderColor As OLE_COLOR)
   m_DisThumbBorderColor = New_DisThumbBorderColor
   PropertyChanged "DisThumbBorderColor"
End Property

Public Property Get DisThumbColor1() As OLE_COLOR
Attribute DisThumbColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisThumbColor1 = m_DisThumbColor1
End Property

Public Property Let DisThumbColor1(ByVal New_DisThumbColor1 As OLE_COLOR)
   m_DisThumbColor1 = New_DisThumbColor1
   PropertyChanged "DisThumbColor1"
End Property

Public Property Get DisThumbColor2() As OLE_COLOR
Attribute DisThumbColor2.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisThumbColor2 = m_DisThumbColor2
End Property

Public Property Let DisThumbColor2(ByVal New_DisThumbColor2 As OLE_COLOR)
   m_DisThumbColor2 = New_DisThumbColor2
   PropertyChanged "DisThumbColor2"
End Property

Public Property Get DisTrackbarColor1() As OLE_COLOR
Attribute DisTrackbarColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisTrackbarColor1 = m_DisTrackbarColor1
End Property

Public Property Let DisTrackbarColor1(ByVal New_DisTrackbarColor1 As OLE_COLOR)
   m_DisTrackbarColor1 = New_DisTrackbarColor1
   PropertyChanged "DisTrackbarColor1"
End Property

Public Property Get DisTrackbarColor2() As OLE_COLOR
Attribute DisTrackbarColor2.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisTrackbarColor2 = m_DisTrackbarColor2
End Property

Public Property Let DisTrackbarColor2(ByVal New_DisTrackbarColor2 As OLE_COLOR)
   m_DisTrackbarColor2 = New_DisTrackbarColor2
   PropertyChanged "DisTrackbarColor2"
End Property

Public Property Get Enabled() As Boolean
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If
   InitListBoxDisplayCharacteristics
   PropertyChanged "Enabled"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
   hwnd = UserControl.hwnd
End Property

Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
   hdc = UserControl.hdc
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "The bitmap to display in lieu of a gradient in the ListBox background."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
   Set m_Picture = New_Picture
   PropertyChanged "Picture"
'  this flag tells Redraw to re-blit the new background to the virtual DC.
   ChangingPicture = True
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get List(ByVal Index As Long) As String
   If Index >= 0 And Index <= UBound(ListArray) Then
      List = ListArray(Index)
   End If
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
   ItemData = ItemDataArray(Index)
End Property

Public Property Let ItemData(ByVal Index As Long, NewValue As Long)
   ItemDataArray(Index) = NewValue
End Property

Public Property Get Selected(ByVal Index As Long) As Boolean
   Selected = SelectedArray(Index)
End Property

Public Property Let Selected(ByVal Index As Long, NewValue As Boolean)

'*************************************************************************
'* processes programmatic selection/deselection of specified list item.  *
'* all 3 multiple selection modes (MultiSelect Simple, MultiSelect Exten-*
'* ded and CheckBox) are processed the same way here; only MultiSelect   *
'* None mode is treated differently.                                     *
'*************************************************************************

   Dim PreviouslySelected As Boolean

   PreviouslySelected = SelectedArray(Index)

   If m_Style = [Standard] And m_MultiSelect = vbMultiSelectNone Then
'     in MultiSelect None mode, if Selected(Index) is set to True, all other list items
'     are deselected and variables are set to the index of the newly selected item.
      SetSelectedArrayRange 0, m_ListCount - 1, False
      SelectedArray(Index) = NewValue
      If NewValue Then
         m_ListIndex = Index
         m_SelCount = 1
         LastSelectedItem = Index
         ItemWithFocus = Index
      Else
'        if Selected(Index) is set to False, all items are deselected.  .ListIndex
'        property is set to -1, which the default in this mode for no selected items.
         If PreviouslySelected Then
            m_ListIndex = -1
            m_SelCount = 0
            LastSelectedItem = -1
            ItemWithFocus = 0
         End If
      End If
   Else
'     in modes that allow multiple selections, the item is selected or deselected, and
'     the .SelCount property is adjusted accordingly.  .ListIndex points to Index item.
      SelectedArray(Index) = NewValue
      m_ListIndex = Index
      ItemWithFocus = Index
      LastSelectedItem = Index
      If PreviouslySelected And Not NewValue Then
         m_SelCount = m_SelCount - 1
      ElseIf Not PreviouslySelected And NewValue Then
         m_SelCount = m_SelCount + 1
      End If
   End If

   If m_RedrawFlag Then ' added if 01/25/06 to account for many selections via code.
      DisplayList
   End If

End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_Description = "The index of the most recently selected list item."
Attribute ListIndex.VB_ProcData.VB_Invoke_Property = ";Text"
   ListIndex = m_ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Long)

'*************************************************************************
'* processes programmatic alteration of the .ListIndex property.         *
'*************************************************************************

'  can't modify this property in design mode.
   If Ambient.UserMode = False Then Err.Raise 387

   m_ListIndex = New_ListIndex

   If m_Style = [Standard] And m_MultiSelect = vbMultiSelectNone Then
'     in MultiSelect None mode, if the new .ListIndex is -1,  clear all selections.  If
'     it's > -1, clear any existing selection and select the item pointed to by .ListIndex.
      If m_ListIndex = -1 Then
         SetSelectedArrayRange 0, m_ListCount - 1, False
         LastSelectedItem = m_ListIndex
         m_SelCount = 0
         ItemWithFocus = 0
      Else
         SetSelectedArrayRange 0, m_ListCount - 1, False
         LastSelectedItem = m_ListIndex
         m_SelCount = 1
         SelectedArray(m_ListIndex) = True
         ItemWithFocus = m_ListIndex
      End If
   Else
'     for MultiSelect Simple, MultiSelect Extended and CheckBox modes,
'     move the focus rectangle to the item pointed to by .ListIndex.
      ItemWithFocus = m_ListIndex
      If m_Style = [CheckBox] Then
'        in CheckBox mode, selection bar always moves with the focus rectangle.
         LastSelectedItem = m_ListIndex
      End If
   End If

   DisplayList

   PropertyChanged "ListIndex"

End Property

Public Property Get ListCount() As Long
   ListCount = m_ListCount
End Property

Public Property Get ListFont() As Font
Attribute ListFont.VB_Description = "The font used to display list items."
Attribute ListFont.VB_ProcData.VB_Invoke_Property = ";Text"
   Set ListFont = m_ListFont
End Property

Public Property Set ListFont(ByVal New_ListFont As Font)
   Set m_ListFont = New_ListFont
   Set UserControl.Font = m_ListFont
'  get the height range of characters in the current font.
   ListItemHeight = TextHeight("^j")
   InitTextDisplayCharacteristics
   PropertyChanged "ListFont"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "When True, ListBox items are automatically maintained in ascending order."
Attribute Sorted.VB_ProcData.VB_Invoke_Property = ";Behavior"
   Sorted = m_Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
   m_Sorted = New_Sorted
   PropertyChanged "Sorted"
End Property

Public Property Get RedrawFlag() As Boolean
Attribute RedrawFlag.VB_MemberFlags = "400"
   If Ambient.UserMode Then Err.Raise 393
   RedrawFlag = m_RedrawFlag
End Property

Public Property Let RedrawFlag(ByVal New_RedrawFlag As Boolean)
'  this makes the property unavailable at design time.
   If Ambient.UserMode = False Then Err.Raise 387
   m_RedrawFlag = New_RedrawFlag
   PropertyChanged "RedrawFlag"
   RedrawControl ' if RedrawFlag is now True, the control will redraw.
End Property

Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "The color of list item text when it is not highlighted by the selection bar."
Attribute TextColor.VB_ProcData.VB_Invoke_Property = ";Text"
   TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
   m_TextColor = New_TextColor
   PropertyChanged "TextColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get SelColor1() As OLE_COLOR
Attribute SelColor1.VB_Description = "The first gradient color of the list item selection bar."
Attribute SelColor1.VB_ProcData.VB_Invoke_Property = ";Text"
   SelColor1 = m_SelColor1
End Property

Public Property Let SelColor1(ByVal New_SelColor1 As OLE_COLOR)
   m_SelColor1 = New_SelColor1
   PropertyChanged "SelColor1"
   CalculateHighlightBarGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get SelColor2() As OLE_COLOR
Attribute SelColor2.VB_Description = "The second gradient color of the list item selection bar."
Attribute SelColor2.VB_ProcData.VB_Invoke_Property = ";Text"
   SelColor2 = m_SelColor2
End Property

Public Property Let SelColor2(ByVal New_SelColor2 As OLE_COLOR)
   m_SelColor2 = New_SelColor2
   PropertyChanged "SelColor2"
   CalculateHighlightBarGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get SelTextColor() As OLE_COLOR
Attribute SelTextColor.VB_Description = "The color of list item text when it is highlighted by the selection bar."
Attribute SelTextColor.VB_ProcData.VB_Invoke_Property = ";Text"
   SelTextColor = m_SelTextColor
End Property

Public Property Let SelTextColor(ByVal New_SelTextColor As OLE_COLOR)
   m_SelTextColor = New_SelTextColor
   PropertyChanged "SelTextColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get MultiSelect() As SelectionOptions
Attribute MultiSelect.VB_Description = "Sets the main ListBox operation mode: MultiSelect None, Simple, or Extended."
Attribute MultiSelect.VB_ProcData.VB_Invoke_Property = ";Operation Modes"
   MultiSelect = m_MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As SelectionOptions)
'  can't change at runtime.
   If Ambient.UserMode Then Err.Raise 382
   m_MultiSelect = New_MultiSelect
   PropertyChanged "MultiSelect"
End Property

Public Property Get Style() As ListItemOptions
Attribute Style.VB_Description = "Sets the display and operation of the ListBox to either Standard or CheckBox styles.  CheckBox style supercedes all .MultiSelect operation modes."
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Operation Modes"
   Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As ListItemOptions)
'  can't change at runtime.
   'If Ambient.UserMode Then Err.Raise 382
   m_Style = New_Style
   InitListBoxDisplayCharacteristics
   RedrawControl
   PropertyChanged "Style"
   UserControl.Refresh
End Property

Public Property Get SelCount() As Long
Attribute SelCount.VB_MemberFlags = "400"
   SelCount = m_SelCount
End Property

Public Property Let SelCount(ByVal New_SelCount As Long)
'  can't change at design time or runtime.
   If Ambient.UserMode = False Then Err.Raise 387
   If Ambient.UserMode Then Err.Raise 382
   m_SelCount = New_SelCount
   PropertyChanged "SelCount"
End Property

Public Property Get TrackBarColor1() As OLE_COLOR
Attribute TrackBarColor1.VB_Description = "The first gradient color of the vertical scrollbar trackbar."
Attribute TrackBarColor1.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   TrackBarColor1 = m_TrackBarColor1
End Property

Public Property Let TrackBarColor1(ByVal New_TrackBarColor1 As OLE_COLOR)
   m_TrackBarColor1 = New_TrackBarColor1
   PropertyChanged "TrackBarColor1"
   CalculateVerticalTrackbarGradientUnclicked
   DisplayVerticalScrollBar
   UserControl.Refresh
End Property

Public Property Get TrackBarColor2() As OLE_COLOR
Attribute TrackBarColor2.VB_Description = "The second gradient color of the vertical scrollbar trackbar."
Attribute TrackBarColor2.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   TrackBarColor2 = m_TrackBarColor2
End Property

Public Property Let TrackBarColor2(ByVal New_TrackBarColor2 As OLE_COLOR)
   m_TrackBarColor2 = New_TrackBarColor2
   PropertyChanged "TrackBarColor2"
   CalculateVerticalTrackbarGradientUnclicked
   DisplayVerticalScrollBar
   UserControl.Refresh
End Property

Public Property Get ButtonColor1() As OLE_COLOR
Attribute ButtonColor1.VB_Description = "The first gradient color of the vertical scrollbar up/down buttons."
Attribute ButtonColor1.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   ButtonColor1 = m_ButtonColor1
End Property

Public Property Let ButtonColor1(ByVal New_ButtonColor1 As OLE_COLOR)
   m_ButtonColor1 = New_ButtonColor1
   PropertyChanged "ButtonColor1"
   CalculateScrollbarButtonGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get ButtonColor2() As OLE_COLOR
Attribute ButtonColor2.VB_Description = "The second gradient color of the vertical scrollbar up/down buttons."
Attribute ButtonColor2.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   ButtonColor2 = m_ButtonColor2
End Property

Public Property Let ButtonColor2(ByVal New_ButtonColor2 As OLE_COLOR)
   m_ButtonColor2 = New_ButtonColor2
   PropertyChanged "ButtonColor2"
   CalculateScrollbarButtonGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get ThumbColor1() As OLE_COLOR
Attribute ThumbColor1.VB_Description = "The first gradient color of the vertical scrollbar thumb."
Attribute ThumbColor1.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   ThumbColor1 = m_ThumbColor1
End Property

Public Property Let ThumbColor1(ByVal New_ThumbColor1 As OLE_COLOR)
   m_ThumbColor1 = New_ThumbColor1
   PropertyChanged "ThumbColor1"
   CalculateScrollbarThumbGradient
   DisplayVerticalScrollBar
   UserControl.Refresh
End Property

Public Property Get ThumbColor2() As OLE_COLOR
Attribute ThumbColor2.VB_Description = "The second gradient color of the vertical scrollbar thumb."
Attribute ThumbColor2.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   ThumbColor2 = m_ThumbColor2
End Property

Public Property Let ThumbColor2(ByVal New_ThumbColor2 As OLE_COLOR)
   m_ThumbColor2 = New_ThumbColor2
   PropertyChanged "ThumbColor2"
   CalculateScrollbarThumbGradient
   DisplayVerticalScrollBar
   UserControl.Refresh
End Property

Public Property Get ThumbBorderColor() As OLE_COLOR
Attribute ThumbBorderColor.VB_Description = "The color of the 1-pixel wide border surrounding the vertical scrollbar thumb."
Attribute ThumbBorderColor.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   ThumbBorderColor = m_ThumbBorderColor
End Property

Public Property Let ThumbBorderColor(ByVal New_ThumbBorderColor As OLE_COLOR)
   m_ThumbBorderColor = New_ThumbBorderColor
   PropertyChanged "ThumbBorderColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get ArrowUpColor() As OLE_COLOR
Attribute ArrowUpColor.VB_Description = "The color of the vertical scrollbar up/down arrow when the button is not clicked."
Attribute ArrowUpColor.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   ArrowUpColor = m_ArrowUpColor
End Property

Public Property Let ArrowUpColor(ByVal New_ArrowUpColor As OLE_COLOR)
   m_ArrowUpColor = New_ArrowUpColor
   PropertyChanged "ArrowUpColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get ArrowDownColor() As OLE_COLOR
Attribute ArrowDownColor.VB_Description = "The color of the vertical scrollbar up/down arrow when the button is clicked."
Attribute ArrowDownColor.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   ArrowDownColor = m_ArrowDownColor
End Property

Public Property Let ArrowDownColor(ByVal New_ArrowDownColor As OLE_COLOR)
   m_ArrowDownColor = New_ArrowDownColor
   PropertyChanged "ArrowDownColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CheckboxArrowColor() As OLE_COLOR
Attribute CheckboxArrowColor.VB_Description = "The color of the CheckBox arrow when a list item is selected in CheckBox mode."
Attribute CheckboxArrowColor.VB_ProcData.VB_Invoke_Property = ";CheckBox"
   CheckboxArrowColor = m_CheckboxArrowColor
End Property

Public Property Let CheckboxArrowColor(ByVal New_CheckboxArrowColor As OLE_COLOR)
   m_CheckboxArrowColor = New_CheckboxArrowColor
   PropertyChanged "CheckboxArrowColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CheckBoxColor() As OLE_COLOR
Attribute CheckBoxColor.VB_Description = "The color of the CheckBox in CheckBox mode."
Attribute CheckBoxColor.VB_ProcData.VB_Invoke_Property = ";CheckBox"
   CheckBoxColor = m_CheckBoxColor
End Property

Public Property Let CheckBoxColor(ByVal New_CheckBoxColor As OLE_COLOR)
   m_CheckBoxColor = New_CheckBoxColor
   PropertyChanged "CheckBoxColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get FocusRectColor() As OLE_COLOR
Attribute FocusRectColor.VB_Description = "The color of the 1-pixel custom focus rectangle."
Attribute FocusRectColor.VB_ProcData.VB_Invoke_Property = ";Text"
   FocusRectColor = m_FocusRectColor
End Property

Public Property Let FocusRectColor(ByVal New_FocusRectColor As OLE_COLOR)
   m_FocusRectColor = New_FocusRectColor
   PropertyChanged "FocusRectColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get DblClickBehavior() As DblClickBehaviorOptions
Attribute DblClickBehavior.VB_Description = "If 0, double-clicking an item returns a DblClick event.  If 1, double-clicking item returns two single Click events."
Attribute DblClickBehavior.VB_ProcData.VB_Invoke_Property = ";Behavior"
   DblClickBehavior = m_DblClickBehavior
End Property

Public Property Let DblClickBehavior(ByVal New_DblClickBehavior As DblClickBehaviorOptions)
'  this makes the property read-only at runtime.
   If Ambient.UserMode Then Err.Raise 382
   m_DblClickBehavior = New_DblClickBehavior
   PropertyChanged "DblClickBehavior"
End Property

Public Property Get NewIndex() As Long
Attribute NewIndex.VB_Description = "The index of the most recently added list item.  If the list is empty, or the most recent action was removing a list item, -1 is returned."
   NewIndex = m_NewIndex
End Property

Public Property Let NewIndex(ByVal New_NewIndex As Long)
'  this makes the property unavailable at design time.
   If Ambient.UserMode = False Then Err.Raise 387
   m_NewIndex = New_NewIndex
   PropertyChanged "NewIndex"
End Property

Public Property Get PictureMode() As PictureModeOptions
Attribute PictureMode.VB_Description = "Allows user to display picture background in normal, tiled, or stretch-to-fit manner."
Attribute PictureMode.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   If Ambient.UserMode Then Err.Raise 393
   PictureMode = m_PictureMode
End Property

Public Property Let PictureMode(ByVal New_PictureMode As PictureModeOptions)
'  not available at runtime.
   If Ambient.UserMode Then Err.Raise 382
   m_PictureMode = New_PictureMode
   RedrawControl
   UserControl.Refresh
   PropertyChanged "PictureMode"
End Property

Public Property Get CheckStyle() As CheckStyleOptions
Attribute CheckStyle.VB_Description = "Style of checkmark in CheckBox mode (tick or check)."
Attribute CheckStyle.VB_ProcData.VB_Invoke_Property = ";CheckBox"
   CheckStyle = m_CheckStyle
End Property

Public Property Let CheckStyle(ByVal New_CheckStyle As CheckStyleOptions)
   m_CheckStyle = New_CheckStyle
   PropertyChanged "CheckStyle"
End Property

Public Property Get DragEnabled() As Boolean
Attribute DragEnabled.VB_Description = "If set to True, allows drag operations to take place."
Attribute DragEnabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
   DragEnabled = m_DragEnabled
End Property

Public Property Let DragEnabled(ByVal New_DragEnabled As Boolean)
   m_DragEnabled = New_DragEnabled
   PropertyChanged "DragEnabled"
End Property

Public Property Get TrackClickColor1() As OLE_COLOR
Attribute TrackClickColor1.VB_Description = "First gradient color of portion of scroll track above or below thumb when it is clicked."
Attribute TrackClickColor1.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   TrackClickColor1 = m_TrackClickColor1
End Property

Public Property Let TrackClickColor1(ByVal New_TrackClickColor1 As OLE_COLOR)
   m_TrackClickColor1 = New_TrackClickColor1
   PropertyChanged "TrackClickColor1"
   CalculateVerticalTrackbarGradientClicked
   DisplayVerticalTrackBar
   UserControl.Refresh
End Property

Public Property Get TrackClickColor2() As OLE_COLOR
Attribute TrackClickColor2.VB_Description = "Second gradient color of portion of scroll track above or below thumb when it is clicked."
Attribute TrackClickColor2.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   TrackClickColor2 = m_TrackClickColor2
End Property

Public Property Let TrackClickColor2(ByVal New_TrackClickColor2 As OLE_COLOR)
   m_TrackClickColor2 = New_TrackClickColor2
   PropertyChanged "TrackClickColor2"
   CalculateVerticalTrackbarGradientClicked
   DisplayVerticalTrackBar
   UserControl.Refresh
End Property

Public Property Get Theme() As ThemeOptions
   Theme = m_Theme
End Property

Public Property Get SortAsNumeric() As Boolean
Attribute SortAsNumeric.VB_Description = "When True, forces listbox to sort numeric list items in numeric order as opposed to string order (although list items are still stored as strings)."
Attribute SortAsNumeric.VB_ProcData.VB_Invoke_Property = ";Behavior"
   SortAsNumeric = m_SortAsNumeric
End Property

Public Property Let SortAsNumeric(ByVal New_SortAsNumeric As Boolean)
   m_SortAsNumeric = New_SortAsNumeric
   PropertyChanged "SortAsNumeric"
End Property

Public Property Get ShowItemImages() As Boolean
Attribute ShowItemImages.VB_Description = "If True, displays user-selected icons to the left of each listitem."
Attribute ShowItemImages.VB_ProcData.VB_Invoke_Property = ";Text"
   ShowItemImages = m_ShowItemImages
End Property

Public Property Let ShowItemImages(ByVal New_ShowItemImages As Boolean)
   m_ShowItemImages = New_ShowItemImages
   InitListBoxDisplayCharacteristics
   RedrawControl
   PropertyChanged "ShowItemImages"
End Property

Public Property Get ItemImageSize() As Long
Attribute ItemImageSize.VB_Description = "The height and width, in pixels, of the icon displayed next to listitems.  If 0, height and width are compressed to the height of listitem text."
Attribute ItemImageSize.VB_ProcData.VB_Invoke_Property = ";Text"
   ItemImageSize = m_ItemImageSize
End Property

Public Property Let ItemImageSize(ByVal New_ItemImageSize As Long)
   m_ItemImageSize = New_ItemImageSize
   InitListBoxDisplayCharacteristics
   RedrawControl
   PropertyChanged "ItemImageSize"
End Property

Public Property Get ImageIndex(ByVal Index As Long) As Long
   ImageIndex = ImageIndexArray(Index)
End Property

Public Property Let ImageIndex(ByVal Index As Long, New_ImageIndex As Long)
   ImageIndexArray(Index) = New_ImageIndex
   If m_RedrawFlag And InDisplayedItemRange(Index) Then
      DisplayList
   End If
End Property

Public Property Let Theme(ByVal New_Theme As ThemeOptions)
Attribute Theme.VB_Description = "A list of predefined color schemes that can be selected by the user."
Attribute Theme.VB_ProcData.VB_Invoke_PropertyPut = ";Main Graphics"

'*************************************************************************
'* changes color scheme of listbox to one of eight predefined themes.    *
'*************************************************************************

   m_Theme = New_Theme

   Select Case m_Theme

      Case [Cyan Eyed]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H404000
         m_ArrowUpColor = &HFFFF00
         m_BackAngle = 45
         m_BackColor1 = &H808000
         m_BackColor2 = &HFFFF80
         m_BackMiddleOut = True
         m_BorderColor = &H404000
         m_BorderWidth = 1
         m_ButtonColor1 = &H404000
         m_ButtonColor2 = &H808000
         m_CheckBoxColor = &H404000
         m_CheckboxArrowColor = &H404000
         m_FocusRectColor = &HFFFFC0
         m_SelColor1 = &H808000
         m_SelColor2 = &H808000
         m_SelTextColor = &HFFFFC0
         m_TextColor = &H0
         m_ThumbBorderColor = &HFFFF00
         m_ThumbColor1 = &H404000
         m_ThumbColor2 = &H808000
         m_TrackBarColor1 = &H808000
         m_TrackBarColor2 = &HFFFFC0
         m_TrackClickColor1 = &H404000
         m_TrackClickColor2 = &HFFFF80
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0

      Case [Gunmetal Grey]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H0
         m_ArrowUpColor = &HE0E0E0
         m_BackAngle = 45
         m_BackColor1 = &H606060
         m_BackColor2 = &HE0E0E0
         m_BackMiddleOut = True
         m_BorderColor = &H0
         m_BorderWidth = 1
         m_ButtonColor1 = &H0
         m_ButtonColor2 = &HC0C0C0
         m_CheckBoxColor = &H0
         m_CheckboxArrowColor = &H0
         m_FocusRectColor = &HFFFFFF
         m_SelColor1 = &H404040
         m_SelColor2 = &H404040
         m_SelTextColor = &HE0E0E0
         m_TextColor = &H0
         m_ThumbBorderColor = &HE0E0E0
         m_ThumbColor1 = &H0
         m_ThumbColor2 = &H909090
         m_TrackBarColor1 = &H606060
         m_TrackBarColor2 = &HE0E0E0
         m_TrackClickColor1 = &H0
         m_TrackClickColor2 = &HE0E0E0
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0

      Case [Blue Moon]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H400000
         m_ArrowUpColor = &HFFC0C0
         m_BackAngle = 45
         m_BackColor1 = &HC00000
         m_BackColor2 = &HFFC0C0
         m_BackMiddleOut = True
         m_BorderColor = &H400000
         m_BorderWidth = 1
         m_ButtonColor1 = &H400000
         m_ButtonColor2 = &HFF8080
         m_CheckBoxColor = &H400000
         m_CheckboxArrowColor = &H400000
         m_FocusRectColor = &HFFC0C0
         m_SelColor1 = &H800000
         m_SelColor2 = &H800000
         m_SelTextColor = &HFFC0C0
         m_TextColor = &H0
         m_ThumbBorderColor = &HFFC0C0
         m_ThumbColor1 = &H400000
         m_ThumbColor2 = &HFF8080
         m_TrackBarColor1 = &H800000
         m_TrackBarColor2 = &HFFC0C0
         m_TrackClickColor1 = &H400000
         m_TrackClickColor2 = &HFF8080
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0

      Case [Red Rum]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H40
         m_ArrowUpColor = &HC0C0FF
         m_BackAngle = 45
         m_BackColor1 = &H80
         m_BackColor2 = &HC0C0FF
         m_BackMiddleOut = True
         m_BorderColor = &H40
         m_BorderWidth = 1
         m_ButtonColor1 = &H40
         m_ButtonColor2 = &H8080FF
         m_CheckBoxColor = &H40
         m_CheckboxArrowColor = &H40&
         m_FocusRectColor = &HC0C0FF
         m_SelColor1 = &H80&
         m_SelColor2 = &H80&
         m_SelTextColor = &HC0C0FF
         m_TextColor = &H0
         m_ThumbBorderColor = &HC0C0FF
         m_ThumbColor1 = &H40
         m_ThumbColor2 = &H8080FF
         m_TrackBarColor1 = &H80
         m_TrackBarColor2 = &HC0C0FF
         m_TrackClickColor1 = &H40&
         m_TrackClickColor2 = &H8080FF
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0

      Case [Green With Envy]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H4000&
         m_ArrowUpColor = &HC0FFC0
         m_BackAngle = 45
         m_BackColor1 = &H8000&
         m_BackColor2 = &HC0FFC0
         m_BackMiddleOut = True
         m_BorderColor = &H4000&
         m_BorderWidth = 1
         m_ButtonColor1 = &H4000&
         m_ButtonColor2 = &H80FF80
         m_CheckBoxColor = &H4000&
         m_CheckboxArrowColor = &H4000&
         m_FocusRectColor = &HC0FFC0
         m_SelColor1 = &H8000&
         m_SelColor2 = &H8000&
         m_SelTextColor = &HC0FFC0
         m_TextColor = &H0
         m_ThumbBorderColor = &HC0FFC0
         m_ThumbColor1 = &H4000&
         m_ThumbColor2 = &HFF00&
         m_TrackBarColor1 = &H8000&
         m_TrackBarColor2 = &HC0FFC0
         m_TrackClickColor1 = &H4000&
         m_TrackClickColor2 = &H80FF80
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0

      Case [Purple People Eater]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H800080
         m_ArrowUpColor = &HFF00FF
         m_BackAngle = 45
         m_BackColor1 = &H800080
         m_BackColor2 = &HFF80FF
         m_BackMiddleOut = True
         m_BorderColor = &H400040
         m_BorderWidth = 1
         m_ButtonColor1 = &H400040
         m_ButtonColor2 = &H800080
         m_CheckBoxColor = &H400040
         m_CheckboxArrowColor = &H400040
         m_FocusRectColor = &HFFC0FF
         m_SelColor1 = &H800080
         m_SelColor2 = &H800080
         m_SelTextColor = &HFFC0FF
         m_TextColor = &H0
         m_ThumbBorderColor = &HFF00FF
         m_ThumbColor1 = &H400040
         m_ThumbColor2 = &H800080
         m_TrackBarColor1 = &H800080
         m_TrackBarColor2 = &HFFC0FF
         m_TrackClickColor1 = &H400040
         m_TrackClickColor2 = &HFF80FF
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0

      Case [Golden Goose]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H8080&
         m_ArrowUpColor = &HFFFF&
         m_BackAngle = 45
         m_BackColor1 = &H8080&
         m_BackColor2 = &H80FFFF
         m_BackMiddleOut = True
         m_BorderColor = &H4040&
         m_BorderWidth = 1
         m_ButtonColor1 = &H4040&
         m_ButtonColor2 = &H8080&
         m_CheckBoxColor = &H4040&
         m_CheckboxArrowColor = &H4040&
         m_FocusRectColor = &HC0FFFF
         m_SelColor1 = &H8080&
         m_SelColor2 = &H8080&
         m_SelTextColor = &HC0FFFF
         m_TextColor = &H0
         m_ThumbBorderColor = &HFFFF&
         m_ThumbColor1 = &H4040&
         m_ThumbColor2 = &H8080&
         m_TrackBarColor1 = &H8080&
         m_TrackBarColor2 = &HC0FFFF
         m_TrackClickColor1 = &H4040&
         m_TrackClickColor2 = &H80FFFF
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0

      Case [Penny Wise]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H4080&
         m_ArrowUpColor = &HC0E0FF
         m_BackAngle = 45
         m_BackColor1 = &H4080&
         m_BackColor2 = &H80C0FF
         m_BackMiddleOut = True
         m_BorderColor = &H404080
         m_BorderWidth = 1
         m_ButtonColor1 = &H404080
         m_ButtonColor2 = &H80FF&
         m_CheckBoxColor = &H0
         m_CheckboxArrowColor = &H0
         m_FocusRectColor = &HC0E0FF
         m_SelColor1 = &H4080&
         m_SelColor2 = &H4080&
         m_SelTextColor = &HC0E0FF
         m_TextColor = &H0
         m_ThumbBorderColor = &H80FF&
         m_ThumbColor1 = &H404080
         m_ThumbColor2 = &H40C0&
         m_TrackBarColor1 = &H4080&
         m_TrackBarColor2 = &HC0E0FF
         m_TrackClickColor1 = &H404080
         m_TrackClickColor2 = &H80C0FF
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0

   End Select

   PropertyChanged "Theme"

   GetEnabledDisplayProperties
   CalculateGradients
   RedrawControl
   UserControl.Refresh

End Property

Public Property Get TopIndex() As Long
Attribute TopIndex.VB_Description = "The index of the list item that is displayed at the top of the listbox."
Attribute TopIndex.VB_ProcData.VB_Invoke_Property = ";Text"
'  Note:  In the standard VB listbox, the .TopIndex property is read/write, and you can set
'  the property to display list items beginning with the specified index.  In this control,
'  the .DisplayFrom method replaces the .TopIndex write functionality. Therefore, .TopIndex
'  here is read-only and returns the index of the first displayed list item.
   TopIndex = m_TopIndex
End Property

Public Property Let TopIndex(ByVal New_TopIndex As Long)
   If Ambient.UserMode = False Then Err.Raise 387    ' property is not available at design time.
   If Ambient.UserMode Then Err.Raise 382            ' property is read-only at runtime.
   m_TopIndex = New_TopIndex
   PropertyChanged "TopIndex"
End Property

Public Property Get ScaleWidth() As Long
   ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Get ScaleMode() As Integer
   ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
   m_ScaleMode = New_ScaleMode
   UserControl.ScaleMode = m_ScaleMode
   PropertyChanged "ScaleMode"
End Property

Public Property Get ScaleHeight() As Long
   ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Get AutoRedraw() As Boolean
   AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
   m_AutoRedraw = New_AutoRedraw
   UserControl.AutoRedraw = m_AutoRedraw
   PropertyChanged "AutoRedraw"
End Property

Private Sub GetEnabledDisplayProperties()

'*************************************************************************
'* applies enabled graphics properties to the active display properties. *
'*************************************************************************

   Set m_ActivePicture = m_Picture
   m_ActiveArrowDownColor = m_ArrowDownColor
   m_ActiveArrowUpColor = m_ArrowUpColor
   m_ActiveBackColor1 = m_BackColor1
   m_ActiveBackColor2 = m_BackColor2
   m_ActiveBorderColor = m_BorderColor
   m_ActiveButtonColor1 = m_ButtonColor1
   m_ActiveButtonColor2 = m_ButtonColor2
   m_ActiveCheckboxArrowColor = m_CheckboxArrowColor
   m_ActiveCheckBoxColor = m_CheckBoxColor
   m_ActiveFocusRectColor = m_FocusRectColor
   m_ActivePictureMode = m_PictureMode
   m_ActiveSelColor1 = m_SelColor1
   m_ActiveSelColor2 = m_SelColor2
   m_ActiveSelTextColor = m_SelTextColor
   m_ActiveTextColor = m_TextColor
   m_ActiveThumbBorderColor = m_ThumbBorderColor
   m_ActiveThumbColor1 = m_ThumbColor1
   m_ActiveThumbColor2 = m_ThumbColor2
   m_ActiveTrackBarColor1 = m_TrackBarColor1
   m_ActiveTrackBarColor2 = m_TrackBarColor2

End Sub

Private Sub GetDisabledDisplayProperties()

'*************************************************************************
'* applies disabled graphics properties to active display properties.    *
'*************************************************************************

   Set m_ActivePicture = m_DisPicture
   m_ActiveArrowDownColor = m_DisArrowDownColor
   m_ActiveArrowUpColor = m_DisArrowUpColor
   m_ActiveBackColor1 = m_DisBackColor1
   m_ActiveBackColor2 = m_DisBackColor2
   m_ActiveBorderColor = m_DisBorderColor
   m_ActiveButtonColor1 = m_DisButtonColor1
   m_ActiveButtonColor2 = m_DisButtonColor2
   m_ActiveCheckboxArrowColor = m_DisCheckboxArrowColor
   m_ActiveCheckBoxColor = m_DisCheckboxColor
   m_ActiveFocusRectColor = m_DisFocusRectColor
   m_ActivePictureMode = m_DisPictureMode
   m_ActiveSelColor1 = m_DisSelColor1
   m_ActiveSelColor2 = m_DisSelColor2
   m_ActiveSelTextColor = m_DisSelTextColor
   m_ActiveTextColor = m_DisTextColor
   m_ActiveThumbBorderColor = m_DisThumbBorderColor
   m_ActiveThumbColor1 = m_DisThumbColor1
   m_ActiveThumbColor2 = m_DisThumbColor2
   m_ActiveTrackBarColor1 = m_DisTrackbarColor1
   m_ActiveTrackBarColor2 = m_DisTrackbarColor2

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Subclassing >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<< All subclassing code by Paul Caton. >>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub StartSubclassing()

'*************************************************************************
'* starts up Paul Caton's self-subclassing code.                         *
'*************************************************************************

   If Ambient.UserMode Then                                    ' if we're not in design mode.
      With UserControl
         Call Subclass_Start(.hwnd)                            ' Start subclassing.
         Call Subclass_AddMsg(.hwnd, WM_MOUSEMOVE, MSG_AFTER)  ' for mouse enter detect.
         Call Subclass_AddMsg(.hwnd, WM_MOUSELEAVE, MSG_AFTER) ' for mouse leave detect.
         Call Subclass_AddMsg(.hwnd, WM_MOUSEWHEEL, MSG_AFTER) ' for mouse wheel detect.
         Call Subclass_AddMsg(.hwnd, WM_SETFOCUS, MSG_AFTER)   ' for got focus detect.
         Call Subclass_AddMsg(.hwnd, WM_KILLFOCUS, MSG_AFTER)  ' for lost focus detect.
      End With
   End If

End Sub

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

'*************************************************************************
'* track the mouse leaving the indicated window.                         *
'*************************************************************************

   Dim tme As TRACKMOUSEEVENT_STRUCT

   With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
   End With

   Call TrackMouseEventComCtl(tme)

End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
   With sc_aSubData(zIdx(lng_hWnd))
      If When And eMsgWhen.MSG_BEFORE Then
         Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
      End If
      If When And eMsgWhen.MSG_AFTER Then
         Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
      End If
   End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
   Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long

'  Parameters:
'  lng_hWnd  - The handle of the window to be subclassed
'  Returns;
'  The sc_aSubData() index
   Const CODE_LEN              As Long = 200                      'Length of the machine code in bytes
   Const FUNC_CWP              As String = "CallWindowProcA"      'We use CallWindowProc to call the original WndProc
   Const FUNC_EBM              As String = "EbMode"               'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
   Const FUNC_SWL              As String = "SetWindowLongA"       'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
   Const MOD_USER              As String = "user32"               'Location of the SetWindowLongA & CallWindowProc functions
   Const MOD_VBA5              As String = "vba5"                 'Location of the EbMode function if running VB5
   Const MOD_VBA6              As String = "vba6"                 'Location of the EbMode function if running VB6
   Const PATCH_01              As Long = 18                       'Code buffer offset to the location of the relative address to EbMode
   Const PATCH_02              As Long = 68                       'Address of the previous WndProc
   Const PATCH_03              As Long = 78                       'Relative address of SetWindowsLong
   Const PATCH_06              As Long = 116                      'Address of the previous WndProc
   Const PATCH_07              As Long = 121                      'Relative address of CallWindowProc
   Const PATCH_0A              As Long = 186                      'Address of the owner object
   Static aBuf(1 To CODE_LEN)  As Byte                            'Static code buffer byte array
   Static pCWP                 As Long                            'Address of the CallWindowsProc
   Static pEbMode              As Long                            'Address of the EbMode IDE break/stop/running function
   Static pSWL                 As Long                            'Address of the SetWindowsLong function
   Dim i                       As Long                            'Loop index
   Dim j                       As Long                            'Loop index
   Dim nSubIdx                 As Long                            'Subclass data index
   Dim sHex                    As String                          'Hex code string

'  If it's the first time through here..
   If aBuf(1) = 0 Then

'     The hex pair machine code representation.
      sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
             "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
             "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
             "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'     Convert the string from hex pairs to bytes and store in the static machine code buffer
      i = 1
      Do While j < CODE_LEN
         j = j + 1
         aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                   'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
         i = i + 2
      Loop                                                        'Next pair of hex characters

'     Get API function addresses
      If Subclass_InIDE Then                                      'If we're running in the VB IDE
         aBuf(16) = &H90                                          'Patch the code buffer to enable the IDE state code
         aBuf(17) = &H90                                          'Patch the code buffer to enable the IDE state code
         pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                  'Get the address of EbMode in vba6.dll
         If pEbMode = 0 Then                                      'Found?
            pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)               'VB5 perhaps
         End If
      End If

      pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                        'Get the address of the CallWindowsProc function
      pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                        'Get the address of the SetWindowLongA function
      ReDim sc_aSubData(0 To 0) As tSubData                       'Create the first sc_aSubData element

   Else

      nSubIdx = zIdx(lng_hWnd, True)
      If nSubIdx = -1 Then                                        'If an sc_aSubData element isn't being re-cycled
         nSubIdx = UBound(sc_aSubData()) + 1                      'Calculate the next element
         ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData     'Create a new sc_aSubData element
      End If

      Subclass_Start = nSubIdx

   End If

   With sc_aSubData(nSubIdx)
      .hwnd = lng_hWnd                                            'Store the hWnd
      .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)               'Allocate memory for the machine code WndProc
      .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)  'Set our WndProc in place
      Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)      'Copy the machine code from the static byte array to the code array in sc_aSubData
      Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
      Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)             'Original WndProc address for CallWindowProc, call the original WndProc
      Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                   'Patch the relative address of the SetWindowLongA api function
      Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)             'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
      Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                   'Patch the relative address of the CallWindowProc api function
      Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))             'Patch the address of this object instance into the static machine code buffer
   End With

End Function

'Stop all subclassing
Private Sub Subclass_StopAll()

   Dim i As Long

   i = UBound(sc_aSubData())                                      'Get the upper bound of the subclass data array
   Do While i >= 0                                                'Iterate through each element
      With sc_aSubData(i)
         If .hwnd <> 0 Then                                       'If not previously Subclass_Stop'd
            Call Subclass_Stop(.hwnd)                             'Subclass_Stop
         End If
      End With
      i = i - 1                                                   'Next element
   Loop

End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)

'  Parameters:
'  lng_hWnd  - The handle of the window to stop being subclassed
   With sc_aSubData(zIdx(lng_hWnd))
      Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)         'Restore the original WndProc
      Call zPatchVal(.nAddrSub, PATCH_05, 0)                      'Patch the Table B entry count to ensure no further 'before' callbacks
      Call zPatchVal(.nAddrSub, PATCH_09, 0)                      'Patch the Table A entry count to ensure no further 'after' callbacks
      Call GlobalFree(.nAddrSub)                                  'Release the machine code memory
      .hwnd = 0                                                   'Mark the sc_aSubData element as available for re-use
      .nMsgCntB = 0                                               'Clear the before table
      .nMsgCntA = 0                                               'Clear the after table
      Erase .aMsgTblB                                             'Erase the before table
      Erase .aMsgTblA                                             'Erase the after table
   End With

End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)

   Dim nEntry  As Long                                            'Message table entry index
   Dim nOff1   As Long                                            'Machine code buffer offset 1
   Dim nOff2   As Long                                            'Machine code buffer offset 2

   If uMsg = ALL_MESSAGES Then                                    'If all messages
      nMsgCnt = ALL_MESSAGES                                      'Indicates that all messages will callback
   Else                                                           'Else a specific message number
      Do While nEntry < nMsgCnt                                   'For each existing entry. NB will skip if nMsgCnt = 0
         nEntry = nEntry + 1
         If aMsgTbl(nEntry) = 0 Then                              'This msg table slot is a deleted entry
            aMsgTbl(nEntry) = uMsg                                'Re-use this entry
            Exit Sub                                              'Bail
         ElseIf aMsgTbl(nEntry) = uMsg Then                       'The msg is already in the table!
            Exit Sub                                              'Bail
         End If
      Loop                                                        'Next entry
      nMsgCnt = nMsgCnt + 1                                       'New slot required, bump the table entry count
      ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                'Bump the size of the table.
      aMsgTbl(nMsgCnt) = uMsg                                     'Store the message number in the table
   End If

   If When = eMsgWhen.MSG_BEFORE Then                             'If before
      nOff1 = PATCH_04                                            'Offset to the Before table
      nOff2 = PATCH_05                                            'Offset to the Before table entry count
   Else                                                           'Else after
      nOff1 = PATCH_08                                            'Offset to the After table
      nOff2 = PATCH_09                                            'Offset to the After table entry count
   End If

   If uMsg <> ALL_MESSAGES Then
      Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))            'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
   End If

   Call zPatchVal(nAddr, nOff2, nMsgCnt)                          'Patch the appropriate table entry count

End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
   zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
   Debug.Assert zAddrFunc                                         'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long

'  Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
   zIdx = UBound(sc_aSubData)
   Do While zIdx >= 0                                             'Iterate through the existing sc_aSubData() elements
      With sc_aSubData(zIdx)
         If .hwnd = lng_hWnd Then                                 'If the hWnd of this element is the one we're looking for
            If Not bAdd Then                                      'If we're searching not adding
               Exit Function                                      'Found
            End If
         ElseIf .hwnd = 0 Then                                    'If this an element marked for reuse.
            If bAdd Then                                          'If we're adding
               Exit Function                                      'Re-use it
            End If
         End If
      End With
      zIdx = zIdx - 1                                             'Decrement the index
   Loop

  If Not bAdd Then
    Debug.Assert False                                            'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
   Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
   Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
   zSetTrue = True
   bValue = True
End Function
