VERSION 5.00
Begin VB.UserControl GradientMagicLabel 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "GradientMagicLabel.ctx":0000
   ScaleHeight     =   192
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   234
   ToolboxBitmap   =   "GradientMagicLabel.ctx":0042
   Begin VB.Shape shResize 
      Height          =   495
      Left            =   825
      Top             =   945
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "GradientMagicLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type tRGB
    R As Integer
    G As Integer
    b As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Enum typBordo
    [Normal] = 0
    [3D] = 1
End Enum

Enum VAlign
    vTop = 0
    vCenter = 1
    vBottom = 2
End Enum

Enum HAlign
    hLeft = 0
    hCenter = 1
    hRight = 2
End Enum

Enum gOrientation
   Horizontal = 0
   Vertical = 1
End Enum

Private Type RECT
    rLeft    As Long
    rTop     As Long
    rRight   As Long
    rBottom  As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String
End Type

Const GRADIENT_FILL_RECT_H As Long = &H0
Const GRADIENT_FILL_RECT_V  As Long = &H1
Const GRADIENT_FILL_TRIANGLE As Long = &H2
Const GRADIENT_FILL_OP_FLAG As Long = &HFF
Const PixelPerMillimeter As Double = 3.77933333333333

Const DT_WORDBREAK = &H10
Const DT_CALCRECT = &H400

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Dim tmpBorderStyle As typBordo
Dim m_BorderStyle As typBordo
Dim m_Caption As String
Dim m_BeginColor As OLE_COLOR
Dim m_EndColor As OLE_COLOR
Dim m_InResize As Boolean
Dim m_Gradient As Boolean
Dim m_AlignmentH As HAlign
Dim m_CaptionOffSet As Integer
Dim m_GradientAlignment As gOrientation
Dim m_AlignmentV As VAlign

Const m_def_BeginColor = &HFFFFFF
Const m_def_EndColor = &HC0FFFF
Const m_def_BorderStyle = 0
Const m_def_Caption = "No label specified"
Const m_def_InResize = False
Const m_def_Gradient = True
Const m_def_AlignmentH = 0
Const m_def_AlignmentV = 0
Const m_def_CaptionOffSet = 0
Const m_def_GradientAlignment = False

Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Viene generato quando si preme e quindi si rilascia un pulsante del mouse su un oggetto."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Viene generato quando si preme il pulsante del mouse mentre lo stato attivo si trova su un oggetto."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Viene generato quando si sposta il mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Viene generato quando si rilascia il pulsante del mouse mentre lo stato attivo si trova su un oggetto."

Const m_def_Filigree = "Filigree"
Dim m_Filigree As String




'MemberInfo=10,0,0,0
Public Property Get BeginColor() As OLE_COLOR
Attribute BeginColor.VB_Description = "Imposta il colore di partenza del gradiente."
    BeginColor = m_BeginColor
End Property

Public Property Let BeginColor(ByVal New_BeginColor As OLE_COLOR)
    m_BeginColor = New_BeginColor
    PropertyChanged "BeginColor"
    MagicLabelRedraw
End Property

'MemberInfo=10,0,0,00
Public Property Get EndColor() As OLE_COLOR
Attribute EndColor.VB_Description = "Imposta il colore finale del gradiente."
    EndColor = m_EndColor
End Property

Public Property Let EndColor(ByVal New_EndColor As OLE_COLOR)
    m_EndColor = New_EndColor
    PropertyChanged "EndColor"
    MagicLabelRedraw
End Property

Private Sub UserControl_EnterFocus()
Debug.Print "Enter Focus"
End Sub

Private Sub UserControl_GotFocus()
Debug.Print "Got Focus"

End Sub

Private Sub UserControl_InitProperties()
   Set UserControl.Font = Ambient.Font
   m_BeginColor = m_def_BeginColor
   m_EndColor = m_def_EndColor
   m_Caption = m_def_Caption
   m_BorderStyle = m_def_BorderStyle
   m_AlignmentV = m_def_AlignmentV
   m_GradientAlignment = m_def_GradientAlignment
   m_CaptionOffSet = m_def_CaptionOffSet
   m_AlignmentH = m_def_AlignmentH
   m_Gradient = m_def_Gradient
   m_InResize = m_def_InResize
   m_Filigree = m_def_Filigree
   UserControl.BorderStyle = m_BorderStyle
End Sub

Private Sub UserControl_LostFocus()
Debug.Print "Lost Focus"

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
   Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
   m_BeginColor = PropBag.ReadProperty("BeginColor", m_def_BeginColor)
   m_EndColor = PropBag.ReadProperty("EndColor", m_def_EndColor)
   m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
   m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
   m_AlignmentV = PropBag.ReadProperty("AlignmentV", m_def_AlignmentV)
   m_GradientAlignment = PropBag.ReadProperty("GradientAlignment", m_def_GradientAlignment)
   m_CaptionOffSet = PropBag.ReadProperty("CaptionOffSet", m_def_CaptionOffSet)
   m_AlignmentH = PropBag.ReadProperty("AlignmentH", m_def_AlignmentH)
   m_Gradient = PropBag.ReadProperty("Gradient", m_def_Gradient)
   m_InResize = PropBag.ReadProperty("InResize", m_def_InResize)
   m_Filigree = PropBag.ReadProperty("Filigree", m_def_Filigree)
   UserControl.BorderStyle = m_BorderStyle
   UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
   UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
   UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
End Sub

Private Sub UserControl_Resize()
MagicLabelRedraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "BeginColor", m_BeginColor, m_def_BeginColor
      .WriteProperty "EndColor", m_EndColor, m_def_EndColor
      .WriteProperty "Font", UserControl.Font, Ambient.Font
      .WriteProperty "Caption", m_Caption, m_def_Caption
      .WriteProperty "BorderStyle", m_BorderStyle, m_def_BorderStyle
      .WriteProperty "AlignmentV", m_AlignmentV, m_def_AlignmentV
      .WriteProperty "GradientAlignment", m_GradientAlignment, m_def_GradientAlignment
      .WriteProperty "CaptionOffSet", m_CaptionOffSet, m_def_CaptionOffSet
      .WriteProperty "AlignmentH", m_AlignmentH, m_def_AlignmentH
      .WriteProperty "Gradient", m_Gradient, m_def_Gradient
      .WriteProperty "ForeColor", UserControl.ForeColor, &H80000012
      .WriteProperty "MouseIcon", MouseIcon, Nothing
      .WriteProperty "MousePointer", UserControl.MousePointer, 0
      .WriteProperty "InResize", m_InResize, m_def_InResize
      .WriteProperty "AutoRedraw", UserControl.AutoRedraw, True
   End With
   Call PropBag.WriteProperty("Filigree", m_Filigree, m_def_Filigree)
End Sub

Public Function LongToUShort(Unsigned As Long) As Integer
If Unsigned <> 0 Then LongToUShort = CInt(Unsigned - &H10000)
End Function

'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Restituisce un oggetto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    MagicLabelRedraw
End Property

'MemberInfo=13,0,0,No label specified
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Testo che verrà visualizzato nel corpo dell'oggetto. NOTA: sono supportati i testi multi riga."
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    MagicLabelRedraw
End Property

Private Function SeparaCanali(Colore As OLE_COLOR) As tRGB
On Error Resume Next
Dim R, G, b As Integer
Dim Tmp As String

Tmp = Right("000000" & Hex(Colore), 6)
SeparaCanali.R = LongToUShort(Val("&H" & Right(Tmp, 2) & "00&"))
SeparaCanali.G = LongToUShort(Val("&H" & Mid(Tmp, 3, 2) & "00&"))
SeparaCanali.b = LongToUShort(Val("&H" & Left(Tmp, 2) & "00&"))
End Function

'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As typBordo
Attribute BorderStyle.VB_Description = "Imposta lo stile del bordo dell'oggetto."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As typBordo)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    UserControl.BorderStyle = New_BorderStyle
    MagicLabelRedraw
End Property

'MemberInfo=7,0,0,0
Public Property Get AlignmentV() As VAlign
Attribute AlignmentV.VB_Description = "Imposta l'allineamento verticale del testo"
    AlignmentV = m_AlignmentV
End Property

Public Property Let AlignmentV(ByVal New_AlignmentV As VAlign)
    m_AlignmentV = New_AlignmentV
    PropertyChanged "AlignmentV"
    MagicLabelRedraw
End Property

'MemberInfo=0,0,0,False
Public Property Get GradientAlignment() As gOrientation
Attribute GradientAlignment.VB_Description = "Imposta lo stile di disegno del gradiente (orizzontale o verticale)."
    GradientAlignment = m_GradientAlignment
End Property

Public Property Let GradientAlignment(ByVal New_GradientAlignment As gOrientation)
    m_GradientAlignment = New_GradientAlignment
    PropertyChanged "GradientAlignment"
    MagicLabelRedraw
End Property

'MemberInfo=7,0,0,0
Public Property Get CaptionOffSet() As Integer
Attribute CaptionOffSet.VB_Description = "Specifica (in mm) la distanza da sinistra del testo."
    CaptionOffSet = m_CaptionOffSet
End Property

Public Property Let CaptionOffSet(ByVal New_CaptionOffSet As Integer)
    m_CaptionOffSet = New_CaptionOffSet
    PropertyChanged "CaptionOffSet"
    MagicLabelRedraw
End Property

Private Sub MagicLabelRedraw()
On Error Resume Next

With UserControl
   .Cls
   
   Dim vert(1) As TRIVERTEX
   Dim gRect As GRADIENT_RECT
   
   With vert(0)
       .X = 0
       .Y = 0
       .Red = SeparaCanali(m_BeginColor).R
       .Green = SeparaCanali(m_BeginColor).G
       .Blue = SeparaCanali(m_BeginColor).b
       .Alpha = 0&
   End With
   
   With vert(1)
       .X = UserControl.ScaleWidth
       .Y = UserControl.ScaleHeight
       .Red = SeparaCanali(IIf(m_Gradient, m_EndColor, m_BeginColor)).R
       .Green = SeparaCanali(IIf(m_Gradient, m_EndColor, m_BeginColor)).G
       .Blue = SeparaCanali(IIf(m_Gradient, m_EndColor, m_BeginColor)).b
       .Alpha = 0&
   End With
   
   gRect.UpperLeft = 0
   gRect.LowerRight = 1

   GradientFillRect .hdc, vert(0), 2, gRect, 1, m_GradientAlignment
   
   'Disegno la Filigree
   SetTextColor .hdc, &HC0C0C0
   DisegnaText m_Filigree, 0, True
   
   'Disegno il Text
   SetTextColor .hdc, .ForeColor
   DisegnaText m_Caption, m_CaptionOffSet
   
   shResize.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
   shResize.Visible = InResize
End With
End Sub

'MemberInfo=7,0,0,0
Public Property Get AlignmentH() As HAlign
Attribute AlignmentH.VB_Description = "Imposta l'allineamento orizzontale del testo"
   AlignmentH = m_AlignmentH
End Property

Public Property Let AlignmentH(ByVal New_AlignmentH As HAlign)
   m_AlignmentH = New_AlignmentH
   PropertyChanged "AlignmentH"
   MagicLabelRedraw
End Property

Private Sub DisegnaText(ByVal Text As String, ByVal OffSet As Integer, Optional ByVal Filigree As Boolean = False)
   Dim vh As Long
   Dim hRect As RECT
   Dim Compensazione As Long
   Dim dFont As LOGFONT
   Dim rFont As Long, defaultFont As Long
   Dim bkpFont As New StdFont
   Dim tmpFont As New StdFont
   
   Set bkpFont = UserControl.Font
   
   'Init dFont
   If Filigree Then
      tmpFont.Name = "Arial"
      tmpFont.Size = 14
    Else
      Set tmpFont = UserControl.Font
   End If
   
   Set UserControl.Font = tmpFont
      
   With UserControl
     
      SetRect hRect, 4, 0, ScaleWidth - 4, ScaleHeight
     
      If Filigree Then
         Compensazione = OffSet * PixelPerMillimeter
         OffsetRect hRect, Compensazione, 0
         vh = DrawText(.hdc, Text, -1, hRect, DT_CALCRECT Or hCenter Or DT_WORDBREAK)
       Else
         Select Case AlignmentH
            Case hLeft:
               Compensazione = (OffSet - 1) * PixelPerMillimeter
            Case hRight:
               Compensazione = (OffSet + 1) * PixelPerMillimeter
            Case hCenter:
               Compensazione = OffSet * PixelPerMillimeter
         End Select
         Select Case AlignmentV
           Case 0: 'Top
              OffsetRect hRect, Compensazione, -hRect.rTop
           Case 1: 'Center
              OffsetRect hRect, Compensazione, 0
           Case 2: 'Right
              OffsetRect hRect, Compensazione, ScaleHeight - TextHeight(Text) - hRect.rTop
         End Select
         vh = DrawText(.hdc, Text, -1, hRect, DT_CALCRECT Or AlignmentH Or DT_WORDBREAK)
      End If
        
      SetRect hRect, 4, (ScaleHeight - vh) / 2, ScaleWidth - 4, ScaleHeight
     
      If Filigree Then
         OffsetRect hRect, Compensazione, 0
         DrawText .hdc, Text, -1, hRect, DT_WORDBREAK Or hCenter
       Else
         Select Case AlignmentV
           Case 0: 'Top
              OffsetRect hRect, Compensazione, -hRect.rTop
           Case 1: 'Center
              OffsetRect hRect, Compensazione, 0
           Case 2:
              OffsetRect hRect, Compensazione, ScaleHeight - TextHeight(Text) - hRect.rTop
         End Select
         DrawText .hdc, Text, -1, hRect, DT_WORDBREAK Or AlignmentH
      End If
   End With
   Set UserControl.Font = bkpFont
End Sub

'MemberInfo=0,0,0,True
Public Property Get Gradient() As Boolean
Attribute Gradient.VB_Description = "Attiva/disattiva la creazione dinamica del gradiente di colore."
   Gradient = m_Gradient
End Property

Public Property Let Gradient(ByVal New_Gradient As Boolean)
   m_Gradient = New_Gradient
   PropertyChanged "Gradient"
   If Not m_Gradient Then UserControl.BackColor = m_BeginColor
   MagicLabelRedraw
End Property

'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Restituisce o imposta il colore di primo piano utilizzato per la visualizzazione di testo e grafica in un oggetto."
   ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   UserControl.ForeColor() = New_ForeColor
   PropertyChanged "ForeColor"
   MagicLabelRedraw
End Property

'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Restituisce un handle (da Microsoft Windows) al contesto di periferica di un oggetto."
   hdc = UserControl.hdc
End Property

'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Restituisce un handle (da Microsoft Windows) alla finestra di un oggetto."
   hWnd = UserControl.hWnd
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Imposta un'icona personalizzata per il puntatore del mouse."
   Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
   Set UserControl.MouseIcon = New_MouseIcon
   PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Restituisce o imposta il tipo di puntatore del mouse visualizzato quando il puntatore si trova su una parte specifica di un oggetto."
   MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
   UserControl.MousePointer() = New_MousePointer
   PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

'MemberInfo=0,0,2,False
Public Property Get InResize() As Boolean
Attribute InResize.VB_MemberFlags = "400"
   InResize = m_InResize
End Property

Public Property Let InResize(ByVal New_InResize As Boolean)
   If Ambient.UserMode = False Then Err.Raise 387
   m_InResize = New_InResize
   PropertyChanged "InResize"
   
   With UserControl
     If Not New_InResize Then
         MagicLabelRedraw
         .BorderStyle = tmpBorderStyle
         .BackStyle = 1
      Else
         tmpBorderStyle = m_BorderStyle
         .BorderStyle = 0
         .BackStyle = 0
     End If
   End With
End Property

'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Restituisce o imposta l'output di un metodo grafico in una bitmap fissa."
   AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
   UserControl.AutoRedraw() = New_AutoRedraw
   PropertyChanged "AutoRedraw"
End Property

'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Ridisegna completamente un oggetto."
   MagicLabelRedraw
   UserControl.Refresh
End Sub

'MemberInfo=13,0,0,Filigree
Public Property Get Filigree() As String
   Filigree = m_Filigree
End Property

Public Property Let Filigree(ByVal New_Filigree As String)
   m_Filigree = New_Filigree
   PropertyChanged "Filigree"
   MagicLabelRedraw
End Property

