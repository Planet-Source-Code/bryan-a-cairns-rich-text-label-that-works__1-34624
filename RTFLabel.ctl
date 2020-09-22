VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl RTFLabel 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3413
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"RTFLabel.ctx":0000
   End
End
Attribute VB_Name = "RTFLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Click()
Event DblClick()

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const SRCCOPY = &HCC0020
Private Const WM_USER = &H400
Private Const EM_SETTARGETDEVICE = (WM_USER + 72)

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Dim XX As Single
Dim YY As Single

Dim POSX As Single
Dim POSY As Single
Dim POSW As Single
Dim POSH As Single
Dim SID As String
Dim m_AllowMove As Boolean

Public Sub DisplayMode()
    RichTextBox1.Visible = False
    Picture1.Visible = True
    CopyBackGround
End Sub

Public Sub EditMode()
    RichTextBox1.Visible = True
    Picture1.Visible = False
End Sub

Public Sub RenderText()
'this is where we actually draw the text on the control

Dim VDC As VirtualDC
Dim bWrapText As Boolean
Dim LColor As Long
Set VDC = New VirtualDC
LColor = VDC.GetColor(RichTextBox1.BackColor)
'make sure the RTF box is NOT in word wrap mode
bWrapText = False
SendMessageLong RichTextBox1.HWND, EM_SETTARGETDEVICE, 0, CLng(Not bWrapText)
'create a Virtual Device Context that we can draw on

VDC.Create UserControl.hdc, RichTextBox1.Width / Screen.TwipsPerPixelX, RichTextBox1.Height / Screen.TwipsPerPixelY
VDC.Cls
VDC.BackColor = LColor
VDC.ForeColor = LColor
VDC.FillArea 0, 0, VDC.Width, VDC.Height, LColor

'RichTextBox1.LoadFile App.Path & "\Document.rtf", 0
'Make sure the user has no text selected or it will only output the selected text
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = 0
RichTextBox1.SelPrint VDC.hdc, 0
TransparentBlt Picture1.hdc, 0, 0, Picture1.Width / 15, Picture1.Height / 15, VDC.hdc, 0, 0, VDC.Width, VDC.Height, LColor
Picture1.Refresh
End Sub

Public Sub UpdateMoving()
On Error Resume Next
If m_AllowMove = False Then Exit Sub
If HasBeenMoved = False Then Exit Sub
UserControl.Parent.AutoRedraw = False
CopyBackGround
UserControl.Parent.AutoRedraw = True
End Sub

Private Function HasBeenMoved() As Boolean
On Error Resume Next
Dim OBJ As Object
Dim bOK As Boolean
bOK = False
Set OBJ = GetUserControl
    If POSX <> OBJ.Left Or POSY <> OBJ.Top Or POSW <> OBJ.Width Or POSH <> OBJ.Height Then
        POSX = OBJ.Left
        POSY = OBJ.Top
        POSW = OBJ.Width
        POSH = OBJ.Height
        bOK = True
    End If
    HasBeenMoved = bOK
End Function


Public Sub CopyBackGround()
On Error Resume Next
Dim OBJ As Object
Dim rc As RECT
Set OBJ = GetUserControl
If OBJ Is Nothing Then Exit Sub

With rc
.Left = OBJ.Left / Screen.TwipsPerPixelX
.Top = OBJ.Top / Screen.TwipsPerPixelY
.Right = OBJ.Width / Screen.TwipsPerPixelX
.Bottom = OBJ.Height / Screen.TwipsPerPixelY
End With
OBJ.Visible = False
BitBlt Picture1.hdc, 0&, 0&, rc.Right, rc.Bottom, UserControl.Parent.hdc, rc.Left, rc.Top, vbSrcCopy
Picture1.Refresh
OBJ.Visible = True
Picture1.Picture = Picture1.Image
Picture1.Cls
RenderText
End Sub

Private Function GetUserControl() As Object

Dim i As Long
Dim OBJ As Object
On Error Resume Next
For i = 0 To UserControl.ParentControls.Count - 1
    If TypeOf UserControl.ParentControls.Item(i) Is RTFLabel Then
    If UserControl.ParentControls.Item(i).GetID = SID Then
    Set OBJ = UserControl.ParentControls.Item(i)
        Exit For
    End If
    End If
Next i
Set GetUserControl = OBJ
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    XX = X
    YY = Y
    If Button = 1 Then MoveBoundControl X, Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If Button = 1 Then MoveBoundControl X, Y
    
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button = 1 Then MoveBoundControl X, Y
End Sub

Private Sub UserControl_Initialize()
CopyBackGround
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Picture1.Width = UserControl.Width
Picture1.Height = UserControl.Height
RichTextBox1.Width = UserControl.Width
RichTextBox1.Height = UserControl.Height
CopyBackGround
End Sub

Private Sub MoveBoundControl(X As Single, Y As Single)
If m_AllowMove = False Then Exit Sub
Dim i As Long
Dim OBJ As Object
On Error Resume Next
Set OBJ = GetUserControl

If OBJ Is Nothing Then Exit Sub
    OBJ.Move OBJ.Left + (X - XX), OBJ.Top + (Y - YY)
    
CopyBackGround

End Sub


Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Picture1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Picture1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get CurrentX() As Single
Attribute CurrentX.VB_Description = "Returns/sets the horizontal coordinates for next print or draw method."
    CurrentX = Picture1.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
    Picture1.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property

Public Property Get CurrentY() As Single
Attribute CurrentY.VB_Description = "Returns/sets the vertical coordinates for next print or draw method."
    CurrentY = Picture1.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
    Picture1.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property

Private Sub Picture1_Click()
    RaiseEvent Click
End Sub

Private Sub Picture1_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = Picture1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Picture1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = Picture1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Picture1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = Picture1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Picture1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = Picture1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Picture1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = Picture1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Picture1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

Public Property Get FontTransparent() As Boolean
Attribute FontTransparent.VB_Description = "Returns/sets a value that determines whether background text/graphics on a Form, Printer or PictureBox are displayed."
    FontTransparent = Picture1.FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    Picture1.FontTransparent() = New_FontTransparent
    PropertyChanged "FontTransparent"
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = Picture1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Picture1.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Picture1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Picture1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Picture1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Picture1.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get TextRTF() As String
    TextRTF = RichTextBox1.TextRTF
End Property

Public Property Let TextRTF(ByVal New_TextRTF As String)
    PropertyChanged "TextRTF"
    'redraw the text
    RichTextBox1.TextRTF = New_TextRTF
    RenderText
End Property

Public Sub LoadFile(sFilename As String)
    RichTextBox1.LoadFile sFilename, rtfRTF
    CopyBackGround
End Sub

Public Sub SaveFile(sFilename As String)
    RichTextBox1.SaveFile sFilename, rtfRTF
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_TextRTF = m_def_TextRTF
End Sub

Public Sub SetID(sTXT As String)
    SID = sTXT
End Sub
Public Function GetID() As String
    GetID = SID
End Function

Public Property Get AllowMove() As Boolean
    AllowMove = m_AllowMove
End Property

Public Property Let AllowMove(ByVal New_AllowMove As Boolean)
    PropertyChanged "AllowMove"
    m_AllowMove = New_AllowMove
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Picture1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Picture1.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    Picture1.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    Picture1.FontBold = PropBag.ReadProperty("FontBold", 0)
    Picture1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Picture1.FontName = PropBag.ReadProperty("FontName", Picture1.FontName)
    Picture1.FontSize = PropBag.ReadProperty("FontSize", Picture1.FontSize)
    Picture1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    Picture1.FontTransparent = PropBag.ReadProperty("FontTransparent", True)
    Picture1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    Picture1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set Picture1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_TextRTF = PropBag.ReadProperty("TextRTF", "")
    m_AllowMove = PropBag.ReadProperty("AllowMove", False)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderStyle", Picture1.BorderStyle, 1)
    Call PropBag.WriteProperty("CurrentX", Picture1.CurrentX, 0)
    Call PropBag.WriteProperty("CurrentY", Picture1.CurrentY, 0)
    Call PropBag.WriteProperty("FontBold", Picture1.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", Picture1.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", Picture1.FontName, "")
    Call PropBag.WriteProperty("FontSize", Picture1.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", Picture1.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontTransparent", Picture1.FontTransparent, True)
    Call PropBag.WriteProperty("FontUnderline", Picture1.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", Picture1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", Picture1.Font, Ambient.Font)
    Call PropBag.WriteProperty("TextRTF", RichTextBox1.TextRTF, "")
    Call PropBag.WriteProperty("AllowMove", m_AllowMove, False)
    
End Sub

