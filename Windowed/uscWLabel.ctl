VERSION 5.00
Begin VB.UserControl WLabel 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4128
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   HasDC           =   0   'False
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "uscWLabel.ctx":0000
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2856
      Left            =   0
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   0
      Top             =   180
      Width           =   2244
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   804
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2712
      End
   End
End
Attribute VB_Name = "WLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum wlStandardBorderStyleConstants
    wlNone = 0
    wlFixedSingle = 1
End Enum

Public Enum wlStandardAppearanceConstants
    wlAppearanceFlat = 0
    wlAppearance3D = 1
End Enum

Public Enum wlTextOrientationConstants
    wlHorizontal = 0
    wlVertical = 1
    wlReverse = 2
    wlVerticalFromTop = 3
End Enum

Private Type XFORM
  eM11 As Single
  eM12 As Single
  eM21 As Single
  eM22 As Single
  eDx As Single
  eDy As Single
End Type

Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Private Declare Function SetWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As XFORM) As Long
Private Declare Function ModifyWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As XFORM, ByVal iMode As Long) As Long
Private Const MWT_IDENTITY = 1
Private Const MWT_LEFTMULTIPLY = 2
Private Const MWT_RIGHTMULTIPLY = 3

Private Const GM_ADVANCED = 2
Private Const GM_COMPATIBLE = 1

'Default Property Values:
Private Const mdef_BackColor = vbWindowBackground
Private Const mdef_ForeColor = vbWindowText
Private Const mdef_Enabled = True
Private Const mdef_Appearance = 1
Private Const mdef_BorderStyle = 0
Private Const mdef_Caption = ""
Private Const mdef_Alignment = vbLeftJustify
Private Const mdef_AutoSize = False
Private Const mdef_FontUnderline = False
Private Const mdef_FontStrikethru = False
Private Const mdef_FontSize = 0
Private Const mdef_FontName = ""
Private Const mdef_FontItalic = False
Private Const mdef_FontBold = False
Private Const mdef_WordWrap = False
Private Const mdef_Orientation = wlHorizontal ' wlVertical

'Property Variables:
Private mBackColor As Long
Private mForeColor As Long
Private mEnabled As Boolean
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private mAppearance As Integer
Private mBorderStyle As Integer
Private mCaption As String
Private mAlignment As Integer
Private mAutoSize As Boolean
Private mFontUnderline As Boolean
Private mFontStrikethru As Boolean
Private mFontSize As Single
Private mFontName As String
Private mFontItalic As Boolean
Private mFontBold As Boolean
Private mWordWrap As Boolean
Private mOrientation As wlTextOrientationConstants
Private mGMPrev As Long

'Event Declarations:
Public Event Click()
Public Event Change()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


Private Sub mFont_FontChanged(ByVal PropertyName As String)
    If PropertyName = "Name" Then
        mFontName = mFont.Name
    ElseIf PropertyName = "Size" Then
        mFontSize = mFont.Size
    ElseIf PropertyName = "Bold" Then
        mFontBold = mFont.Bold
    ElseIf PropertyName = "Italic" Then
        mFontItalic = mFont.Italic
    ElseIf PropertyName = "Strikethrough" Then
        mFontStrikethru = mFont.Strikethrough
    ElseIf PropertyName = "Underline" Then
        mFontUnderline = mFont.Underline
    End If
    PropertyChanged "Font"
End Sub

Private Sub Picture1_Click()
    RaiseEvent Click
End Sub

Private Sub Picture1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Picture1_Paint()
    Dim mtx1 As XFORM, mtx2 As XFORM, c As Single, s As Single, p As IPicture
    
    ModifyWorldTransform Picture1.hdc, mtx1, MWT_IDENTITY
    If mOrientation = wlVertical Then
        c = 0
        s = -1
        mtx1.eM11 = c: mtx1.eM12 = s: mtx1.eM21 = -s: mtx1.eM22 = c: mtx1.eDx = 0: mtx1.eDy = Picture1.ScaleHeight
        mtx2.eM11 = 1: mtx2.eM22 = 1: mtx2.eDx = -0: mtx2.eDy = -Picture1.ScaleHeight
    ElseIf mOrientation = wlVerticalFromTop Then
        c = 0
        s = 1
        mtx1.eM11 = c: mtx1.eM12 = s: mtx1.eM21 = -s: mtx1.eM22 = c: mtx1.eDx = Picture1.ScaleWidth: mtx1.eDy = 0
        mtx2.eM11 = 1: mtx2.eM22 = 1: mtx2.eDx = -Picture1.ScaleWidth: mtx2.eDy = -0
    ElseIf mOrientation = wlHorizontal Then
        c = 1
        s = 0
        mtx1.eM11 = c: mtx1.eM12 = s: mtx1.eM21 = -s: mtx1.eM22 = c: mtx1.eDx = Picture1.ScaleWidth / 2: mtx1.eDy = Picture1.ScaleHeight / 2
        mtx2.eM11 = 1: mtx2.eM22 = 1: mtx2.eDx = -Picture1.ScaleWidth / 2: mtx2.eDy = -Picture1.ScaleHeight / 2
    Else
        c = -1
        s = 0.0001
        mtx1.eM11 = c: mtx1.eM12 = s: mtx1.eM21 = -s: mtx1.eM22 = c: mtx1.eDx = Picture1.ScaleWidth / 2: mtx1.eDy = Picture1.ScaleHeight / 2
        mtx2.eM11 = 1: mtx2.eM22 = 1: mtx2.eDx = -Picture1.ScaleWidth / 2: mtx2.eDy = -Picture1.ScaleHeight / 2
    End If
    SetWorldTransform Picture1.hdc, mtx1
    ModifyWorldTransform Picture1.hdc, mtx2, MWT_LEFTMULTIPLY
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mGMPrev = SetGraphicsMode(Picture1.hdc, GM_ADVANCED)
    
    mBackColor = PropBag.ReadProperty("BackColor", mdef_BackColor)
    mForeColor = PropBag.ReadProperty("ForeColor", mdef_ForeColor)
    mEnabled = PropBag.ReadProperty("Enabled", mdef_Enabled)
    Set mFont = PropBag.ReadProperty("Font", Ambient.Font)
    mAppearance = PropBag.ReadProperty("Appearance", mdef_Appearance)
    mBorderStyle = PropBag.ReadProperty("BorderStyle", mdef_BorderStyle)
    mCaption = PropBag.ReadProperty("Caption", mdef_Caption)
    mAlignment = PropBag.ReadProperty("Alignment", mdef_Alignment)
    mAutoSize = PropBag.ReadProperty("AutoSize", mdef_AutoSize)
    mFontUnderline = PropBag.ReadProperty("FontUnderline", mdef_FontUnderline)
    mFontStrikethru = PropBag.ReadProperty("FontStrikethru", mdef_FontStrikethru)
    mFontSize = PropBag.ReadProperty("FontSize", mdef_FontSize)
    mFontName = PropBag.ReadProperty("FontName", mdef_FontName)
    mFontItalic = PropBag.ReadProperty("FontItalic", mdef_FontItalic)
    mFontBold = PropBag.ReadProperty("FontBold", mdef_FontBold)
    mWordWrap = PropBag.ReadProperty("WordWrap", mdef_WordWrap)
    mOrientation = PropBag.ReadProperty("Orientation", mdef_Orientation)
    
    SetLabel
End Sub

Private Sub UserControl_Resize()
    Picture1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    If mOrientation = wlVertical Then
        Label1.Move 0, Picture1.ScaleHeight, Picture1.ScaleHeight, Picture1.ScaleWidth
    ElseIf mOrientation = wlVerticalFromTop Then
        Label1.Move Picture1.ScaleWidth, 0, Picture1.ScaleHeight, Picture1.ScaleWidth
    ElseIf mOrientation = wlHorizontal Then
        Label1.Move 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight - 1
    Else
        Label1.Move 0, 0, Picture1.ScaleWidth - 1, Picture1.ScaleHeight
    End If
    Picture1.Cls
    UserControl.Cls
    Picture1.Refresh
    UserControl.Refresh
    Label1.Refresh
End Sub

'MemberInfo=8,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gr�ficos en un objeto."
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal nBackColor As OLE_COLOR)
    If nBackColor <> mBackColor Then
        mBackColor = nBackColor
        PropertyChanged "BackColor"
        Label1.BackColor = mBackColor
        Picture1.Refresh
    End If
End Property


'MemberInfo=8,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gr�ficos en un objeto."
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal nForeColor As OLE_COLOR)
    If nForeColor <> mForeColor Then
        mForeColor = nForeColor
        PropertyChanged "ForeColor"
        Label1.ForeColor = mForeColor
        Picture1.Refresh
    End If
End Property


'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal nEnabled As Boolean)
    If nEnabled <> mEnabled Then
        mEnabled = nEnabled
        PropertyChanged "Enabled"
        Label1.Enabled = mEnabled
        UserControl.Enabled = mEnabled
        Picture1.Refresh
    End If
End Property


'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = mFont
End Property

Public Property Set Font(ByVal nFont As Font)
    If Not nFont Is mFont Then
        Set mFont = nFont
        PropertyChanged "Font"
        Set Label1.Font = mFont
        Picture1.Refresh
    End If
End Property

Public Property Let Font(ByVal nFont As Font)
    Set Font = nFont
End Property


'MemberInfo=7,0,0,0
Public Property Get Appearance() As wlStandardAppearanceConstants
Attribute Appearance.VB_Description = "Devuelve o establece si los objetos se dibujan en tiempo de ejecuci�n con efectos 3D."
    Appearance = mAppearance
End Property

Public Property Let Appearance(ByVal nAppearance As wlStandardAppearanceConstants)
    If nAppearance <> mAppearance Then
        If nAppearance < wlAppearanceFlat Then nAppearance = wlAppearanceFlat
        If nAppearance > wlAppearance3D Then nAppearance = wlAppearance3D
        mAppearance = nAppearance
        PropertyChanged "Appearance"
        Label1.Appearance = mAppearance
        mBackColor = Label1.BackColor
        mForeColor = Label1.ForeColor
        Picture1.Refresh
    End If
End Property


'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As wlStandardBorderStyleConstants
Attribute BorderStyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(ByVal nBorderStyle As wlStandardBorderStyleConstants)
    If nBorderStyle <> mBorderStyle Then
        If nBorderStyle < wlNone Then nBorderStyle = wlNone
        If nBorderStyle > wlFixedSingle Then nBorderStyle = wlFixedSingle
        mBorderStyle = nBorderStyle
        PropertyChanged "BorderStyle"
        Label1.BorderStyle = mBorderStyle
        Picture1.Refresh
    End If
End Property


'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Obliga a volver a dibujar un objeto."
     Label1.Refresh
End Sub


'MemberInfo=13,0,0,
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Devuelve o establece el texto mostrado en la barra de t�tulo de un objeto o bajo el icono de un objeto."
    Caption = mCaption
End Property

Public Property Let Caption(ByVal nCaption As String)
    If nCaption <> mCaption Then
        mCaption = nCaption
        PropertyChanged "Caption"
        Label1.Caption = mCaption
        SetAutoSize
        Picture1.Refresh
        RaiseEvent Change
    End If
End Property


'MemberInfo=7,0,0,0
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Devuelve o establece la alineaci�n de un control CheckBox u OptionButton, o el texto de un control."
    Alignment = mAlignment
End Property

Public Property Let Alignment(ByVal nAlignment As AlignmentConstants)
    If nAlignment <> mAlignment Then
        If (nAlignment < vbLeftJustify) Then nAlignment = vbLeftJustify
        If (nAlignment > vbCenter) Then nAlignment = vbCenter
        mAlignment = nAlignment
        PropertyChanged "Alignment"
        Label1.Alignment = mAlignment
        Picture1.Refresh
    End If
End Property


'MemberInfo=0,0,0,0
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determina si un control cambia de tama�o autom�ticamente para mostrar todo su contenido."
    AutoSize = mAutoSize
End Property

Public Property Let AutoSize(ByVal nAutoSize As Boolean)
    
    If nAutoSize <> mAutoSize Then
        mAutoSize = nAutoSize
        PropertyChanged "AutoSize"
        Label1.AutoSize = mAutoSize
        Label1.Refresh
        SetAutoSize
        Picture1.Refresh
    End If
End Property


'MemberInfo=0,0,0,0
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Devuelve o establece el estilo subrayado de una fuente."
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = mFontUnderline
End Property

Public Property Let FontUnderline(ByVal nFontUnderline As Boolean)
    If nFontUnderline <> mFontUnderline Then
        mFontUnderline = nFontUnderline
        PropertyChanged "FontUnderline"
        Label1.FontUnderline = mFontUnderline
        Picture1.Refresh
    End If
End Property


'MemberInfo=0,0,0,0
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Devuelve o establece el estilo tachado de una fuente."
Attribute FontStrikethru.VB_MemberFlags = "400"
    FontStrikethru = mFontStrikethru
End Property

Public Property Let FontStrikethru(ByVal nFontStrikethru As Boolean)
    If nFontStrikethru <> mFontStrikethru Then
        mFontStrikethru = nFontStrikethru
        PropertyChanged "FontStrikethru"
        Label1.FontStrikethru = mFontStrikethru
        Picture1.Refresh
    End If
End Property


'MemberInfo=12,0,0,0
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Especifica el tama�o (en puntos) de la fuente que aparece en cada fila del nivel especificado."
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = mFontSize
End Property

Public Property Let FontSize(ByVal nFontSize As Single)
    If nFontSize <> mFontSize Then
        mFontSize = nFontSize
        PropertyChanged "FontSize"
        Label1.FontSize = mFontSize
        Picture1.Refresh
    End If
End Property


'MemberInfo=13,0,0,
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Especifica el nombre de la fuente que aparece en cada fila del nivel especificado."
Attribute FontName.VB_MemberFlags = "400"
    FontName = mFontName
End Property

Public Property Let FontName(ByVal nFontName As String)
    If nFontName <> mFontName Then
        mFontName = nFontName
        PropertyChanged "FontName"
        Label1.FontName = mFontName
        Picture1.Refresh
    End If
End Property


'MemberInfo=0,0,0,0
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Devuelve o establece el estilo cursiva de una fuente."
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = mFontItalic
End Property

Public Property Let FontItalic(ByVal nFontItalic As Boolean)
    If nFontItalic <> mFontItalic Then
        mFontItalic = nFontItalic
        PropertyChanged "FontItalic"
        Label1.FontItalic = mFontItalic
        Picture1.Refresh
    End If
End Property


'MemberInfo=0,0,0,0
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Devuelve o establece el estilo negrita de una fuente."
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = mFontBold
End Property

Public Property Let FontBold(ByVal nFontBold As Boolean)
    If nFontBold <> mFontBold Then
        mFontBold = nFontBold
        PropertyChanged "FontBold"
        Label1.FontBold = mFontBold
        Picture1.Refresh
    End If
End Property


'MemberInfo=0,0,0,0
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Devuelve o establece un valor que determina si un control se expande para ajustarse al texto de su t�tulo."
    WordWrap = mWordWrap
End Property

Public Property Let WordWrap(ByVal nWordWrap As Boolean)
    If nWordWrap <> mWordWrap Then
        mWordWrap = nWordWrap
        PropertyChanged "WordWrap"
        Label1.WordWrap = mWordWrap
        Picture1.Refresh
    End If
End Property


Public Property Get Orientation() As wlTextOrientationConstants
    Orientation = mOrientation
End Property

Public Property Let Orientation(nOrientation As wlTextOrientationConstants)
    If nOrientation <> mOrientation Then
        If nOrientation < wlHorizontal Then nOrientation = wlHorizontal
        If nOrientation > wlVerticalFromTop Then nOrientation = wlVerticalFromTop
        mOrientation = nOrientation
        PropertyChanged "Orientaion"
        SetAutoSize
        UserControl_Resize
    End If
End Property

Private Sub UserControl_InitProperties()
    mGMPrev = SetGraphicsMode(Picture1.hdc, GM_ADVANCED)
    
    mBackColor = mdef_BackColor
    mForeColor = mdef_ForeColor
    mEnabled = mdef_Enabled
    Set mFont = Ambient.Font
    mAppearance = mdef_Appearance
    mBorderStyle = mdef_BorderStyle
    mCaption = Ambient.DisplayName
    mAlignment = mdef_Alignment
    mAutoSize = mdef_AutoSize
    mFontUnderline = mdef_FontUnderline
    mFontStrikethru = mdef_FontStrikethru
    mFontSize = mdef_FontSize
    mFontName = mdef_FontName
    mFontItalic = mdef_FontItalic
    mFontBold = mdef_FontBold
    mWordWrap = mdef_WordWrap
    mOrientation = mdef_Orientation
    
    SetLabel
End Sub

Private Sub UserControl_Terminate()
    SetGraphicsMode Picture1.hdc, mGMPrev
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", mBackColor, mdef_BackColor)
    Call PropBag.WriteProperty("ForeColor", mForeColor, mdef_ForeColor)
    Call PropBag.WriteProperty("Enabled", mEnabled, mdef_Enabled)
    Call PropBag.WriteProperty("Font", mFont, Ambient.Font)
    Call PropBag.WriteProperty("Appearance", mAppearance, mdef_Appearance)
    Call PropBag.WriteProperty("BorderStyle", mBorderStyle, mdef_BorderStyle)
    Call PropBag.WriteProperty("Caption", mCaption, mdef_Caption)
    Call PropBag.WriteProperty("Alignment", mAlignment, mdef_Alignment)
    Call PropBag.WriteProperty("AutoSize", mAutoSize, mdef_AutoSize)
    Call PropBag.WriteProperty("FontUnderline", mFontUnderline, mdef_FontUnderline)
    Call PropBag.WriteProperty("FontStrikethru", mFontStrikethru, mdef_FontStrikethru)
    Call PropBag.WriteProperty("FontSize", mFontSize, mdef_FontSize)
    Call PropBag.WriteProperty("FontName", mFontName, mdef_FontName)
    Call PropBag.WriteProperty("FontItalic", mFontItalic, mdef_FontItalic)
    Call PropBag.WriteProperty("FontBold", mFontBold, mdef_FontBold)
    Call PropBag.WriteProperty("WordWrap", mWordWrap, mdef_WordWrap)
    Call PropBag.WriteProperty("Orientation", mOrientation, mdef_Orientation)
End Sub

Private Sub SetLabel()
    Label1.Appearance = mAppearance
    Label1.BackColor = mBackColor
    Label1.ForeColor = mForeColor
    Label1.Enabled = mEnabled
    UserControl.Enabled = mEnabled
    Set Label1.Font = mFont
    Label1.BorderStyle = mBorderStyle
    Label1.Caption = mCaption
    Label1.Alignment = mAlignment
    Label1.AutoSize = mAutoSize
    Label1.FontUnderline = mFontUnderline
    Label1.FontStrikethru = mFontStrikethru
    If mFontSize > 0 Then Label1.FontSize = mFontSize
    If mFontName <> "" Then Label1.FontName = mFontName
    Label1.FontItalic = mFontItalic
    Label1.FontBold = mFontBold
    Label1.WordWrap = mWordWrap
    SetAutoSize
    UserControl_Resize
End Sub

Private Sub SetAutoSize()
    If mAutoSize Then
        If (mOrientation = wlVertical) Or (mOrientation = wlVerticalFromTop) Then
            UserControl.Size ScaleX(Label1.Height, vbPixels, vbTwips), ScaleY(Label1.Width + 1, vbPixels, vbTwips)
        Else
            UserControl.Size ScaleY(Label1.Width + 1, vbPixels, vbTwips), ScaleX(Label1.Height, vbPixels, vbTwips)
        End If
    End If
End Sub

Public Property Get hWnd() As Long
    hWnd = Picture1.hWnd
End Property

Public Property Get UserControlhWnd() As Long
    UserControlhWnd = UserControl.hWnd
End Property

