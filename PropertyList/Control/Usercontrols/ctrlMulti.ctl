VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "Lvbuttons.ocx"
Begin VB.UserControl ctrlMulti 
   BackColor       =   &H80000009&
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   ScaleHeight     =   18
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   204
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH""H""mm"""""""""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   7177
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
   Begin LVbuttons.LaVolpeButton cmdEllipse 
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "ctrlMulti.ctx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
End
Attribute VB_Name = "ctrlMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum ctrlType1
    eText = 1
    eEllipse = 2
    eReadOnly = 3
End Enum

Private pControlType As ctrlType1
Private sTagVariant As String

'Event Declarations:
Event EllipseClick()
Event Change()
Event ControlLostFocus()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get Text() As String
    Text = txtText.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtText.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get TagVariant() As String
    TagVariant = sTagVariant
End Property

Public Property Let TagVariant(ByVal vNewValue As String)
    sTagVariant = vNewValue
    PropertyChanged "TagVariant"
End Property

Public Property Get ControlType() As ctrlType1
    ControlType = pControlType
End Property

Public Property Let ControlType(ByVal vNewValue As ctrlType1)
    pControlType = vNewValue
    Call UserControl_Resize
    PropertyChanged "ControlType"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = txtText.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    txtText.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub txtText_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Get MultiLine() As Boolean
    MultiLine = txtText.MultiLine
End Property

Public Property Get PasswordChar() As String
    PasswordChar = txtText.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtText.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

Public Property Get ScrollBars() As Integer
    ScrollBars = txtText.ScrollBars
End Property

Public Function TextHeight(ByVal Str As String) As Single
End Function

Public Function TextWidth(ByVal Str As String) As Single
End Function

Public Property Get ToolTipText() As String
    ToolTipText = txtText.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    txtText.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Property Get SelLength() As Long
    SelLength = txtText.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtText.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

Public Property Get SelStart() As Long
    SelStart = txtText.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtText.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

Public Property Get SelText() As String
    SelText = txtText.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtText.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

Public Property Get Font() As Font
    Set Font = txtText.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtText.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ControlType = PropBag.ReadProperty("ControlType", eText)
    txtText.Text = PropBag.ReadProperty("Text", "")
    sTagVariant = PropBag.ReadProperty("TagVariant", "")
    txtText.Text = PropBag.ReadProperty("Text", "")
    txtText.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    txtText.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    txtText.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    txtText.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtText.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtText.SelText = PropBag.ReadProperty("SelText", "")
    Set txtText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    'If ControlType = eDate Then
    '    txtText.Text = Format(Now, "dd/mm/yyyy")
    'End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ControlType", pControlType, eText)
    Call PropBag.WriteProperty("Text", txtText.Text, "")
    Call PropBag.WriteProperty("TagVariant", sTagVariant, "")
    Call PropBag.WriteProperty("Text", txtText.Text, "")
    Call PropBag.WriteProperty("MousePointer", txtText.MousePointer, 0)
    Call PropBag.WriteProperty("PasswordChar", txtText.PasswordChar, "")
    Call PropBag.WriteProperty("ToolTipText", txtText.ToolTipText, "")
    Call PropBag.WriteProperty("SelLength", txtText.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtText.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtText.SelText, "")
    Call PropBag.WriteProperty("Font", txtText.Font, Ambient.Font)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Select Case ControlType
        Case eText
            txtText.Visible = True
            txtText.Enabled = True
            cmdEllipse.Visible = False
            txtText.Move 0, 0, UserControl.Extender.Width, UserControl.Extender.Height
        Case eEllipse
            txtText.Visible = True
            txtText.Enabled = False
            cmdEllipse.Visible = True
            cmdEllipse.Move UserControl.Extender.Width - UserControl.Extender.Height, 0, UserControl.Extender.Height, UserControl.Extender.Height
            txtText.Move 0, 0, UserControl.Extender.Width - UserControl.Extender.Height, UserControl.Extender.Height
        Case eReadOnly
            txtText.Visible = True
            txtText.Enabled = False
            cmdEllipse.Visible = False
            txtText.Move 0, 0, UserControl.Extender.Width, UserControl.Extender.Height
    End Select
End Sub

'Events ...
Private Sub txtText_LostFocus()
    RaiseEvent ControlLostFocus
End Sub

Private Sub txtText_Change()
    RaiseEvent Change
End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub txtText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub cmdEllipse_Click()
    RaiseEvent EllipseClick
End Sub


