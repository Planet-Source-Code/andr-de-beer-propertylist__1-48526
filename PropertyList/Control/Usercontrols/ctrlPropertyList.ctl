VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ctrlPropertyList 
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
   Begin PropertyListBox.ctrlMulti txtTextField 
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Property"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctrlPropertyList.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctrlPropertyList.ctx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctrlPropertyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Event ButtonClick(ItemIndex As Long, PropertyDescription As String, PropertyValue As Variant, PropertyAdditionalValues As Variant, Cancel As Boolean)
Event ItemClick(Item As MSComctlLib.ListItem)

Private lngItemIndex  As Long
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVIR_BOUNDS = 0
Private Const LVIR_ICON = 1
Private Const LVIR_LABEL = 2
Private Const LVIR_SELECTBOUNDS = 3
Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Const LVSCW_AUTOSIZE As Long = -1
Const LVSCW_AUTOSIZE_USEHEADER As Long = -2 'Note: On last column, its width fills remaining width
                                            '   of list-view according to Micro$oft. This does not
                                            '   appear to be the case when I do it.
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private mCurrentItem As Long
Private mDirty As Boolean
Enum ctrlType
    eText = 1
    eEllipse = 2
    eReadOnly = 3
End Enum
Dim d As Dictionary

Private m_cFH As cFlatHeader

'==============================================================
Public Property Get ListItems() As MSComctlLib.ListItems
    Set ListItems = lvwMain.ListItems
End Property

'Properties ....
Public Property Get DescriptionColumnCaption() As String
    DescriptionColumnCaption = lvwMain.ColumnHeaders(1).Text
End Property
Public Property Let DescriptionColumnCaption(ByVal vNewValue As String)
    lvwMain.ColumnHeaders(1).Text = vNewValue
    PropertyChanged "DescriptionColumnCaption"
End Property
Public Property Get ValueColumnCaption() As String
    ValueColumnCaption = lvwMain.ColumnHeaders(2).Text
End Property
Public Property Let ValueColumnCaption(ByVal vNewValue As String)
    lvwMain.ColumnHeaders(2).Text = vNewValue
    PropertyChanged "ValueColumnCaption"
End Property
Public Property Get HideColumnHeaders() As Boolean
    HideColumnHeaders = lvwMain.HideColumnHeaders
End Property
Public Property Let HideColumnHeaders(ByVal vNewValue As Boolean)
    lvwMain.HideColumnHeaders = vNewValue
    PropertyChanged "HideColumnHeaders"
End Property
'==============================================================

Private Sub txtTextField_EllipseClick()
    Dim it As clsItem
    'Set it = d.Item(pIndex)
    Dim bcancel As Boolean
    Dim sRetDescription As String
    Dim sRetValue As String
    Dim sRetAddValues As String
    Dim lstitem As ListItem
    Set lstitem = lvwMain.ListItems(lngItemIndex)
    sRetDescription = lstitem.Text
    sRetValue = lstitem.SubItems(1)
    sRetAddValues = lstitem.Tag
    
    RaiseEvent ButtonClick(lngItemIndex, sRetDescription, sRetValue, sRetAddValues, bcancel)
    If bcancel = True Then
        Exit Sub
    Else
        txtTextField.Text = sRetValue
        lstitem.SubItems(1) = sRetValue
        lstitem.Tag = sRetAddValues
    End If
End Sub
Private Sub UserControl_Initialize()
    Set d = New Dictionary
    Set txtTextField.Font = lvwMain.Font
    Set m_cFH = New cFlatHeader
    m_cFH.Attach lvwMain.hwnd
    m_cFH.LVFlatScrollBars(lvwMain.hwnd) = True
End Sub

'==============================================================
'Property Bag ....
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lvwMain.ColumnHeaders(1).Text = PropBag.ReadProperty("DescriptionColumnCaption", "Property")
    lvwMain.ColumnHeaders(2).Text = PropBag.ReadProperty("ValueColumnCaption", "Value")
    lvwMain.HideColumnHeaders = PropBag.ReadProperty("HideColumnHeaders", False)
End Sub

Private Sub UserControl_Terminate()
    m_cFH.Detach
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DescriptionColumnCaption", lvwMain.ColumnHeaders(1).Text, "Property")
    Call PropBag.WriteProperty("ValueColumnCaption", lvwMain.ColumnHeaders(2).Text, "Value")
    Call PropBag.WriteProperty("HideColumnHeaders", lvwMain.HideColumnHeaders, False)
End Sub
'==============================================================

'==============================================================
'Events ....
Private Sub UserControl_Resize()
    lvwMain.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    If UserControl.Ambient.UserMode = True Then
        AutoSize 0
    End If
End Sub
Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lstitem As ListItem
    
    If mCurrentItem = 0 Then mCurrentItem = 1: lvwMain.ListItems(1).Selected = True: Call lvwMain_ItemClick(lvwMain.ListItems(1)): Exit Sub
    
    txtTextField.Visible = True
    txtTextField.ZOrder 0
    If mDirty Then
        lvwMain.ListItems(mCurrentItem).SubItems(1) = txtTextField.Text
    End If
    mCurrentItem = Item.Index
    Display mCurrentItem, 1
    RaiseEvent ItemClick(Item)
    
        
    For Each lstitem In lvwMain.ListItems
        If lstitem.Selected = True Then
            lstitem.SmallIcon = 1
        Else
            lstitem.SmallIcon = 2
        End If
    Next lstitem
    
    'Display mCurrentItem, 1
    
End Sub
Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    Dim it As New clsItem
    If mCurrentItem = 0 Then Exit Sub
    Set it = d.Item(mCurrentItem)
    If it.ItemType = 3 Or it.ItemType = 2 Then Exit Sub
    txtTextField.Text = Chr(KeyAscii)
    txtTextField.SelStart = 1
    If txtTextField.Visible Then
        txtTextField.SetFocus
    Else
        txtTextField.Visible = True
        txtTextField.SetFocus
    End If
End Sub
Private Sub txtTextField_Change()
    mDirty = True
End Sub
Private Sub txtTextField_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lvwMain.ListItems.Item(mCurrentItem).SubItems(1) = txtTextField.Text
        lvwMain.SetFocus
    End If
End Sub
Private Sub txtTextField_ControlLostFocus()
    On Error Resume Next
    lvwMain.ListItems.Item(mCurrentItem).SubItems(1) = txtTextField.Text
    lvwMain.SetFocus
End Sub
'==============================================================

Public Sub AddProperty(pText As String, pValue As String, Optional pAddValues As Variant = "", Optional pType As ctrlType = eText) ', ListValues As Variant
    Dim pItem As ListItem
    Dim pTestWidth As Long
    'Calculate optimal width
    pTestWidth = getTextWidth(pText) + 30
    If pTestWidth > lvwMain.ColumnHeaders(1).Width Then
        lvwMain.ColumnHeaders(1).Width = pTestWidth + 30
    End If
    'Calculate optimal width
    pTestWidth = getTextWidth(pValue) + 40
    If pTestWidth > lvwMain.ColumnHeaders(2).Width Then
       lvwMain.ColumnHeaders(2).Width = pTestWidth + 40
    End If
    Set pItem = lvwMain.ListItems.Add(, , pText, 0, 2)
        pItem.SubItems(1) = pValue
        pItem.Tag = pAddValues
    'Add to the dictionary ...
    Dim it As New clsItem
    With it
        .ItemIndex = pItem.Index
        .ItemType = pType
    End With
    d.Add pItem.Index, it
    'Set to nothing ...
    Set pItem = Nothing
    AutoSize 0
End Sub

Private Function getTextWidth(pString As String) As Long
    getTextWidth = UserControl.TextWidth(pString)
End Function

Sub Display(pIndex As Long, pSubitem As Long)
    Const lvBorder = 3
    Const lvGrid = 1
    Dim pMode As String
    Dim pFont As New StdFont
    Dim pRect As RECT
    Dim APIItemIndex As Long
    Dim a
    APIItemIndex = pIndex - 1
    pRect.Top = pSubitem
    pRect.Left = LVIR_LABEL
    a = SendMessage(lvwMain.hwnd, LVM_GETSUBITEMRECT, APIItemIndex, pRect)
    lngItemIndex = pIndex
    Dim it As clsItem
    Set it = d.Item(pIndex)
    With txtTextField
        .Visible = True
        .ZOrder 0
        .Top = lvwMain.Top + pRect.Top '+ lvBorder + lvGrid
        .Left = lvwMain.Left + pRect.Left ' + lvBorder + lvGrid
        .Height = (pRect.Bottom) - (pRect.Top) '- lvGrid
        .Width = (pRect.Right - pRect.Left) '- lvGrid
        .Text = lvwMain.ListItems(pIndex).SubItems(pSubitem)
        .ControlType = it.ItemType
    End With
End Sub

Private Sub AutoSize(Index As Integer)
    LockWindowUpdate lvwMain.hwnd               ' Lock update of ListView. Prevents ghostly text
                                                ' from appearing. I have seen it happen in other
                                                ' projects, but not this one. Always a good idea
                                                ' to use nonetheless.
    SendMessage lvwMain.hwnd, LVM_SETCOLUMNWIDTH, Index, LVSCW_AUTOSIZE_USEHEADER ' The magic of auotosize
    LockWindowUpdate 0                          ' Unlock
    
    lvwMain.ColumnHeaders(2).Width = (lvwMain.Width - lvwMain.ColumnHeaders(1).Width)
    
End Sub
