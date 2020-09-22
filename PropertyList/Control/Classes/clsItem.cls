VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lngItemIndex As Long
Private lngItemType As Long

Public Property Get ItemIndex() As Long
    ItemIndex = lngItemIndex
End Property

Public Property Let ItemIndex(ByVal vNewValue As Long)
    lngItemIndex = vNewValue
End Property

Public Property Get ItemType() As Long
    ItemType = lngItemType
End Property

Public Property Let ItemType(ByVal vNewValue As Long)
    lngItemType = vNewValue
End Property

