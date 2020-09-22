VERSION 5.00
Object = "{01CBB486-93F2-4FF0-9E8F-DA9077A16A32}#1.0#0"; "PropertyListBox.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   3045
   ClientTop       =   2445
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   6585
   Begin PropertyListBox.ctrlPropertyList ctrlPropertyList1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5318
   End
   Begin VB.Menu mnuEnum 
      Caption         =   "&Enumerate Values"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ctrlPropertyList1_ButtonClick(ItemIndex As Long, PropertyDescription As String, PropertyValue As Variant, PropertyAdditionalValues As Variant, Cancel As Boolean)
    Dim splits() As String
    Select Case PropertyDescription
        Case "Date"
            splits = Split(frmPopup.GetDate(CStr(PropertyValue)), ";", -1, vbTextCompare)
            If splits(0) = "" Then
                Cancel = True
                Exit Sub
            Else
                PropertyValue = splits(0)
                PropertyAdditionalValues = ""
                Cancel = False
            End If
    End Select
End Sub

Private Sub ctrlPropertyList1_ItemClick(Item As MSComctlLib.ListItem)
'    MsgBox Item.Text & ":" & Item.SubItems(1)
End Sub

Private Sub Form_Load()
    ctrlPropertyList1.AddProperty "First Name", "Peter", "", eText
    ctrlPropertyList1.AddProperty "Middle Name", "Jones", "", eText
    ctrlPropertyList1.AddProperty "Last Name", "Pan", "", eText
    ctrlPropertyList1.AddProperty "Date", Format(Now, "yyyy/mm/dd"), "", eEllipse
End Sub

Private Sub Form_Resize()
    ctrlPropertyList1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuEnum_Click()
    Dim lstitem As MSComctlLib.ListItem
    Dim strVal As String
    
    For Each lstitem In ctrlPropertyList1.ListItems
        strVal = strVal & lstitem.Text & ":" & lstitem.SubItems(1) & vbCrLf
    Next lstitem

    MsgBox strVal

End Sub
