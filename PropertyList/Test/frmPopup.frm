VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmPopup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Popup"
   ClientHeight    =   2580
   ClientLeft      =   4995
   ClientTop       =   3405
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin LVbuttons.LaVolpeButton cmdOK 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&OK"
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
      MICON           =   "frmPopup.frx":0000
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
   Begin MSComCtl2.MonthView Main 
      Height          =   2310
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      StartOfWeek     =   19595265
      CurrentDate     =   37879
   End
   Begin LVbuttons.LaVolpeButton cmdCancel 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancel"
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
      MICON           =   "frmPopup.frx":001C
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
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bCancel As Boolean

Public Function GetDate(Optional PassedDate As String) As String
    If Len(PassedDate) > 0 Then
        Main.Value = Format(PassedDate, "yyyy/mm/dd")
    Else
        Main.Value = Now
    End If
    Me.Show vbModal
    If bCancel = True Then
        GetDate = ";"
    Else
        GetDate = Format(Main.Value, "yyyy/mm/dd") & ";"
    End If
    
    Unload Me

End Function


Private Sub cmdCancel_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    bCancel = False
    Me.Hide
End Sub

Private Sub Main_DblClick()
    Call cmdOK_Click
End Sub
