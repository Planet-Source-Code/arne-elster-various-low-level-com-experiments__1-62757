VERSION 5.00
Begin VB.Form frmEnum 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Enum object members"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox lst 
      Height          =   2205
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   4665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "enum"
      Height          =   315
      Left            =   3900
      TabIndex        =   1
      Top             =   150
      Width           =   915
   End
   Begin VB.TextBox txtObj 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Text            =   "MSComctlLib.ImageComboCtl"
      Top             =   150
      Width           =   3690
   End
End
Attribute VB_Name = "frmEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim obj             As Object
    Dim Members()       As enmeinf

    Dim i               As Integer
    Dim strName         As String

    ' create the requested object
    Set obj = CreateObject(txtObj)

    ' get its members (properties, functions)
    Members = GetObjMembers(obj)

    lst.Clear
    For i = 0 To UBound(Members) - 1
        strName = Members(i).name
        Select Case Members(i).invkind
            Case INVOKE_FUNC:
                strName = "Function "
            Case INVOKE_PROPERTY_GET:
                strName = "Property Get "
            Case INVOKE_PROPERTY_PUT:
                strName = "Property Let "
            Case INVOKE_PROPERTY_PUTREF:
                strName = "Proeprty Set "
        End Select

        strName = strName & Members(i).name & " ("
        strName = strName & Members(i).params & " params)"

        lst.AddItem strName
    Next
End Sub
