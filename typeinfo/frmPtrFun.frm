VERSION 5.00
Begin VB.Form frmPtrFun 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "fun with pointers"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "call the function by its VTable entry (both EXE and IDE)"
      Height          =   615
      Left            =   270
      TabIndex        =   1
      Top             =   1050
      Width           =   4140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "call function by its name (only IDE)"
      Height          =   615
      Left            =   270
      TabIndex        =   0
      Top             =   300
      Width           =   4140
   End
End
Attribute VB_Name = "frmPtrFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cSay As clsSay

Public Sub Command1_Click()
    Dim pfnc    As Long
    Dim msg     As String

    msg = "hi"

    ' this isn't comparable to CallByName,
    ' as it get's the function's address
    ' from ITypeInfo
    ' CallByName uses IDispatch->Invoke
    ' which will also work compiled.
    ' The problem is, when compiled
    ' classes have no ITypeInfo,
    ' so we can't get the VTable offset
    ' of a function in a VB class.
    pfnc = GetFncInfo(cSay, "say").addr

    If pfnc = 0 Then
        MsgBox "Could not get address of ""say"".", vbExclamation
        Exit Sub
    End If

    CallPointer pfnc, ObjPtr(cSay), StrPtr(msg)
End Sub

Private Sub Command2_Click()
    Dim pfnc    As Long
    Dim msg     As String

    msg = "hi"

    ' get the value of the first (actually 8) entry
    pfnc = VTableEntry(cHi, 1)

    CallPointer pfnc, ObjPtr(cSay), StrPtr(msg)
End Sub

Private Sub Form_Load()
    Set cSay = New clsSay
End Sub
