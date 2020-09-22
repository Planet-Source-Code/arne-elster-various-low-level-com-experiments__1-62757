VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "low level COM examples"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "enum functions/properties of objects"
      Height          =   540
      Left            =   645
      TabIndex        =   1
      Top             =   1125
      Width           =   3390
   End
   Begin VB.CommandButton Command2 
      Caption         =   "fun with pointers"
      Height          =   540
      Left            =   645
      TabIndex        =   0
      Top             =   375
      Width           =   3390
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
    frmPtrFun.Show vbModal
End Sub

Private Sub Command3_Click()
    frmEnum.Show vbModal
End Sub
