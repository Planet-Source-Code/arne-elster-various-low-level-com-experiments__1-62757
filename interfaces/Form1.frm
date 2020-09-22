VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Active Desktop Wallpaper Changer"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   91
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4275
      Top             =   750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "BMP (*.bmp)|*.bmp|JPG (*.jpg;*.jpeg)|*.jpg;*.jpeg"
   End
   Begin VB.CommandButton cmdBrowse 
      Cancel          =   -1  'True
      Caption         =   "..."
      Height          =   315
      Left            =   4275
      TabIndex        =   2
      Top             =   300
      Width           =   465
   End
   Begin VB.TextBox txtWallpaper 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   300
      Width           =   4065
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "set new wallpaper"
      Height          =   390
      Left            =   900
      TabIndex        =   0
      Top             =   750
      Width           =   3315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    On Error Resume Next
    dlg.ShowOpen
    If Err Then Exit Sub
    txtWallpaper.Text = dlg.FileName
End Sub

Private Sub cmdSet_Click()
    Dim oDesktop        As IActiveDesktop
    Dim strWallpaper    As String

    ' this won't create a new IActiveDesktop
    ' like you may know it from other languages,
    ' this just creates an instance of our
    ' fake interface.
    Set oDesktop = New IActiveDesktop

    ' know we create the real IActiveDesktop instance
    If CreateInterface(oDesktop) <> 0 Then
        MsgBox "Couldn't create IActiveDesktop instance.", vbExclamation
        Exit Sub
    End If

    strWallpaper = txtWallpaper.Text

    ' OLE strings are unicode, perfect for us :)
    If oDesktop.SetWallpaper(StrPtr(strWallpaper), 0) <> 0 Then
        MsgBox "Couldn't set new wallpapre", vbExclamation
    Else
        ' now that we have set the new wallpaper, apply the changes
        ' and refresh the desktop
        If oDesktop.ApplyChanges(AD_APPLY_ALL Or AD_APPLY_FORCE) <> 0 Then
            MsgBox "Couldn't apply changes", vbExclamation
        End If
    End If

    ' last but not least we need to clean up
    oDesktop.Release
End Sub
