VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' interface - create any interfaces in VB
'
' low level COM project by [rm] 2005

' IID
Public Property Get IID() As String: End Property
' CLSID
Public Property Get CLSID() As String: End Property
' object pointer
Public Property Get object() As Long: End Property
' object pointer
Public Property Let object(lng As Long): End Property
