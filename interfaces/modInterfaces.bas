Attribute VB_Name = "modInterfaces"
Option Explicit

' modInterfaces - create any interfaces in VB
'
' low level COM project by [rm] 2005


Private Declare Function IIDFromString Lib "ole32" ( _
    ByVal lpszIID As Long, _
    iid As Any _
) As Long

Private Declare Function CoCreateInstance Lib "ole32" ( _
    rclsid As Any, _
    ByVal pUnkOuter As Long, _
    ByVal dwClsContext As Long, _
    riid As Any, _
    ByVal ppv As Long _
) As Long

Private Declare Function CallWindowProcA Lib "user32" ( _
    ByVal addr As Long, _
    ByVal p1 As Long, _
    ByVal p2 As Long, _
    ByVal p3 As Long, _
    ByVal p4 As Long _
) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal dlen As Long _
)

Private Type GUID
    data1           As Long
    data2           As Integer
    data3           As Integer
    data4(7)        As Byte
End Type

Private Const CLSCTX_INPROC_SERVER As Long = 1&

Public Function CreateInterface(iface As interface) As Long
    Dim classid             As GUID
    Dim iid                 As GUID
    Dim obj                 As Long
    Dim hRes                As Long

    ' CLSID string to GUID struct
    If 0 <> IIDFromString(StrPtr(iface.CLSID), classid) Then
        Exit Function
    End If

    ' IID string to IID struct
    If 0 <> IIDFromString(StrPtr(iface.iid), iid) Then
        Exit Function
    End If

    ' create an instance of the requested interface
    hRes = CoCreateInstance(classid, 0, &H1, iid, VarPtr(obj))
    If hRes <> 0 Then
        Exit Function
    End If
    CreateInterface = hRes

    ' return the object pointer to our pseudo interface
    iface.object = obj
End Function

Public Function CallPointer(ByVal fnc As Long, ParamArray params()) As Long
    Dim btASM(&HEC00& - 1)  As Byte
    Dim pASM                As Long
    Dim i                   As Integer

    pASM = VarPtr(btASM(0))

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    For i = UBound(params) To 0 Step -1
        AddPush pASM, CLng(params(i))   ' PUSH dword
    Next

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallPointer = CallWindowProcA(VarPtr(btASM(0)), 0, 0, 0, 0)
End Function

Private Sub AddPush(pASM As Long, lng As Long)
    AddByte pASM, &H68
    AddLong pASM, lng
End Sub

Private Sub AddCall(pASM As Long, addr As Long)
    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong(pASM As Long, lng As Long)
    RtlMoveMemory ByVal pASM, lng, 4
    pASM = pASM + 4
End Sub

Private Sub AddByte(pASM As Long, bt As Byte)
    RtlMoveMemory ByVal pASM, bt, 1
    pASM = pASM + 1
End Sub
