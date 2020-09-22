Attribute VB_Name = "modCallPtr"
Option Explicit

Private Declare Function CallWindowProcA Lib "user32" ( _
            ByVal adr As Long, _
            ByVal p1 As Long, _
            ByVal p2 As Long, _
            ByVal p3 As Long, _
            ByVal p4 As Long) As Long

Private Declare Sub RtlFillMemory Lib "kernel32" ( _
            pDst As Any, _
            ByVal dlen As Long, _
            ByVal Fill As Byte)

Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
            pDst As Any, _
            pSrc As Any, _
            ByVal dlen As Long)

Private Const MAXCODE       As Long = &HEC00&

Public Function CallPointer(ByVal fnc As Long, ParamArray params()) As Long
    Dim btASM(MAXCODE - 1)  As Byte
    Dim pASM                As Long
    Dim i                   As Integer

    pASM = VarPtr(btASM(0))

    RtlFillMemory ByVal pASM, MAXCODE, &HCC

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    If UBound(params) = 0 Then
        If IsArray(params(0)) Then
            For i = UBound(params(0)) To 0 Step -1
                AddPush pASM, CLng(params(0)(i))    ' PUSH dword
            Next
        Else
            For i = UBound(params) To 0 Step -1
                AddPush pASM, CLng(params(i))       ' PUSH dword
            Next
        End If
    Else
        For i = UBound(params) To 0 Step -1
            AddPush pASM, CLng(params(i))           ' PUSH dword
        Next
    End If

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallPointer = CallWindowProcA(VarPtr(btASM(0)), _
                                  0, 0, 0, 0)
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
