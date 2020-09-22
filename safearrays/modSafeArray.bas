Attribute VB_Name = "modSafeArray"
Option Explicit

' modSafeArray - extend VB`s array features
'
' low level COM project by [rm] 2005


Public Type SAFEARRAYBOUND
    cElements   As Long             ' elements in this bound
    lLBound     As Long             ' lower bound
End Type

Public Type SAFEARRAY
    cDims       As Integer          ' Count of dimensions in this array
    fFeatures   As Integer          ' Flags used by the SafeArray
    cbElements  As Long             ' Size of an element of the array
    cLocks      As Long             ' Number of times the array has been locked without corresponding unlock
    pvData      As Long             ' Pointer to the data
    rgsabound() As SAFEARRAYBOUND   ' One bound for each dimension
End Type

Public Enum FeatureFlags
    FADF_AUTO = &H1                 ' An array that is allocated on the stack
    FADF_STATIC = &H2               ' An array that is statically allocated
    FADF_EMBEDDED = &H4             ' An array that is embedded in a structure
    FADF_FIXEDSIZE = &H10           ' An array that may not be resized or reallocated
    FADF_RECORD = &H20              ' An array containing records
    FADF_HAVEIID = &H40             ' An array that has an IID identifying interface
    FADF_HAVEVARTYPE = &H80         ' An array that has a VT type
    FADF_BSTR = &H100               ' An array of BSTRs
    FADF_UNKNOWN = &H200            ' An array of IUnknown*
    FADF_DISPATCH = &H400           ' An array of IDispatch*
    FADF_VARIANT = &H800            ' An array of VARIANTs
    FADF_RESERVED = &HF0E8          ' Bits reserved for future use
End Enum

Public Enum Varenum
    VT_EMPTY = 0&                   '
    VT_NULL = 1&                    ' 0
    VT_I2 = 2&                      ' signed 2 bytes integer
    VT_I4 = 3&                      ' signed 4 bytes integer
    VT_R4 = 4&                      ' 4 bytes float
    VT_R8 = 5&                      ' 8 bytes float
    VT_CY = 6&                      ' currency
    VT_DATE = 7&                    ' date
    VT_BSTR = 8&                    ' BStr
    VT_DISPATCH = 9&                ' IDispatch
    VT_ERROR = 10&                  ' error value
    VT_BOOL = 11&                   ' boolean
    VT_VARIANT = 12&                ' variant
    VT_UNKNOWN = 13&                ' IUnknown
    VT_DECIMAL = 14&                ' decimal
    VT_I1 = 16&                     ' signed byte
    VT_UI1 = 17&                    ' unsigned byte
    VT_UI2 = 18&                    ' unsigned 2 bytes integer
    VT_UI4 = 19&                    ' unsigned 4 bytes integer
    VT_I8 = 20&                     ' signed 8 bytes integer
    VT_UI8 = 21&                    ' unsigned 8 bytes integer
    VT_INT = 22&                    ' integer
    VT_UINT = 23&                   ' unsigned integer
    VT_VOID = 24&                   ' 0
    VT_HRESULT = 25&                ' HRESULT
    VT_PTR = 26&                    ' pointer
    VT_SAFEARRAY = 27&              ' safearray
    VT_CARRAY = 28&                 ' carray
    VT_USERDEFINED = 29&            ' userdefined
    VT_LPSTR = 30&                  ' LPStr
    VT_LPWSTR = 31&                 ' LPWStr
    VT_RECORD = 36&                 ' Record
    VT_FILETIME = 64&               ' File Time
    VT_BLOB = 65&                   ' Blob
    VT_STREAM = 66&                 ' Stream
    VT_STORAGE = 67&                ' Storage
    VT_STREAMED_OBJECT = 68&        ' Streamed Obj
    VT_STORED_OBJECT = 69&          ' Stored Obj
    VT_BLOB_OBJECT = 70&            ' Blob Obj
    VT_CF = 71&                     ' CF
    VT_CLSID = 72&                  ' Class ID
    VT_BSTR_BLOB = &HFFF&           ' BStr Blob
    VT_VECTOR = &H1000&             ' Vector
    VT_ARRAY = &H2000&              ' Array
    VT_BYREF = &H4000&              ' ByRef
    VT_RESERVED = &H8000&           ' Reserved
    VT_ILLEGAL = &HFFFF&            ' illegal
End Enum

Public Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    Optional ByVal dwLen As Long = 4 _
)

Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" ( _
    arr() As Any _
) As Long

Private Declare Function SafeArrayDestroyData Lib "oleaut32" ( _
    ByVal psa As Long _
) As Long

Private Declare Function SafeArrayDestroy Lib "oleaut32" ( _
    ByVal psa As Long _
) As Long

Public Sub DestroyArray( _
    ByVal ppsa As Long _
)

    Dim sa  As SAFEARRAY

    sa = GetSafeArray(ppsa)
    sa.cDims = 0
    sa.pvData = 0
    sa.cLocks = 0
    sa.fFeatures = 0
    sa.cbElements = 0
    SetSafeArray ppsa, sa
End Sub

Public Function ArrayFromPointer( _
    pData As Long, _
    elements As Long, _
    elementsize As Long, _
    vt As VbVarType, _
    Optional flags As FeatureFlags _
) As Variant

    Dim arr(0)  As Long
    Dim var     As Variant
    Dim sa      As SAFEARRAY
    Dim var2    As Variant

    var = arr

    ' get safearray of arr()
    sa = GetSafeArray(VarPtr(var) + 8)
    sa.pvData = pData                       ' new data ptr
    sa.cbElements = elementsize             ' size of 1 element
    sa.fFeatures = sa.fFeatures Or flags    ' new flags
    sa.rgsabound(0).cElements = elements    ' element count

    ' destroy data in arr()
    If 0 <> SafeArrayDestroyData(DeRefI4(VarPtr(var) + 8)) Then
        Debug.Print "ArrayFromPointer: Couldn' destroy array data"
    End If

    ' set new safearray for arr()
    SetSafeArray VarPtr(var) + 8, sa

    ' change the variant`s type
    CpyMem var, vt Or vbArray

    ' return the array
    ArrayFromPointer = var
End Function

' if an sa has a vartype defined, it is
' stored in the 4 bytes before the sa
Public Function SafeArrayVarType( _
    ByVal ppsa As Long _
) As Varenum

    Dim psa As Long

    If 0 = (GetSafeArray(ppsa).fFeatures And FADF_HAVEVARTYPE) Then
        Exit Function
    End If

    psa = DeRefI4(ppsa)
    CpyMem SafeArrayVarType, ByVal psa - 4
End Function

' overwrites a safearray
' YOU SHOULD NOT ADD NEW DIMENSIONS TO AN ARRAY THIS WAY!
Public Sub SetSafeArray( _
    ByVal ppsa As Long, _
    sa As SAFEARRAY _
)

    Dim psa As Long
    Dim cI  As Long

    psa = DeRefI4(ppsa)
    CpyMem ByVal psa, sa, 16

    For cI = 0 To sa.cDims - 1
        CpyMem ByVal psa + 16 + cI * 8, sa.rgsabound(cI), 8
    Next
End Sub

' returns the safearray struct of an array
Public Function GetSafeArray( _
    ByVal ppsa As Long _
) As SAFEARRAY

    Dim cDims   As Long
    Dim psa     As Long
    Dim cI      As Long

    psa = DeRefI4(ppsa)
    If psa = 0 Then Exit Function

    cDims = DeRefI2(psa)
    ReDim GetSafeArray.rgsabound(cDims - 1) As SAFEARRAYBOUND

    ' crashed for me when trying to copy the whole safearray at once.
    ' I think VB stores the sa of sa.rgsabound directly in the UDT.
    CpyMem GetSafeArray, ByVal psa, 16

    For cI = 0 To cDims - 1
        CpyMem GetSafeArray.rgsabound(cI), ByVal psa + 16 + cI * 8, 8
    Next
End Function

' dereferences 4-bytes-integer values
Private Function DeRefI4(ByVal ptr As Long) As Long
    If ptr = 0 Then Exit Function
    CpyMem ByVal VarPtr(DeRefI4), ByVal ptr
End Function

' dereferences 2-bytes-integer values
Private Function DeRefI2(ByVal ptr As Long) As Integer
    If ptr = 0 Then Exit Function
    CpyMem ByVal VarPtr(DeRefI2), ByVal ptr, 2
End Function
