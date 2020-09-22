VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IActiveDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' pseudo interface
Implements interface

' CLSID of this interface
Private Const CLSID As String = "{75048700-EF1F-11D0-9888-006097DEACF9}"
Private Const IID As String = "{F490EB00-1240-11D1-9888-006097DEACF9}"

Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
    pDst As Any, pSrc As Any, ByVal dlen As Long)

Public Enum AD_APPLY
    AD_APPLY_SAVE = &H1
    AD_APPLY_HTMLGEN = &H2
    AD_APPLY_REFRESH = &H4
    AD_APPLY_ALL = &H7
    AD_APPLY_FORCE = &H8
    AD_APPLY_BUFFERED_REFRESH = &H10
    AD_APPLY_DYNAMICREFRESH = &H20
End Enum

Public Enum COMP_ELEM
    COMP_ELEM_TYPE = &H1
    COMP_ELEM_CHECKED = &H2
    COMP_ELEM_DIRTY = &H4
    COMP_ELEM_NOSCROLL = &H8
    COMP_ELEM_POS_LEFT = &H10
    COMP_ELEM_POS_TOP = &H20
    COMP_ELEM_SIZE_WIDTH = &H40
    COMP_ELEM_SIZE_HEIGHT = &H80
    COMP_ELEM_POS_ZINDEX = &H100
    COMP_ELEM_SOURCE = &H200
    COMP_ELEM_FRIENDLYNAME = &H400
    COMP_ELEM_SUBSCRIBEDURL = &H800
    COMP_ELEM_ORIGINAL_CSI = &H1000
    COMP_ELEM_RESTORED_CSI = &H2000
    COMP_ELEM_CURITEMSTATE = &H4000
    COMP_ELEM_ALL = &H7FFF
End Enum

Public Enum DTI_ADTIWUI
    DTI_ADDUI_DEFAULT = &H0
    DTI_ADDUI_DISPSUBWIZARD = &H1
    DTI_ADDUI_POSITIONITEM = &H2
End Enum

Public Enum AddUrl
    ADDURL_SILENT = &H1
End Enum

' VTable of IActiveDesktop : IUnknown
Private Type ActiveDesktop
    ' IUnknown
    QueryInterface          As Long
    AddRef                  As Long
    Release                 As Long
    ' IActiveDesktop
    ApplyChanges            As Long
    GetWallpaper            As Long
    SetWallpaper            As Long
    GetWallpaperOptions     As Long
    SetWallpaperOptions     As Long
    GetPattern              As Long
    SetPattern              As Long
    GetDesktopItemOptions   As Long
    SetDesktopItemOptions   As Long
    AddDesktopItem          As Long
    AddDesktopItemWithUI    As Long
    ModifyDesktopItem       As Long
    RemoveDesktopItem       As Long
    GetDesktopItemCount     As Long
    GetDesktopItem          As Long
    GetDesktopItemByID      As Long
    GenerateDesktopItemHtml As Long
    AddUrl                  As Long
    GetDesktopItemBySource  As Long
End Type
Private vtbl                As ActiveDesktop

Private obj                 As Long     ' Object pointer
Private vtblptr             As Long     ' VTable pointer

' IActiveDesktop VTable wrapper functions
Public Function SetWallpaperOptions(ByVal pwpo As Long, ByVal dwReserved As Long) As Long
    SetWallpaperOptions = CallPointer(vtbl.SetWallpaperOptions, obj, pwpo, dwReserved)
End Function

Public Function SetPattern(ByVal pwszPattern As Long, ByVal dwReserved As Long) As Long
    SetPattern = CallPointer(vtbl.SetPattern, obj, pwszPattern, dwReserved)
End Function

Public Function SetDesktopItemOptions(ByVal pcomp As Long, ByVal dwReserved As Long) As Long
    SetDesktopItemOptions = CallPointer(vtbl.SetDesktopItemOptions, obj, pcomp, dwReserved)
End Function

Public Function RemoveDesktopItem(ByVal pcomp As Long, ByVal dwReserved As Long) As Long
    RemvoeDesktopItem = CallPointer(vtbl.RemoveDesktopItem, obj, pcomp, dwReserved)
End Function

Public Function ModifyDesktopItem(ByVal pcomp As Long, ByVal dwFlags As Long) As Long
    ModifyDesktopItem = CallPointer(vtbl.ModifyDesktopItem, obj, pcomp, dwFlags)
End Function

Public Function GetWallpaperOptions(ByVal pwpo As Long, ByVal dwReserved As Long) As Long
    GetWallpaperOptions = CallPointer(vtbl.GetWallpaperOptions, obj, pwpo, dwReserved)
End Function

Public Function GetPattern(ByVal pwszPattern As String, ByVal cchPattern As Long, ByVal dwReserved As Long) As Long
    GetPattern = CallPointer(vtbl.GetPattern, obj, pwszPattern, cchPattern, dwReserved)
End Function

Public Function GetDesktopItemOptions(ByVal pco As Long, ByVal dwReserved As Long) As Long
    GetDesktopItemOptions = CallPointer(vtbl.GetDesktopItemCount, obj, pco, dwReserved)
End Function

Public Function GetDesktopItemCount(ByVal lpiCount As Long, ByVal dwReserved As Long) As Long
    GetDeskopItemCount = CallPointer(vtbl.GetDesktopItemCount, obj, lpiCount, dwReserved)
End Function

Public Function GetDesktopItemBySource(ByVal pszSource As Long, ByVal pcomp As Long, ByVal dwReserved As Long) As Long
    GetDesktopItemBySource = CallPointer(vtbl.GetDesktopItemBySource, obj, pszSource, pcomp, dwReserved)
End Function

Public Function GetDesktopItemByID(ByVal dwID As Long, ByVal pcomp As Long, ByVal dwReserved As Long) As Long
    GetDesktopItemByID = CallPointer(vtbl.GetDesktopItemByID, obj, dwID, pcomp, dwReserved)
End Function

Public Function GetDesktopItem(ByVal nComponent As Long, ByVal pcomp As Long, ByVal dwReserved As Long) As Long
    GetDeskopItem = CallPointer(vtbl.GetDesktopItem, obj, nComponent, pcomp, dwReserved)
End Function

Public Function GenerateDesktopItemHtml(ByVal pwszFilename As String, ByVal pcomp As Long, ByVal dwResverved As Long) As Long
    GenerateDesktopItemHtml = CallPointer(vtbl.GenerateDesktopItemHtml, obj, pwszFilename, pcomp, dwResverved)
End Function

Public Function AddUrl(ByVal hwnd As Long, ByVal pszSource As Long, ByVal pcomp As Long, ByVal dwFlags As Long) As Long
    AddUrl = CallPointer(vtbl.AddUrl, obj, hwnd, pszSource, pcomp, dwFlags)
End Function

Public Function AddDesktopItemWithUI(ByVal hwnd As Long, ByVal pcomp As Long, ByVal dwFlags As Long) As Long
    AddDesktopItemWithUI = CallPointer(vtbl.AddDesktopItemWithUI, obj, hwnd, pcomp, dwFlags)
End Function

Public Function AddDesktopItem(ByVal pcomp As Long, ByVal rsvd As Long) As Long
    AddDesktopItem = CallPointer(vtbl.AddDesktopItem, obj, pcomp, rsvd)
End Function

Public Function ApplyChanges(ByVal flags As Long) As Long
    ApplyChanges = CallPointer(vtbl.ApplyChanges, obj, flags)
End Function

Public Function SetWallpaper(ByVal pwszWallpaper As Long, ByVal rsvd As Long) As Long
    SetWallpaper = CallPointer(vtbl.SetWallpaper, obj, pwszWallpaper, rsvd)
End Function

Public Function GetWallpaper(ByVal pwszWallpaper As Long, ByVal cchWallpaper As Long, ByVal rsvd As Long) As Long
    GetWallpaper = CallPointer(vtbl.GetWallpaper, obj, pwszWallpaper, cchWallpaper, rsvd)
End Function

' IUnknown VTable wrapper functions
Public Function Release() As Long
    Release = CallPointer(vtbl.Release, obj)
End Function

Public Function AddRef() As Long
    AddRef = CallPointer(vtbl.AddRef, obj)
End Function

Public Function QueryInterface(ByVal riid As Long, ByVal ppvObject As Long) As Long
    QueryInterface = CallPointer(vtbl.QueryInterface, obj, riid, ppvObject)
End Function

' CLSID of this interface
Private Property Get interface_CLSID() As String
    interface_CLSID = CLSID
End Property

Private Property Get interface_IID() As String
    interface_IID = IID
End Property

' new Object, new VTable
Private Property Let interface_object(RHS As Long)
    obj = RHS
    RtlMoveMemory vtblptr, ByVal obj, 4
    RtlMoveMemory vtbl, ByVal vtblptr, Len(vtbl)
End Property

Private Property Get interface_object() As Long
    interface_object = obj
End Property
