Attribute VB_Name = "MyNewModIcon"
'DrawIconEx Minimum operating systems Windows 95, Windows NT 3.5
'OleCreatePictureIndirect Requirements
  'Windows NT/2000/XP: Requires Windows NT 4.0 or later.
  'Windows 95/98: Requires Windows 95 or later.
  'Header: Declared in olectl.h.
  'Import Library: Included as a resource in olepro32.dll.
'ExtractIcon Minimum operating systems Windows 95, Windows NT 3.1
'DestroyIcon Minimum operating systems Windows 95, Windows NT 3.1


'API For GetIconFromFileIndex2
'*****************************
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_NORMAL = 3
'*****************************
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function ExtractIcon Lib "shell32" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type
Private Type CLSID
    id((123)) As Byte
End Type
Public Function GetIconFromFileIndex(sIconFile As String, icon_index As Integer, APicture As PictureBox)
Dim lpUnk As IUnknown
Dim cls_id As CLSID
Dim new_icon As TypeIcon
Dim thisRow As Long
Dim hIcon As Long
APicture.Picture = Nothing
hIcon = ExtractIcon(0&, sIconFile, icon_index)
With new_icon
    .cbSize = Len(new_icon)
    .picType = vbPicTypeIcon
    .hIcon = hIcon
End With
With cls_id
    .id(8) = &HC0
    .id(15) = &H46
End With
hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
If hRes = 0 Then Set icon_pic = lpUnk
APicture = icon_pic
Call DestroyIcon(hIcon)
End Function
'GetIconFromFileIndex2 not used because it generates a SavePicture error. I left it here so to someday get it to work, the reason why is it generates a smaller icon using DrawIconEx.
Public Function GetIconFromFileIndex2(sIconFile As String, icon_index As Integer, APicture As PictureBox)
Dim lpUnk As IUnknown
Dim cls_id As CLSID
Dim new_icon As TypeIcon
Dim thisRow As Long
Dim hIcon As Long
APicture.Picture = Nothing
hIcon = ExtractIcon(0&, sIconFile, icon_index)
  With APicture
    .Picture = LoadPicture("")
    .AutoRedraw = True
     Call DrawIconEx(.hdc, 0, 0, hIcon, 16, 16, 0, 0, DI_NORMAL)
    .Refresh
End With
APicture = icon_pic
Call DestroyIcon(hIcon)
End Function
