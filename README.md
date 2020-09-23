<div align="center">

## Common Dialog without OCX


</div>

### Description

Hi All,

The Perpose of this Progarm is to Use windows Common Dialog Control Control Without the COMDLG32.OCX file. This will work even if the File is not Present

This is only for Open and Save Functions. But You can append it to get Color and other Dialog Boxes too,

Just Send any comments to

visual_basic@manjulapra.com

Visit me at

http://www.manjulapra.com

Thank You
 
### More Info
 
The Filter for the Common Dialog

The Default Extention for the Common Dialog

Optionally the Dialog Titile

The Path of the Selected File

None Identified


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Manjula Dharmawardhana](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/manjula-dharmawardhana.md)
**Level**          |Intermediate
**User Rating**    |4.8 (58 globes from 12 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/manjula-dharmawardhana-common-dialog-without-ocx__1-13368/archive/master.zip)

### API Declarations

```
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private strfileName As OPENFILENAME
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
```


### Source Code

```
'This Function sets the Filters for the Common Dialog
'It is basically the Same as in Commondialog OCX But when You want Multiple Filter Use as
'"All Files|*.*|Executable Files|*.exe"
Private Sub DialogFilter(WantedFilter As String)
  Dim intLoopCount As Integer
  strfileName.lpstrFilter = ""
  For intLoopCount = 1 To Len(WantedFilter)
    If Mid(WantedFilter, intLoopCount, 1) = "|" Then strfileName.lpstrFilter = _
    strfileName.lpstrFilter + Chr(0) Else strfileName.lpstrFilter = _
    strfileName.lpstrFilter + Mid(WantedFilter, intLoopCount, 1)
  Next intLoopCount
  strfileName.lpstrFilter = strfileName.lpstrFilter + Chr(0)
End Sub
'This is The Function To get the File Name to Open
'Even If U don't specify a Title or a Filter it is OK
Public Function fncGetFileNametoOpen(Optional strDialogTitle As String = "Open", Optional strFilter As String = "All Files|*.*", Optional strDefaultExtention As String = "*.*") As String
Dim lngReturnValue As Long
Dim intRest As Integer
  strfileName.lpstrTitle = strDialogTitle
  strfileName.lpstrDefExt = strDefaultExtention
  DialogFilter (strFilter)
  strfileName.hInstance = App.hInstance
  strfileName.lpstrFile = Chr(0) & Space(259)
  strfileName.nMaxFile = 260
  strfileName.flags = &H4
  strfileName.lStructSize = Len(strfileName)
  lngReturnValue = GetOpenFileName(strfileName)
  fncGetFileNametoOpen = strfileName.lpstrFile
End Function
'This Function Returns the Save File Name
'Remember, U have to Specify a Filter and default Extention for this
Public Function fncGetFileNametoSave(strFilter As String, strDefaultExtention As String, Optional strDialogTitle As String = "Save") As String
Dim lngReturnValue As Long
Dim intRest As Integer
  strfileName.lpstrTitle = strDialogTitle
  strfileName.lpstrDefExt = strDefaultExtention
  DialogFilter (strFilter)
  strfileName.hInstance = App.hInstance
  strfileName.lpstrFile = Chr(0) & Space(259)
  strfileName.nMaxFile = 260
  strfileName.flags = &H80000 Or &H4
  strfileName.lStructSize = Len(strfileName)
  lngReturnValue = GetSaveFileName(strfileName)
  fncGetFileNametoSave = strfileName.lpstrFile
End Function
```

