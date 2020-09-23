<div align="center">

## Show Printer Document Properties setup dialog


</div>

### Description

Shows the printer document properties dialog box from code.
 
### More Info
 
This entry in Shell32.dll is only present in version 4.71 and above (Windows NT 4 and Internet Explorer 4.0 or above)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Duncan Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-jones.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/duncan-jones-show-printer-document-properties-setup-dialog__1-22127/archive/master.zip)

### API Declarations

```
'SHSTDAPI_(BOOL) SHInvokePrinterCommandA(HWND hwnd, UINT uAction, LPCSTR lpBuf1, LPCSTR lpBuf2, BOOL fModal);
Private Declare Function SHInvokePrinterCommand Lib "shell32.dll" Alias "SHInvokePrinterCommandA" (ByVal hWnd As Long, ByVal uAction As enPrinterActions, ByVal Buffer1 As String, ByVal Buffer2 As String, ByVal Modal As Long) As Long
Public Enum enPrinterActions
   PRINTACTION_OPEN = 0
   PRINTACTION_PROPERTIES = 1
   PRINTACTION_NETINSTALL = 2
   PRINTACTION_NETINSTALLLINK = 3
   PRINTACTION_TESTPAGE = 4
   PRINTACTION_OPENNETPRN = 5
   PRINTACTION_DOCUMENTDEFAULTS = 6
   PRINTACTION_SERVERPROPERTIES = 7
End Enum
```


### Source Code

```
Public Sub DisplayDocumentDefaults(ByVal PrinterName As String, ByVal hWnd As Long)
Dim lRet As Long
'\\ Only version 4.71 and above have this :. jump over error
On Error Resume Next
lRet = SHInvokePrinterCommand(hWnd, PRINTACTION_DOCUMENTDEFAULTS, PrinterName, "", 0)
End Sub
```

