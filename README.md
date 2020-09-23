<div align="center">

## Validate\_Drive


</div>

### Description

Validates whether a given hard/floppy/network drive is valid
 
### More Info
 
strDrive--drive to validate

i the dirve exists returns TRUE, otherwise FALSE


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\)
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-validate-drive__1-39/archive/master.zip)





### Source Code

```
Function Validate_Drive (ByVal strDrive As String)
  On Error GoTo BAD2
    'Dim strOldDrive As String
    'strOldDrive = Get_Drive_Name(CurDir$)
    ChDrive (strDrive)
    'ChDrive (strOldDrive)
  On Error GoTo 0
  Validate_Drive = True
Exit Function
BAD2:
  Validate_Drive = False
  Resume Exit2
Exit2:
  Exit Function
End Function
```

