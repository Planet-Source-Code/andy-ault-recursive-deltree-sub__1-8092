<div align="center">

## Recursive Deltree Sub


</div>

### Description

Deletes a specified directory tree.
 
### More Info
 
sFolder: Folder tree to remove (ie "T:\TEMP")

Note that "Dir" doesn't work as you may expect with recursive programs. This code works around that problem by re-doing the Dir w/parameters after calling itself recursively.

No return. If you need to use this in situations where filesor folders might have read-only attributes or can't be deleted for other reasons, you will need to modify the code to add error handling and return an error code.

As always with this type of code, use it carefully. You could wipe valuable data.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andy Ault](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andy-ault.md)
**Level**          |Intermediate
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andy-ault-recursive-deltree-sub__1-8092/archive/master.zip)





### Source Code

```
Public Sub KillFolderTree(sFolder As String)
 Dim sCurrFilename As String
 sCurrFilename = Dir(sFolder & "\*.*", vbDirectory)
 Do While sCurrFilename <> ""
 If sCurrFilename <> "." And sCurrFilename <> ".." Then
  If (GetAttr(sFolder & "\" & sCurrFilename) And vbDirectory) = vbDirectory Then
  Call KillFolderTree(sFolder & "\" & sCurrFilename)
  sCurrFilename = Dir(sFolder & "\*.*", vbDirectory)
  Else
  On Error Resume Next
  Kill sFolder & "\" & sCurrFilename
  On Error Goto 0
  sCurrFilename = Dir
  End If
 Else
  sCurrFilename = Dir
 End If
 Loop
 On Error Resume Next
 RmDir sFolder
End Sub
```

