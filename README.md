<div align="center">

## Path manipulation


</div>

### Description

There are many methods you can use to return certain parts of a fully qualified path to a file. Here is the SHORTEST and FASTEST way to 1) return just the path, 2) return just the filename, and 3) change the extension of a filename. The code is so short that it is probably faster to keep it inline than to create additional functions. (I've done so here to better illustrate the parameters.
 
### More Info
 
I've intentionally left out doing any checks for a valid path in order to show the simplicity of the code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Wilson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-wilson.md)
**Level**          |Advanced
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 6\.0, VB Script
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-wilson-path-manipulation__1-11326/archive/master.zip)





### Source Code

```
Public Sub PathTest
' Return just the path "c:\test\"
' TRUE strips the backslash, FALSE retains it
Debug.Print JustPath("c:\test\myfile.txt", "\", True)
' Return just the filename "myfile.txt"
' Change "\" to "/" to handle UNIX or URL pathnames!
Debug.Print JustFile("c:\test\myfile.txt", "\")
' Change the extension to "bak" and return "c:\test\myfile.bak"
Debug.Print ChangeExt("c:\test\myfile.txt", "bak")
' Change the extension and return just the filename "myfile.bak"
' Change "\" to "/" to handle UNIX or URL pathnames!
Debug.Print JustFile(ChangeExt("c:\test\myfile.txt", "bak"), "\")
End Sub
Public Function JustPath(ByVal filepath As String, ByVal dirchar As String, ByVal stripbs As Integer) As String
	' Returns just the path
	' TRUE evaluates to -1, FALSE evaluates to 0 so
	' simple addition is all we need at the end to remove the slash
	JustPath = Mid$(filepath, 1, InStrRev(filepath, dirchar) + stripbs)
End Function
Public Function JustFile(ByVal filepath As String, ByVal dirchar As String) As String
 ' Returns just the filename
 JustFile = Mid$(filepath, InStrRev(filepath, dirchar) + 1)
End Function
Public Function ChangeExt(ByVal filepath As String, ByVal newext As String) As String
 ' Changes the extension
 ChangeExt = Mid$(filepath, 1, InStrRev(filepath, ".")) & newext
End Function
```

