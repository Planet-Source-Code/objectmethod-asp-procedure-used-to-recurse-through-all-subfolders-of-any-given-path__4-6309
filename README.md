<div align="center">

## ASP Procedure Used To Recurse Through All SubFolders Of Any Given Path


</div>

### Description

Use it to recurse (walk-through) every sub-directory of any given path. Use it to find a file or list all files in any given path.
 
### More Info
 
"PATH" = Starting location of search

You can write your own code for each file or folder found.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ObjectMethod](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/objectmethod.md)
**Level**          |Intermediate
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/objectmethod-asp-procedure-used-to-recurse-through-all-subfolders-of-any-given-path__4-6309/archive/master.zip)





### Source Code

```
<%
	Sub Recurse(Path)
		Dim fso, Root, WindowsFolder, Files, _
			Folders, File, i, FoldersArray(100)
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		Set Root = fso.getfolder(Path)
		Set Files = Root.Files
		Set Folders = Root.SubFolders
		For Each File In Files
			'''''''''''''''''''''''''''''
			'Code For Each File Found
			'Goes Here!
			'''''''''''''''''''''''''''''
		Next
		For Each Folder In Folders
			'''''''''''''''''''''''''''''
			'Code For Each Folder Found
			'Goes Here!
			'''''''''''''''''''''''''''''
			FoldersArray(i) = Folder.Path
			i = i + 1
		Next
		For i = 0 To UBound(FoldersArray)
			If FoldersArray(i) <> "" Then
				Recurse FoldersArray(i)
			Else
				Exit For
			End If
		Next
	End Sub
%>
```

