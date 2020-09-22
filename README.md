<div align="center">

## Recursive Scan Directory


</div>

### Description

The first procedure scans any directory specified and all its subdirectories and fills an array with all the files found.

The other procedure returns the number of files in the directory specified and all its sub directories.
 
### More Info
 
the inputs are the "Directory to scan", the number of files(which is only required if u wanna have a progress bar,the fpgrid (third party grid)not used if u don't have the FarPoint Grid...can be modifed to use any other grid, listbox etc.

include the Microsoft Scripting Runtime from references

nothing, it has a public filelist array which stores all the files with their paths.

U must specify the sFol(directory to scan) parameter as i have not included any checks in this code for that.(if it is blank)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Anumeet Soni](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/anumeet-soni.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/anumeet-soni-recursive-scan-directory__1-44631/archive/master.zip)





### Source Code

```
Option Explicit
Dim fso As New FileSystemObject 'The file system object
Dim ParFolder As Folder     'parent folder variable
Dim n As Long          'for counting
Public Filelist() As String   'array to hold the list of files with path
'---------------------------------------------------------------------------------------
' Procedure : Public Function FindFile(Optional ByVal sFol As String, Optional ByVal NumberFiles As Long, Optional ByVal fp As vaSpread)
' DateTime : April 8th 2003, 3:48 PM
' Author  : Anumeet Son
' Purpose  : Gets the List of files and store in array, in a specified folder
'       and all its subfolders(using FileSystemObject)
'       using "Microsoft Scripting Runtime"
'       YOU CAN EITHER STORE IT IN THE ARRAY OR USE IT AS REQUIRED FROM HERE
'       ONLY
'---------------------------------------------------------------------------------------
'PURPOSE OF NUMBERFILES IS FOR INCLUDING A PROGRESS BAR(OPTIONAL AND FP AND VASPREAD
'ARE THIRD PARTY CONTROLS(GRID) IN WHICH I AM POPULATING THE FILE
Public Function FindFile(Optional ByVal sFol As String, Optional ByVal NumberFiles As Long, Optional ByVal fp As vaSpread)
    Dim CurFile As File
    Dim CurFolder As Folder
    Dim NFiles As Long
    Set ParFolder = fso.GetFolder(sFol)
    NFiles = ParFolder.Files.Count
    If NFiles > 0 Then
      For Each CurFile In ParFolder.Files
        Filelist(n) = CurFile.Path   'STORE THE FILE IN ARRAY
        fp.SetText 1, n, Filelist(n)
        n = n + 1            'INCREASE COUNTER BY 1
      Next
    End If
    For Each CurFolder In ParFolder.SubFolders 'IF SUBFOLDERS OF CURRENT FOLDER ARE THERE
      FindFile CurFolder, , fp        'call itself to get the files of subfolders
    Next
End Function
'---------------------------------------------------------------------------------------
' Procedure : Public Function FindNoFiles(ByVal sFol As String)
' DateTime : April 8th 2003, 3:48 PM
' Author  : Anumeet Soni
' Purpose  : Gets the number of files, in a specified folder
'       and all its subfolders(using FileSystemObject)
'       using Microsoft Scripting Runtime
'---------------------------------------------------------------------------------------
Public Function FindNoFiles(ByVal sFol As String)
    Dim tFld As Folder
    Set ParFolder = fso.GetFolder(sFol)
    FindNoFiles = ParFolder.Files.Count
    If ParFolder.SubFolders.Count > 0 Then
      For Each tFld In ParFolder.SubFolders
        FindNoFiles = FindNoFiles + FindNoFiles(tFld.Path)
      Next
    End If
End Function
```

