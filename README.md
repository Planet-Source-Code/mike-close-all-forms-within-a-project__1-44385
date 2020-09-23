<div align="center">

## Close all forms within a project\.


</div>

### Description

Add this code to a module and simply use CloseAll to unload all forms in your project.
 
### More Info
 
None that I know off, I haven't had any problems yet... a decrease in use of Unload and End?


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-close-all-forms-within-a-project__1-44385/archive/master.zip)





### Source Code

```
Sub CloseAll()
On Error Resume Next
Dim intFrmNum As Integer
 intFrmNum = Forms.Count
Do Until intFrmNum = 0
 Unload Forms(intFrmNum - 1)
 intFrmNum = intFrmNum - 1
Loop
End Sub
```

