<div align="center">

## Auto\-Scrolling TextBox


</div>

### Description

A few lines of useful code to make a text box auto scroll to the bottom every time it changes. (like mIRC and msn)

Simply replace Text1 with the name of your textbox.

By Alex. (http://www.alexchia.com)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alexander Chia](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alexander-chia.md)
**Level**          |Beginner
**User Rating**    |4.9 (94 globes from 19 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alexander-chia-auto-scrolling-textbox__1-3511/archive/master.zip)





### Source Code

```
Private Sub Text1_Change()
 On Error Resume Next
 Text1.SelLength = 0
 If Len(Text1.Text) > 0 Then
 If Right$(Text1.Text,1) = vbCrLf Then
  Text1.SelStart = Len(Text1.Text) - 1
  Exit Sub
 End If
 Text1.SelStart = Len(Text1.Text)
 End If
End Sub
```

