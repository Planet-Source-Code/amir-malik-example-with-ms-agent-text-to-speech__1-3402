<div align="center">

## Example with MS Agent, text\-to\-speech


</div>

### Description

Microsoft Agent will speak (computerized voice) the contents of a textbox or richtextbox to you!
 
### More Info
 
the text to speak, or just paste it

components/objects:

Microsoft Direct Text-to-Speech control

rich text box

2 command buttons

the sp is the DirectSS control

text box for the computer voice speed

sound!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Amir Malik](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/amir-malik.md)
**Level**          |Unknown
**User Rating**    |3.8 (30 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/amir-malik-example-with-ms-agent-text-to-speech__1-3402/archive/master.zip)





### Source Code

```
' MSAGENT example by Amir Malik
' website: http://amir142.cjb.net
' e-mail : amir@infoteen.com
Private Sub cmdPaste_Click()
  TextData.Text = Clipboard.GetText
End Sub
Private Sub cmdPauseR_Click()
  If cmdPauseR.Caption = "&Pause / Stop" Then
    sp.AudioPause
    cmdPauseR.Caption = "&Resume"
  ElseIf cmdPauseR.Caption = "&Resume" Then
    sp.AudioResume
    cmdPauseR.Caption = "&Pause / Stop"
  End If
End Sub
Private Sub cmdSpeak_Click()
  sp.Speak TextData.Text
  sp.Speed = txtSpeed.Text
  Sspeak = True
End Sub
Private Sub txtSpeed_LostFocus()
  If txtSpeed.Text < 50 Then
    MsgBox "Speed is too low."
    txtSpeed.Text = "150"
  End If
  If txtSpeed.Text > 250 Then
    MsgBox "Speed is too high."
    txtSpeed.Text = "150"
  End If
End Sub
```

