<div align="center">

## HTML to VB color converter \- A must have for web programmers\!


</div>

### Description

I have a program I'm working on right now where the user settings for the app are done on a website. What I didn't realize was that with HTML colours being stored in RRGGBB format, you can't automatically use it in VB, because VB uses the &HBBGGRR format. Here's a little function to bring back the right colour.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brandon McPherson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brandon-mcpherson.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brandon-mcpherson-html-to-vb-color-converter-a-must-have-for-web-programmers__1-23058/archive/master.zip)





### Source Code

<font size="2">Function MakeVBColour(hColor) As Long<br>
' 20010509 BWM - Used to flip the <br>
' #RRGGBB HTML colour format to the<br>
' VB-style &HBBGGRR format<br>
' Note: the variable 'RED' refers to 'BLUE'<br>
' in HTML, and 'BLUE' refers to 'RED' in HTML.<br>
' There's no standard.<br>
 Dim Red As Long<br>
 Dim Green As Long<br>
 Dim Blue As Long<br>
 Dim sRed As String<br>
 Dim sBlue As String<br>
 Dim sGreen As String<br>
' Fill a long variable with the colour
<br>
 hColor = CLng(hColor)<br>
' Separate the colours into their own variables<br>
 Red = hColor And 255<br>
 Green = (hColor And 65280) \ 256<br>
 Blue = (hColor And 16711680) \ 65535<br>
' Get the hex equivalents<br>
 sRed = Hex(Red)<br>
 sBlue = Hex(Blue)<br>
 sGreen = Hex(Green)<br>
' Pad each colour, to make sure it's 2 chars<br>
 sRed = String(2 - Len(sRed), "0") & sRed<br>
 sBlue = String(2 - Len(sBlue), "0") & sBlue<br>
 sGreen = String(2 - Len(sGreen), "0") & sGreen<br>
'reassemble' the colour<br>
 MakeVBColour = CLng("&H" & sRed & sGreen & sBlue)<br>
End Function<br></font>

