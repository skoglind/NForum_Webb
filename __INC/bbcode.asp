<%
' VBScript BBCode to HTML convertor by snuke
' http://alizahid.net/
' snukeNetwork
Function Reggex(strString, strPattern, strReplace)

  Dim RE: Set RE = New RegExp

  With RE
    .Pattern = strPattern
    .Global = True
    .IgnoreCase = True
    Reggex = .Replace(strString, strReplace)
  End With
  
End Function

Function BBCode(sText, bSmilies)
  sText = " " & sText & " "  

  sText = Reggex(sText, "\[newline\]", Chr(13) & Chr(10))
  sText = Reggex(sText, "\[ampersand\]", "&")
  sText = Reggex(sText, "\[bracket\]", "#")

  sText = Reggex(sText, "\[list\][\r\n]{0,3}", "</p><ul class='postlist'>")
  sText = Reggex(sText, "\[\/list\][\r\n]{0,3}", "</ul><p>")
  
  sText = Reggex(sText, "\[\*\](.*?)\n", "<li>$1</li>")

  sText = Reggex(sText, "\[quote\]", "</p><div class='quote'><p>")
  sText = Reggex(sText, "\[\/quote\][\r\n]{0,3}", "</p></div><p>")
  
  sText = Reggex(sText, "[\r\n]{0,3}\[rubrik\]", "</p><h3>")
  sText = Reggex(sText, "\[\/rubrik\][\r\n]{0,3}", "</h3><p>")
  
  sText = Reggex(sText, "\[table\][\r\n]{0,3}", "</p><table>")
  sText = Reggex(sText, "\[\/table\][\r\n]{0,3}", "</table><p>")
  sText = Reggex(sText, "\[tr\][\r\n]{0,3}", "<tr>")
  sText = Reggex(sText, "\[\/tr\][\r\n]{0,3}", "</tr>")
  sText = Reggex(sText, "\[td\][\r\n]{0,3}", "<td>")
  sText = Reggex(sText, "\[\/td\][\r\n]{0,3}", "</td>")
  sText = Reggex(sText, "\[th\][\r\n]{0,3}", "<th>")
  sText = Reggex(sText, "\[\/th\][\r\n]{0,3}", "</th>")
  
  sText = Replace(sText, Chr(13) & Chr(10), "<br>")

  sText = Reggex(sText, "\[url=([^\]]+?)\](.+?)\[\/url\]", "<a href='$1' rel='nofollow' target='_blank' title='$1'>$2</a>")
  sText = Reggex(sText, "\[url\](.+?)\[\/url\]", "<a href='$1' rel='nofollow' target='_blank' title='$1'>$1</a>")
  sText = Reggex(sText, "\[img\](.+?)\[\/img\]", "<img class='imgInner' src='$1' alt='$1'>")
  
  sText = Reggex(sText, "\[dbimg\]([0-9]+)\[\/dbimg\]", "</p><div class='imgInnerDB'><a href='" & config_ImageLocation & "?e=$1&amp;w=800&h=600' target='_blank' rel='lightbox[intext]'><img class='imgInnerDB' src='" & config_ImageLocation & "?e=$1&amp;w=320&h=240' alt='$1' title='Klicka för att se bilden i originalformat.'></a></div><p style='margin-top: 0;'>")
  sText = Reggex(sText, "\[dblnk\]([0-9]+)\[\/dblnk\]", "<a href='" & config_ImageLocation & "?e=$1&amp;w=640&h=480' target='_blank' rel='lightbox' alt='Bild' title='Klicka för att se bilden.'>[BILD]</a>")
  sText = Reggex(sText, "\[dbthumb\]([0-9]+)\[\/dbthumb\]", "<a href='" & config_ImageLocation & "?e=$1&amp;w=640&h=480' target='_blank' rel='lightbox' alt='Bild' title='Klicka för att se bilden.'><img src='" & config_ImageLocation & "?e=$1&amp;w=80&h=80' style='float:left;width:80px;height:80px;margin:2px;padding:2px;border:solid 1px #CCC;background-color:#FFF;'></a>")

  sText = Reggex(sText, "\[b\]", "<strong>")
  sText = Reggex(sText, "\[\/b\]", "</strong>")
  sText = Reggex(sText, "\[i\]", "<em>")
  sText = Reggex(sText, "\[\/i\]", "</em>")
  sText = Reggex(sText, "\[u\]", "<span class=""underline"">")
  sText = Reggex(sText, "\[\/u\]", "</span>")
  sText = Reggex(sText, "\[s\]", "<strike>")
  sText = Reggex(sText, "\[\/s\]", "</strike>")

  sText = Reggex(sText, "\[youtube\](.+?)\[\/youtube\]", "<object width=438 height=360><param name='movie' value='http://www.youtube.com/v/$1'></param><param name='allowFullScreen' value='true'></param><param name='allowscriptaccess' value='always'></param><embed src='http://www.youtube.com/v/$1' type='application/x-shockwave-flash' allowscriptaccess='always' allowfullscreen='true' width=438 height=360></embed></object>")
  sText = Reggex(sText, "\[gametrailers\](.+?)\[\/gametrailers\]", "<object classid='clsid:d27cdb6e-ae6d-11cf-96b8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0' id='gtembed' width='438' height='358'> <param name='allowScriptAccess' value='sameDomain' /> <param name='allowFullScreen' value='true' /><param name='movie' value='http://www.gametrailers.com/remote_wrap.php?mid=$1'/><param name='quality' value='high' /> <embed src='http://www.gametrailers.com/remote_wrap.php?mid=$1' swLiveConnect='true' name='gtembed' align='middle' allowScriptAccess='sameDomain' allowFullScreen='true' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='438' height='358'></embed></object>")

  If bSmilies Then sText = ActivateSmilies(sText)
  
  BBCode = Trim(sText)
End Function

Function TinyCode(sText)
  sText = " " & sText & " "  
  
  sText = Replace(sText, Chr(13) & Chr(10), "<br>")

  sText = Reggex(sText, "\[url=([^\]]+?)\](.+?)\[\/url\]", "<a href='$1' rel='nofollow' target='_blank' title='$1'>$2</a>")
  sText = Reggex(sText, "\[url\](.+?)\[\/url\]", "<a href='$1' rel='nofollow' target='_blank' title='$1'>$1</a>")
 
  sText = Reggex(sText, "\[b\]", "<strong>")
  sText = Reggex(sText, "\[\/b\]", "</strong>")
  sText = Reggex(sText, "\[i\]", "<em>")
  sText = Reggex(sText, "\[\/i\]", "</em>")
  sText = Reggex(sText, "\[u\]", "<span class=""underline"">")
  sText = Reggex(sText, "\[\/u\]", "</span>")
  sText = Reggex(sText, "\[s\]", "<strike>")
  sText = Reggex(sText, "\[\/s\]", "</strike>")

  sText = ActivateSmilies(sText)
  
  TinyCode = Trim(sText)
End Function

Function BBCode_Remove(sText)
    Set ra = New RegExp 
      ra.Global = True
      ra.IgnoreCase = True
  
      ra.Pattern = "\[[^\]]*\]"
      sText = ra.Replace(sText, "")
      
      ra.Pattern = "<[^>]*>"
      sText = ra.Replace(sText, "")
      
    Set rs = Nothing
    
    BBCode_Remove = sText
  End Function
  
  Function ActivateSmilies(sText)
    sText = " " & sText & " "

    sText = Replace(sText, " :)"      ," <img src='" & config_GFXLocation & "icons/smilies/glad.gif' alt='Glad' title='Glad'>")
    sText = Replace(sText, " :("      ," <img src='" & config_GFXLocation & "icons/smilies/sorgsen.gif' alt='Sorgsen' title='Sorgsen'>")
    sText = Replace(sText, " 8)"      ," <img src='" & config_GFXLocation & "icons/smilies/cool.gif' alt='Cool' title='Cool'>")
    sText = Replace(sText, " >:("     ," <img src='" & config_GFXLocation & "icons/smilies/arg.gif' alt='Arg' title='Arg'>")
    sText = Replace(sText, " :D"      ," <img src='" & config_GFXLocation & "icons/smilies/flin.gif' alt='Flin' title='Flin'>")
    sText = Replace(sText, " :/"      ," <img src='" & config_GFXLocation & "icons/smilies/fundersam.gif' alt='Fundersam' title='Fundersam'>")
    sText = Replace(sText, " :s"      ," <img src='" & config_GFXLocation & "icons/smilies/forvirrad.gif' alt='Förvirrad' title='Förvirrad'>")
    sText = Replace(sText, " :$"      ," <img src='" & config_GFXLocation & "icons/smilies/generad.gif' alt='Generad' title='Generad'>")
    sText = Replace(sText, " :uack:"  ," <img src='" & config_GFXLocation & "icons/smilies/illamaende.gif' alt='Illamående' title='Illamående'>")
    sText = Replace(sText, " :p"      ," <img src='" & config_GFXLocation & "icons/smilies/lipa.gif' alt='Lipa' title='Lipa'>")
    sText = Replace(sText, " :roll:"  ," <img src='" & config_GFXLocation & "icons/smilies/rulla.gif' alt='Rullögon' title='Rullögon'>")
    sText = Replace(sText, " ;)"      ," <img src='" & config_GFXLocation & "icons/smilies/blink.gif' alt='Blink' title='Blink'>")

    ActivateSmilies = Trim(sText)
  End Function
%>