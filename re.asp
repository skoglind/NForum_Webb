<%
regkod = UCase(request.QueryString("z"))

Set ra = New RegExp 
  ra.Global = True
  ra.IgnoreCase = True
  
  ra.Pattern = "([ABCDEGLMNOPRSTUV]{0,4})[-\s\.]{0,1}([ABCDEFGHIJKLMNOPQRSTUVWXYZ]{0,4})[-\s\.]{0,1}([ABCDEFGHIJKLMNOPQRSTUVWXYZ]{0,3})(.*)"
  
  reConsole = ra.Replace(regKod, "$1")
  reGame = ra.Replace(regKod, "$2")
  reRegion = ra.Replace(regKod, "$3")
  
  Response.Write "P1=" & reConsole & "<br>"
  Response.Write "P2=" & reGame & "<br>"
  Response.Write "P3=" & reRegion & "<br>"
Set ra = Nothing

If Len(reConsole) > 2 And Len(reGame) > 1 Then
  qSimple = reConsole & "-" & reGame
  If Len(reRegion) > 0 Then qFull = qSimple & "-" & reRegion
Else
  Response.Write "No hit"
End If
%>