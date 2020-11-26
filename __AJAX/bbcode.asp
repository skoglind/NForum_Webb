<!--#INCLUDE FILE="../__INC/includes.asp"-->
  <%
  retValue = GetQ("txt", "ABC", 20000)
  retAnchor = GetQ("a", "CHK", 0)
  retSmilies = GetQ("s", "CHK", 0)
  If Len(retValue) > 0 Then
    
    If retAnchor Then retValue = TraceHyperlinks(retValue)
    retValue = BBCode(sEncode(retValue), retSmilies)
    
    Response.Write retValue
    
  End If
  %>

<!--#INCLUDE FILE="../__INC/includes_end.asp"-->