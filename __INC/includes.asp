<!--#INCLUDE FILE="configuration.asp"-->
<!--#INCLUDE FILE="md5.asp"-->
<!--#INCLUDE FILE="constant.asp"-->
<!--#INCLUDE FILE="lists.asp"-->
<!--#INCLUDE FILE="functions.asp"-->
<!--#INCLUDE FILE="bbcode.asp"-->
<!--#INCLUDE FILE="db_reqs.asp"-->
<!--#INCLUDE FILE="_banmail.asp"-->

<% Con_Open %>

<%
  ' ## CookieLogin

  If NOT Session.Value("NFORUM_Login") Then
    lA = Request.Cookies("NFORUM")("A")
    lP = Request.Cookies("NFORUM")("P")

    If Len(lA) > 1 And Len(lP) > 1 Then
      ' ### KÖR INLOGGNINGSRUTINEN
      LoginUser lA, lP, True, True, ""
      
      If Session.Value("NFORUM_Login") Then SetConstants()
    End If
  End If
  
  onURL = Server.URLEncode(Request.ServerVariables("PATH_INFO") & "?" & Request.ServerVariables("QUERY_STRING"))
%>