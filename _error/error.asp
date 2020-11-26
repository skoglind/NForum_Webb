<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<% On Error Resume Next %>

<!--#INCLUDE FILE="../__INC/configuration.asp"-->
<!--#INCLUDE FILE="../__INC/constant.asp"-->
<!--#INCLUDE FILE="../__INC/functions.asp"-->

<%
Dim oErr
Set oErr = Server.GetLastError

sCategory     = Trim(oErr.Category & " ")
sColumn       = Trim(oErr.Column & " ")
sDescription  = Trim(oErr.Description & " ")
sFile         = Trim(oErr.File & " ")
sLine         = Trim(oErr.Line & " ")
sNumber       = Trim(oErr.Number & " ")
sSource       = Trim(oErr.Source & " ")

sIP           = Left(Request.ServerVariables("REMOTE_ADDR"),50)
sUserID       = CLng(CONST_USERID)

Con_Open

  RS_Open 1, "SELECT * FROM cms_Error WHERE eIP = '" & sIP & "' AND eUser = " & sUserID & " AND eNumber = '" & sNumber & "' AND eSource = '" & sSource & "' AND DATEDIFF(d, eDatum, '" & Now & "') = 0", True
    If rsDB(1).EOF Then
      rsDB(1).AddNew
        
        rsDB(1)("eIP")          = sIP
        rsDB(1)("eDatum")       = Now
        rsDB(1)("eUser")        = sUserID
        rsDB(1)("eReferer")     = Left(Request.ServerVariables("HTTP_REFERER"),500)
        rsDB(1)("eBrowser")     = Left(Request.ServerVariables("HTTP_USER_AGENT"),500)
        rsDB(1)("eCategory")    = Left(sCategory,255)
        rsDB(1)("eColumn")      = Left(sColumn,50)
        rsDB(1)("eDescription") = Left(sDescription,255)
        rsDB(1)("eFile")        = Left(sFile,255)
        rsDB(1)("eLine")        = Left(sLine,50)
        rsDB(1)("eNumber")      = Left(sNumber,50)
        rsDB(1)("eSource")      = Left(sSource,255)
      
      rsDB(1).Update
    End If
  RS_Close 1

Con_Close
%>

<html>
  <head>
    <title> 500 Error - Ett oväntat fel uppstod! </title>
    <style type="text/css">
      #holder {width: 400px; margin: 0 auto 0 auto;}
      #box {float: left; width: 400px; margin-top: 50px;}
      #felinfo {float: left; width: 400px; margin-top: 5px; border-top: dotted 2px #AAA;}
      #return {float: left; width: 400px; margin-top: 5px; border-top: dotted 2px #AAA;}
      
      #box h1 {margin: 0; padding: 0; color: #333; font: bold 26px Arial;}
      #box p {margin: 0 0 10px 0; padding: 0; color: #AAA; font: 15px Arial;}
      
      #felinfo h2 {margin: 6px 0 0 0; padding: 0; color: #333; font: bold 18px Arial;}
      #felinfo p {margin: 3px 0 3px 0; padding: 0; color: #666; font: 11px Arial;}
      
      #return h3 {margin: 20px 0 0 0; padding: 0; color: #333; font: bold 16px Arial;}
    </style>
  </head>
  <body>
  
    <div id="holder">
      <div id="box">
        <h1>Ett oväntat fel uppstod!</h1>
        <p>Vi ber så hemskt mycket om ursäkt för besväret.</p>
        <p>Vi har lagrat all information om felet och kommer åtgärda det snarast möjligt.</p>
        <p>Om felet skulle återkomma ofta kan du kontakta oss per e-post (<a href="mailto: info@n-forum.se">info@n-forum.se</a>) och bifoga informationen nedan.</p>
      </div>
      <div id="felinfo">
        <h2>Information om felet</h2>
        <p><strong>Felkod:</strong> <% = sNumber %></p>
        <p><strong>Beskrivning:</strong> <% = sDescription %></p>
        <p><strong>Kategori:</strong> <% = sCategory %></p>
        <p><strong>Fil:</strong> <% = sFile %></p>
        <p><strong>Rad:</strong> <% = sLine %></p>
        <p><strong>Kolumn:</strong> <% = sColumn %></p>
      </div>
      <div id="return">
        <h3><a href="/">« återgå till N-Forum.se</a></h3>
      </div>
    </div>
  
  </body>
</html>