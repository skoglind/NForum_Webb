<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">

  <!--#INCLUDE FILE="../__INC/configuration.asp"-->
  <!--#INCLUDE FILE="../__INC/md5.asp"-->
  <!--#INCLUDE FILE="../__INC/functions.asp"-->

  <%
  Response.ContentType = "text/xml"
  
  If Not config_LockDown_Forum Then
    Con_Open()
  
      RS_Open 1, "SELECT TOP 2500 tID, tStatus_Trad, tStatus_UnderTrad, tDatum_Skapad " & _
                 "FROM fsBB_Tradar AS tbTrad " & _
                 "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
                 "WHERE tDatum_Skapad <= '" & Now & "' AND fSec_View = '0' AND tStatus_Raderad = 0 ORDER BY tDatum_Skapad DESC", False
      
        If rsDB(1).EOF Then
          any_Tradar = False
        Else
          any_Tradar = True
          list_Tradar = rsDB(1).GetRows
        End If
      
      RS_Close 1
    
    Con_Close()
  End If
  %>
  
  <url>
    <loc>http://www.n-forum.se/avdelning/forum/</loc>
    <lastmod><% = Date %>T<% = Time %>+01:00</lastmod>
    <changefreq>always</changefreq>
  </url>
  
  <url>
    <loc>http://www.n-forum.se/avdelning/forum/forum.asp</loc>
    <lastmod><% = Date %>T<% = Time %>+01:00</lastmod>
    <changefreq>always</changefreq>
  </url>
  
  <% If any_Tradar Then %>
    <% For zx = 0 To UBound(list_Tradar, 2) %>
      <%
        If list_Tradar(1,zx) Then
          tradAdd = list_Tradar(0,zx)
        Else
          tradAdd = list_Tradar(2,zx) & "&amp;go2=" & list_Tradar(0,zx)
        End If
      %>
      <url>
        <loc>http://www.n-forum.se/avdelning/forum/trad.asp?e=<% = tradAdd %></loc>
        <lastmod><% = FormatDateTime(list_Tradar(3,zx), vbShortDate) %>T<% = FormatDateTime(list_Tradar(3,zx), vbLongTime) %>+01:00</lastmod>
        <changefreq>hourly</changefreq>
      </url>
    <% Next %>
  <% End if %>
</urlset>