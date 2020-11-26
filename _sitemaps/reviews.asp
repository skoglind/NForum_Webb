<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">

  <!--#INCLUDE FILE="../__INC/configuration.asp"-->
  <!--#INCLUDE FILE="../__INC/md5.asp"-->
  <!--#INCLUDE FILE="../__INC/functions.asp"-->

  <%
  Response.ContentType = "text/xml"
  
  Con_Open()

    RS_Open 1, "SELECT TOP 2500 rID, rDatumPublicerad " & _
               "FROM cms_Recensioner " & _
               "WHERE rStatus = 4 " & _
               "AND rDatumPublicerad <= '" & Now & "' " & _
               "ORDER BY rDatumPublicerad DESC", False
    
      If rsDB(1).EOF Then
        any_Txt = False
      Else
        any_Txt = True
        list_Txt = rsDB(1).GetRows
      End If
    
    RS_Close 1
  
  Con_Close()
  %>
  
  <url>
    <loc>http://www.n-forum.se/avdelning/recensioner/default.asp</loc>
    <lastmod><% = Date %>T<% = Time %>+01:00</lastmod>
    <changefreq>always</changefreq>
  </url>
  
  <% If any_Txt Then %>
    <% For zx = 0 To UBound(list_Txt, 2) %>
      <url>
        <loc>http://www.n-forum.se/avdelning/recensioner/recension_visa.asp?e=<% = list_Txt(0,zx) %></loc>
        <lastmod><% = FormatDateTime(list_Txt(1,zx), vbShortDate) %>T<% = FormatDateTime(list_Txt(1,zx), vbLongTime) %>+01:00</lastmod>
        <changefreq>daily</changefreq>
      </url>
    <% Next %>
  <% End if %>
</urlset>