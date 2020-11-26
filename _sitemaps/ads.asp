<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">

  <!--#INCLUDE FILE="../__INC/configuration.asp"-->
  <!--#INCLUDE FILE="../__INC/md5.asp"-->
  <!--#INCLUDE FILE="../__INC/functions.asp"-->

  <%
  Response.ContentType = "text/xml"
  
  Con_Open()

    RS_Open 1, "SELECT TOP 2500 ksID, ksSkapadDatum FROM cms_KopSalj WHERE ksID > 0 AND ksSkapadDatum + " & CLng(config_AdDays) & " > '" & Now & "' AND ksSynlig = 1 ORDER BY ksSkapadDatum DESC", False
    
      If rsDB(1).EOF Then
        any_Kos = False
      Else
        any_Kos = True
        list_Kos = rsDB(1).GetRows
      End If
    
    RS_Close 1
  
  Con_Close()
  %>
  
  <url>
    <loc>http://www.n-forum.se/avdelning/annonser/default.asp</loc>
    <lastmod><% = Date %>T<% = Time %>+01:00</lastmod>
    <changefreq>always</changefreq>
  </url>
  
  <% If any_Kos Then %>
    <% For zx = 0 To UBound(list_Kos, 2) %>
      <url>
        <loc>http://www.n-forum.se/avdelning/annonser/annons_visa.asp?e=<% = list_Kos(0,zx) %></loc>
        <lastmod><% = FormatDateTime(list_Kos(1,zx), vbShortDate) %>T<% = FormatDateTime(list_Kos(1,zx), vbLongTime) %>+01:00</lastmod>
        <changefreq>daily</changefreq>
      </url>
    <% Next %>
  <% End if %>
</urlset>