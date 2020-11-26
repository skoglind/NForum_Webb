<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">

  <!--#INCLUDE FILE="../__INC/configuration.asp"-->
  <!--#INCLUDE FILE="../__INC/md5.asp"-->
  <!--#INCLUDE FILE="../__INC/functions.asp"-->

  <%
  Response.ContentType = "text/xml"
  
  Con_Open()

    RS_Open 1, "SELECT TOP 2500 nID, nDatumPublicerad FROM cms_Nyheter WHERE nStatus = 4 AND nDatumPublicerad <= '" & Now & "' ORDER BY nDatumPublicerad DESC", False
    
      If rsDB(1).EOF Then
        any_News = False
      Else
        any_News = True
        list_News = rsDB(1).GetRows
      End If
    
    RS_Close 1
  
  Con_Close()
  %>
  
  <url>
    <loc>http://www.n-forum.se/avdelning/nyheter/default.asp</loc>
    <lastmod><% = Date %>T<% = Time %>+01:00</lastmod>
    <changefreq>always</changefreq>
  </url>
  
  <% If any_News Then %>
    <% For zx = 0 To UBound(list_News, 2) %>
      <url>
        <loc>http://www.n-forum.se/avdelning/nyheter/nyheter_visa.asp?e=<% = list_News(0,zx) %></loc>
        <lastmod><% = FormatDateTime(list_News(1,zx), vbShortDate) %>T<% = FormatDateTime(list_News(1,zx), vbLongTime) %>+01:00</lastmod>
        <changefreq>daily</changefreq>
      </url>
    <% Next %>
  <% End if %>
</urlset>