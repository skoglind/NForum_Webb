<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">

  <!--#INCLUDE FILE="../__INC/configuration.asp"-->
  <!--#INCLUDE FILE="../__INC/md5.asp"-->
  <!--#INCLUDE FILE="../__INC/functions.asp"-->

  <%
  Response.ContentType = "text/xml"
  
  Con_Open()

    RS_Open 1, "SELECT TOP 2500 tID, sDatumSparad " & _
               "FROM cms_SpelTitlar " & _
               "LEFT JOIN cms_Spel ON cms_SpelTitlar.tSpelID = cms_Spel.sID " & _ 
               "WHERE tID = sStandard_Titel AND sSynlig = 1 AND sDatumSparad <> 0 " & _
               "ORDER BY sDatumSparad DESC", False
    
      If rsDB(1).EOF Then
        any_Spel = False
      Else
        any_Spel = True
        list_Spel = rsDB(1).GetRows
      End If
    
    RS_Close 1
  
  Con_Close()
  %>
  
  <url>
    <loc>http://www.n-forum.se/avdelning/spel/default.asp</loc>
    <lastmod><% = Date %>T<% = Time %>+01:00</lastmod>
    <changefreq>always</changefreq>
  </url>
  
  <% If any_Spel Then %>
    <% For zx = 0 To UBound(list_Spel, 2) %>
      <url>
        <loc>http://www.n-forum.se/avdelning/spel/spel_visa_info.asp?e=<% = list_Spel(0,zx) %></loc>
        <lastmod><% = FormatDateTime(list_Spel(1,zx), vbShortDate) %>T<% = FormatDateTime(list_Spel(1,zx), vbLongTime) %>+01:00</lastmod>
        <changefreq>daily</changefreq>
      </url>
    <% Next %>
  <% End if %>
</urlset>