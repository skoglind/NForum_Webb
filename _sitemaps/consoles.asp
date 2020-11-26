<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">

  <!--#INCLUDE FILE="../__INC/configuration.asp"-->
  <!--#INCLUDE FILE="../__INC/md5.asp"-->
  <!--#INCLUDE FILE="../__INC/functions.asp"-->

  <%
  Response.ContentType = "text/xml"
  
  Con_Open()

    RS_Open 1, "SELECT TOP 2500 tID, kDatumSparad " & _
               "FROM cms_KonsolTitlar " & _
               "LEFT JOIN cms_Konsol ON cms_KonsolTitlar.tKonsolID = cms_Konsol.kID " & _ 
               "WHERE tID = kStandard_Titel AND kSynlig = 1 AND kDatumSparad <> 0 " & _
               "ORDER BY kDatumSparad DESC", False
    
      If rsDB(1).EOF Then
        any_Konsol = False
      Else
        any_Konsol = True
        list_Konsol = rsDB(1).GetRows
      End If
    
    RS_Close 1
  
  Con_Close()
  %>
  
  <url>
    <loc>http://www.n-forum.se/avdelning/konsol/default.asp</loc>
    <lastmod><% = Date %>T<% = Time %>+01:00</lastmod>
    <changefreq>always</changefreq>
  </url>
  
  <% If any_Konsol Then %>
    <% For zx = 0 To UBound(list_Konsol, 2) %>
      <url>
        <loc>http://www.n-forum.se/avdelning/konsol/konsol_visa_info.asp?e=<% = list_Konsol(0,zx) %></loc>
        <lastmod><% = FormatDateTime(list_Konsol(1,zx), vbShortDate) %>T<% = FormatDateTime(list_Konsol(1,zx), vbLongTime) %>+01:00</lastmod>
        <changefreq>daily</changefreq>
      </url>
    <% Next %>
  <% End if %>
</urlset>