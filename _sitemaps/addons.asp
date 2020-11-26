<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">

  <!--#INCLUDE FILE="../__INC/configuration.asp"-->
  <!--#INCLUDE FILE="../__INC/md5.asp"-->
  <!--#INCLUDE FILE="../__INC/functions.asp"-->

  <%
  Response.ContentType = "text/xml"
  
  Con_Open()

    RS_Open 1, "SELECT TOP 2500 tID, iDatumSparad " & _
               "FROM cms_TillbehorTitlar " & _
               "LEFT JOIN cms_Tillbehor ON cms_TillbehorTitlar.tTillbehorID = cms_Tillbehor.iID " & _ 
               "WHERE tID = iStandard_Titel AND iSynlig = 1 AND iDatumSparad <> 0 " & _
               "ORDER BY iDatumSparad DESC", False
    
      If rsDB(1).EOF Then
        any_Tillbehor = False
      Else
        any_Tillbehor = True
        list_Tillbehor = rsDB(1).GetRows
      End If
    
    RS_Close 1
  
  Con_Close()
  %>
  
  <url>
    <loc>http://www.n-forum.se/avdelning/tillbehor/default.asp</loc>
    <lastmod><% = Date %>T<% = Time %>+01:00</lastmod>
    <changefreq>always</changefreq>
  </url>
  
  <% If any_Tillbehor Then %>
    <% For zx = 0 To UBound(list_Tillbehor, 2) %>
      <url>
        <loc>http://www.n-forum.se/avdelning/tillbehor/tillbehor_visa_info.asp?e=<% = list_Tillbehor(0,zx) %></loc>
        <lastmod><% = FormatDateTime(list_Tillbehor(1,zx), vbShortDate) %>T<% = FormatDateTime(list_Tillbehor(1,zx), vbLongTime) %>+01:00</lastmod>
        <changefreq>daily</changefreq>
      </url>
    <% Next %>
  <% End if %>
</urlset>