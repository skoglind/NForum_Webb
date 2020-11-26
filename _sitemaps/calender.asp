<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">

  <%
  Response.ContentType = "text/xml"
  %>
  
  <url>
    <loc>http://www.n-forum.se/nintendo/kalender.asp</loc>
    <lastmod><% = Date %>T<% = Time %>+01:00</lastmod>
    <changefreq>always</changefreq>
  </url>
  
  <% For zx = -50 To 200 %>
    <% myDate = DateAdd("d", -zx, Date) %>
    <url>
      <loc>http://www.n-forum.se/nintendo/kalender.asp?d=<% = myDate %></loc>
      <lastmod><% = FormatDateTime(myDate, vbShortDate) %>T<% = FormatDateTime(myDate, vbLongTime) %>+01:00</lastmod>
      <changefreq>daily</changefreq>
    </url>
  <% Next %>
  
</urlset>