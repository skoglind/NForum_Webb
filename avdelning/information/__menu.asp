<div class="minimenu">
  <ul>
    <li <% If page_Slide = "cookie"       Then Response.Write(" class='c'") %>> <a href="cookies.asp" title="">Cookies</a> </li>
    <li <% If page_Slide = "faq"    Then Response.Write(" class='c'") %>> <a href="faq.asp" title="">F.A.Q.</a> </li>
    <li <% If page_Slide = "link"  Then Response.Write(" class='c'") %>> <a href="linkback.asp" title="">Länka hit</a> </li>
    <li class="last<% If page_Slide = "info"  Then Response.Write(" c") %>"> <a href="information.asp" title="">Information</a> </li>
  </ul> 
</div>