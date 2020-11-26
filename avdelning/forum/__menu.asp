<div class="minimenu">
  <ul>
    <li <% If page_Slide = "index"   Then Response.Write(" class='c'") %>> <a href="default.asp" title="">Forumindex</a> </li>
    <li <% If page_Slide = "allfora" Then Response.Write(" class='c'") %>> <a href="forum.asp" title="">Alla forum</a> </li>
    <li <% If page_Slide = "latest"  Then Response.Write(" class='c'") %>> <a href="nyainlagg.asp" title="">Nya inlägg</a> </li>
    <li class="last<% If page_Slide = "search"  Then Response.Write(" c") %>"> <a href="sokforum.asp" title="">Sök i forumet</a> </li>
  </ul> 
</div>