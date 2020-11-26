<!--#INCLUDE FILE="../__INC/includes.asp"-->

<%
  sbTitel = Session.Value("trans_Titel")
  sbText  = Session.Value("trans_Text")
  sbLank  = Session.Value("trans_Lank")
  
  If Len(sbLank) < 1 Then Response.Redirect("/default.asp")
  
  Session.Value("trans_Titel") = ""
  Session.Value("trans_Text") = ""
  Session.Value("trans_Lank") = ""
%>

<%
  ' ## Globala variabler ##
  page_Title    = sbTitel & " - Meddelande"
  page_Header   = "Meddelande"
  page_WhereAmI = "&gt; Meddelande"
  page_SelMenu  = "--"
  page_Slide    = "--"
  Remove_Distas= True
%>

<!--#INCLUDE FILE="../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full">
      <h1><% = sbTitel %></h1>
      
      <div class="nf_msg_full">
        <p><% = sbText %></p>
        <p> <strong><a href="<% = sbLank %>">» Klicka för att fortsätta</a></strong> </p>
      </div>
      
    </div>
    
    <script type="text/javascript">
      setTimeout("location.href='<% = sbLank %>';", 5000);
    </script>
  </div>
  
<!--#INCLUDE FILE="../_page_bottom.asp"-->
<!--#INCLUDE FILE="../__INC/includes_end.asp"-->