<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  lID = GetQ("e","123",0)
%>

<%

  ' ## Globala variabler ##
  page_Title    = "Ny recension - Recensioner"
  page_Header   = "Recensionen är nu inskickad"
  page_WhereAmI = "&gt; <a href='registreradig.asp' title='Gå till &quot;Registrera dig&quot; ...'>Registrera dig</a> &gt; Registrerad!"
  page_SelMenu  = "texter"
  page_Slide    = "recensioner"
  
  page_description  = "Recensionen du skrivit och skickat in till N-Forum.se, Nintendo Forum, är nu sparad."
  page_keywords     = "sparad recension, "
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1>Recensionen är nu inskickad!</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">

      <div class="nf_msg">
        <p>Din recension är nu inskickad och kommer granskas av oss på N-Forum.se innan den eventuellt publiceras. Du kommer bli kontaktat per PM när den är publicerad!</p>
        <p><strong>Tack för din medverkan!</strong></p>
        <p><a href="spel_visa_info.asp?e=<% = lID %>">» Återgå till spelet</a></p>
      </div>
    </div>
    
    <div class="nf_datablock nf_size_onethird">

    </div>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->