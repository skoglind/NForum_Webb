<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%

  ' ## Globala variabler ##
  page_Title    = "Information - Medlem"
  page_Header   = "Information"
  page_WhereAmI = "&gt; <a href='registreradig.asp' title='Gå till &quot;Information&quot; ...'>Information</a> "
  page_SelMenu  = "user"
  page_Slide    = "info"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_unlogged.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1>Information</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      
      <div class="nf_msg">
        <p><strong>Information till dig som vill bli medlem på N-Forum.se.</strong></p>
        <p>Alla kan bli medlem på N-forum.se GRATIS under förutsättning att man följer de regler som vi har satt upp. Reglerna är till för att hålla en seriös nivå och för att göra detta till en trevlig webbsida för alla.</p><p>Alla som följer nedanstående <strong>regler</strong> är välkomna!</p>
      </div>
      
      <div class="nf_text">
        <!--#INCLUDE FILE="../../__INC/_rules.asp"-->
      </div>
      
      <div class="nf_msg nf_red">
        <p>Reglerna ovan kan ändras närhelst utan vidare meddelande till medlem.</p>
      </div>
      
      <div class="nf_msg nf_green">
        <p>- <em>"Somebody set up us the bomb."</em></p>
        <p>- <em>"All your base are belong to us!"</em></p>
      </div>
    </div>
    
    <div class="nf_datablock nf_size_onethird">

      <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
    
    </div>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->