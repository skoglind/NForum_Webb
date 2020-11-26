<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  errCode = GetQ("fail","123",0)
%>

<%

  ' ## Globala variabler ##
  page_Title    = "Nytt aktiveringsbrev - Medlem"
  page_Header   = "Begär nytt aktiveringsbrev"
  page_WhereAmI = "&gt; <a href='omaktivera.asp' title='Gå till &quot;Nytt aktiveringsbrev&quot; ...'>Nytt aktiveringsbrev </a> "
  page_SelMenu  = "user"
  page_Slide    = "innanlogin"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_unlogged.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <form method="POST" action="_action/sendaktivering.asp">
    
      <div class="nf_datablock nf_size_full">
        <h1>Nytt aktiveringsbrev</h1>
      </div>
    
      <div class="nf_datablock nf_size_twothird">
        
        <div class="nf_msg">
          <p><strong>Inte fått något aktiveringsbrev?</strong></p>
          <p>Då kan du använda detta forumlär för få ett nytt utskickat.</p>
        </div>
        
        <div id="savemess" class="nf_infomsg" style="display: none;"><p>Nytt aktiveringsbrev är skickat!</p></div>
        
        <% If errCode > 0 Then %>
          <div class="nf_msg nf_red">
            <% If errCode = 1 Then %><p><strong>Inget aktiveringsbrev skickades ut!</strong></p><p>E-postadressen finns inte i databasen eller så är användaren redan aktiverad.</p><% End If %>
          </div>
        <% End If %>
        
        <div class="nf_form">
        
          <div class="nf_falt"><label>E-postadress:</label> <input type="text" name="epost" maxlength=255 style="width: 436px;"></div>

          <div class="nf_falt nf_buttons">
            <input type="submit" value="Skicka">
          </div>
          
        </div>
      
      </div>
      
      <div class="nf_datablock nf_size_onethird">

        <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
      
        <div class="nf_minibox nf_blue">
          <h4>Information</h4>
          <div class="nf_inside">
            <p>När du angett din registrerade e-postadress kommer ett e-post med en ny aktiveringslänk att skickas till dig.</p>
          </div>
        </div>

      </div>
    
    </form>
  
    <% If Session.Value("form_saved") Then %>
      <script type="text/javascript">
        show("savemess");
        setTimeout("hide('savemess');", 2500);
      </script>
      <% Session.Value("form_saved") = False %>
    <% End If %>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->