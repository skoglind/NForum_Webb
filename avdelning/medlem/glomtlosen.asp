<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  errCode = GetQ("fail","123",0)
%>

<%

  ' ## Globala variabler ##
  page_Title    = "Gl�mt l�senordet? - Medlem"
  page_Header   = "Beg�r nytt l�senord"
  page_WhereAmI = "&gt; <a href='glomtlosen.asp' title='G� till &quot;Gl�mt l�senordet?&quot; ...'>Gl�mt l�senordet? </a> "
  page_SelMenu  = "user"
  page_Slide    = "innanlogin"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_unlogged.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <form method="POST" action="_action/sendpassword.asp">
    
      <div class="nf_datablock nf_size_full">
        <h1>Gl�mt l�senordet?</h1>
      </div>
    
      <div class="nf_datablock nf_size_twothird">

        <div class="nf_msg">
          <p><strong>Har du gl�mt ditt l�senord?</strong></p>
          <p>D� kan du anv�nda detta foruml�r f�r skaffa ett nytt.</p>
        </div>
        
        <div id="savemess" class="nf_infomsg" style="display: none;"><p>Nytt l�senord �r skickat!</p></div>
        
        <% If errCode > 0 Then %>
          <div class="nf_msg nf_red">
            <% If errCode = 1 Then %><p><strong>Inget l�senord skickades ut!</strong></p><p>E-postadressen finns inte i databasen eller s� �r anv�ndaren bannad.</p><% End If %>
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
            <p>N�r du angett din registrerade e-postadress kommer ett e-post g� ut till dig d�r du f�r en l�nk d�r du kan byta ditt l�senord.</p>
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