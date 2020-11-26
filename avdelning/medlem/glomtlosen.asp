<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  errCode = GetQ("fail","123",0)
%>

<%

  ' ## Globala variabler ##
  page_Title    = "Glömt lösenordet? - Medlem"
  page_Header   = "Begär nytt lösenord"
  page_WhereAmI = "&gt; <a href='glomtlosen.asp' title='Gå till &quot;Glömt lösenordet?&quot; ...'>Glömt lösenordet? </a> "
  page_SelMenu  = "user"
  page_Slide    = "innanlogin"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_unlogged.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <form method="POST" action="_action/sendpassword.asp">
    
      <div class="nf_datablock nf_size_full">
        <h1>Glömt lösenordet?</h1>
      </div>
    
      <div class="nf_datablock nf_size_twothird">

        <div class="nf_msg">
          <p><strong>Har du glömt ditt lösenord?</strong></p>
          <p>Då kan du använda detta forumlär för skaffa ett nytt.</p>
        </div>
        
        <div id="savemess" class="nf_infomsg" style="display: none;"><p>Nytt lösenord är skickat!</p></div>
        
        <% If errCode > 0 Then %>
          <div class="nf_msg nf_red">
            <% If errCode = 1 Then %><p><strong>Inget lösenord skickades ut!</strong></p><p>E-postadressen finns inte i databasen eller så är användaren bannad.</p><% End If %>
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
            <p>När du angett din registrerade e-postadress kommer ett e-post gå ut till dig där du får en länk där du kan byta ditt lösenord.</p>
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