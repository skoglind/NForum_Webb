<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  errCode = GetQ("fail","123",0)
  
  If Session.Value("record_chpass") Then
    text_aNamn             = sEncode(Session.Value("chpass_anamn"))
    text_Nyckel            = sEncode(Session.Value("chpass_nyckel"))
    
    Session.Value("record_chpass") = False
  Else
    text_aNamn             = sEncode(GetQ("ua","ABC",0))
    text_Nyckel            = sEncode(GetQ("x","ABC",10))
  End If
  
  'Call stop_Rec2Session("chpass")
%>

<%

  ' ## Globala variabler ##
  page_Title    = "Nytt lösenord! - Medlem"
  page_Header   = "Byt lösenord"
  page_WhereAmI = "&gt; <a href='glomtlosen.asp' title='Gå till &quot;Glömt lösenordet?&quot; ...'>Glömt lösenordet? </a> &gt; <a href='nyttlosenord.asp' title='Gå till &quot;Byt lösenord!&quot; ...'>Byt lösenord!</a>"
  page_SelMenu  = "user"
  page_Slide    = "innanlogin"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_unlogged.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
   
    <form method="POST" action="_action/newpassword.asp">
   
      <div class="nf_datablock nf_size_full">
        <h1>Nytt lösenord</h1>
      </div>
     
      <div class="nf_datablock nf_size_twothird">

        <div class="nf_msg">
          <p>Använd detta forumlär för att byta ditt lösenord. Nyckeln som ska anges får du med e-postbrevet som går ut när du fyller i din e-postadress i "Glömt lösenordet?" forumläret. Det nya lösenordet måste bestå av minst 7 (sju) tecken.</p>
        </div>
        
        <div id="savemess" class="nf_infomsg" style="display: none;"><p>Ditt lösenord är nu bytt!</p></div>
        
        <% If errCode > 0 Then %>
          <div class="nf_msg nf_red">
            <% If errCode = 1 Then %><p><strong>Lösenordet byttes inte!</strong></p><p>Användarnamnet och/eller nyckeln är felaktig.</p><% End If %>
            <% If errCode = 2 Then %><p><strong>Lösenordet byttes inte!</strong></p><p>Du har inte angett något nytt lösenord.</p><% End If %>
            <% If errCode = 3 Then %><p><strong>Lösenordet byttes inte!</strong></p><p>Ogiltigt lösenord, det måste bestå av minst 7 (sju) tecken.</p><% End If %>
            <% If errCode = 4 Then %><p><strong>Lösenordet byttes inte!</strong></p><p>Lösenorden stämmer inte överrens.</p><% End If %>
          </div>
        <% End If %>
        
        <div class="nf_form">
        
          <div class="nf_falt"><label>Användarnamn:</label> <input type="text" name="anamn" maxlength=30 style="width: 405px;" value="<% = text_aNamn %>"></div>
          <div class="nf_falt"><label>Nyckel:</label> <input type="text" name="nyckel" maxlength=10 style="width: 405px;" value="<% = text_Nyckel %>"></div>
          
          <div class="nf_separator"></div>
          
          <div class="nf_falt"><label>Nytt lösenord:</label> <input type="password" name="passwd1" maxlength=255 style="width: 405px;"></div>
          <div class="nf_falt"><label>Bekräfta nytt lösenord:</label> <input type="password" name="passwd2" maxlength=255 style="width: 405px;"></div>

          <div class="nf_falt nf_buttons">
            <input type="submit" value="Byt lösenord" style="width: 120px;">
          </div>
          
        </div>
      
      </div>
      
      <div class="nf_datablock nf_size_onethird">

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