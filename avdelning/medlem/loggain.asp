<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%

  errCode = GetQ("fail","123",0)

  pBack = Session.Value("login_PB")
  Session.Value("login_PB") = ""
%>

<%

  ' ## Globala variabler ##
  page_Title    = "Logga in - Medlem"
  page_Header   = "Logga in"
  page_WhereAmI = "&gt; <a href='loggain.asp' title='G� till &quot;Logga in&quot; ...'>Logga in</a> "
  page_SelMenu  = "user"
  page_Slide    = "innanlogin"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_unlogged.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <form method="POST" action="/_action/do_login.asp">
    
      <div class="nf_datablock nf_size_full">
        <h1>Logga in</h1>
      </div>
    
      <div class="nf_datablock nf_size_twothird">

        <div class="nf_msg">
          <p><strong>V�lkommen!</strong></p>
          <p>Logga in f�r att komma �t din profil och dina personliga inst�llningar.</p>
        </div>
        
        <% If errCode > 0 Then %>
          <div class="nf_msg nf_red">
            <% If errCode = 1 Then %><p><strong>Inloggningen misslyckades!</strong></p><p>Felaktigt anv�ndarnamn och/eller l�senord.</p><% End If %>
            <% If errCode = 2 Then %><p><strong>Inloggningen misslyckades!</strong></p><p>Anv�ndaren �r bannad, kontakta administrat�ren f�r mer information.</p><% End If %>
            <% If errCode = 3 Then %><p><strong>Inloggningen misslyckades!</strong></p><p>Anv�ndaren �r inte aktiverad, ett e-postbrev ska ha skickats ut vid registreringstillf�llet. Om s� inte skett v�lj d� att <a href="omaktivera.asp">skicka ett nytt e-postbrev</a>.</p><% End If %>
          </div>
        <% End If %>
        
        <div class="nf_form">
        
          <div class="nf_falt"><label>Anv�ndarnamn:</label> <input type="text" name="r" maxlength=30 style="width: 436px;"></div>
          <div class="nf_falt"><label>L�senord</label> <input type="password" name="g" maxlength=255 style="width: 436px;"></div>

          <div class="nf_separator"></div>
          
          <div class="nf_falt"><input type="checkbox" name="s" value="YES"> <span style="float: left; font: 12px Arial; margin: 4px 0 0 0;">Kom ih�g mig!</span></div>
          
          <div class="nf_falt nf_buttons">
            <input type="hidden" name="postback" value="<% = sEncode(pBack) %>">
            <input type="submit" value="Logga in">
          </div>
          
        </div>
      
      </div>
      
      <div class="nf_datablock nf_size_onethird">

        <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
        
        <div class="nf_minibox nf_blue">
          <h4>Gl�mt ditt l�senord?</h4>
          <div class="nf_inside">
            <p> Du kan beg�ra ett nytt l�senord via f�ljande foruml�r... </p>
            <p> <strong>� <a href="/avdelning/medlem/glomtlosen.asp">Gl�mt l�senordet?</a></strong> </p>
            <p> Du kommmer d� f� ett nytt l�senord utskickat till din registrerade e-postadress.</p>
          </div>
        </div>
      
      </div>
    
    </form>  
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->