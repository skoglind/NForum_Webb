<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%

  errCode = GetQ("fail","123",0)

  Randomize
  Do
    lVal1 = CLng(Rnd*9) +1 
    lVal2 = CLng(Rnd*9) +1
  Loop Until lVal1 + lVal2 < 13
  
  Session.Value("svaret") = lVal1 + lVal2
  
  If Session.Value("record_reg") Then
    text_aNamn             = sEncode(Session.Value("reg_anvnamn"))
    text_Epost1            = sEncode(Session.Value("reg_epost1"))
    text_Epost2            = sEncode(Session.Value("reg_epost2"))
    
    Session.Value("record_reg") = False
  End If
%>

<%

  ' ## Globala variabler ##
  page_Title    = "Registrera dig - Medlem"
  page_Header   = "Bli medlem"
  page_WhereAmI = "&gt; <a href='registreradig.asp' title='Gå till &quot;Registrera dig&quot; ...'>Registrera dig</a> "
  page_SelMenu  = "user"
  page_Slide    = "innanlogin"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_unlogged.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <form method="POST" action="_action/doregister.asp">
    
      <div class="nf_datablock nf_size_full">
        <h1>Registrera dig</h1>
      </div>
    
      <div class="nf_datablock nf_size_twothird">
        
        <div class="nf_msg">
          <p><strong>Vill du bli medlem på N-Forum.se?</strong></p>
          <p>Då fyller du i forumläret nedan, att bli medlem på N-Forum.se är helt <strong>GRATIS</strong>.</p>
        </div>
        
        <% If errCode > 0 Then %>
          <div class="nf_msg nf_red">
            <% If errCode = 1 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Ogiltigt användarnamn.</p><% End If %>
            <% If errCode = 2 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Användarnamnet är upptaget.</p><% End If %>
            <% If errCode = 3 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Ogiltig e-postadress.</p><% End If %>
            <% If errCode = 4 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Det finns redan en användare registrerad med den angivna e-postadressen.</p><% End If %>
            <% If errCode = 5 Then %><p><strong>Registreringen misslyckades!</strong></p><p>E-postadresserna stämde inte överrens.</p><% End If %>
            <% If errCode = 6 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Du har inte godkänt reglerna.</p><% End If %>
            <% If errCode = 7 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Du fyllde i fel svar på frågan.</p><% End If %>
            <% If errCode = 666 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Nyregistreringar är nerstängda av systemadministratören, om detta kommer bestå under en längre tid kommer vi med mer information snart.</p><% End If %>
          </div>
        <% End If %>
        
        <div class="nf_form">
        
          <div class="nf_falt"><label>Användarnamn:</label> <input type="text" name="anvnamn" value="<% = text_aNamn %>" maxlength=30 style="width: 436px;"></div>
          
          <div class="nf_separator"></div>
          
          <div class="nf_falt"><label>E-postadress</label> <input type="text" name="epost1" value="<% = text_Epost1 %>" maxlength=255 style="width: 436px;"></div>
          <div class="nf_falt"><label>Bekräfta e-postadress:</label> <input type="text" name="epost2" value="<% = text_Epost2 %>" maxlength=255 style="width: 436px;"></div>
          
          <div class="nf_separator"></div>
          
          <div class="nf_falt"><label>Fyll i svaret på: <span style="color: #080; font: bold 16px Arial;"><% = Server.URLEncode(lVal1) %> + <% = Server.URLEncode(lVal2) %></span> ?</label> <input type="text" name="math" value="<% = text_Titel %>" maxlength=10 style="width: 436px;"></div>
          
          <div class="nf_separator"></div>
          
          <div class="nf_falt"><input type="checkbox" name="avtal" value="YES"> <span style="float: left; font: 12px Arial; margin: 4px 0 0 0;">Ja, jag godkänner <a href="information.asp" target="_blank">reglerna</a> och lovar att följa dem!</span></div>
          
          <div class="nf_falt nf_buttons">
            <input type="submit" value="Registrera mig" style="width: 120px;">
          </div>
          
        </div>
      
      </div>
      
      <div class="nf_datablock nf_size_onethird">

        <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
      
        <div class="nf_minibox nf_blue">
          <h4>Information</h4>
          <div class="nf_inside">
            <p>Alla fält är obligatoriska att fylla i, när du trycker på [<em>Registrera mig</em>] kommer du få ett e-postbrev för att bekräfta din e-postadress.</p>
            <p>När du väl är medlem har du möjlighet att komplettera dina uppgifter.</p>
          </div>
        </div>
        
        <div class="nf_minibox nf_blue">
          <h4>Fördelar som medlem</h4>
          <div class="nf_inside">
            <p> - Egen profil </p>
            <p> - Lista dina spel </p>
            <p> - Ändra sidinställningar </p>
            <p> - Skriva inlägg i forumet </p>
            <p> - Betygsätta spel </p>
            <p> - Se högupplöst boxart </p>
            <p> - Lägga upp köp & sälj annonser </p>
            <p> - Kommentera annonser </p>
            <p> - Ladda upp bilder i din profil </p>
            <p> - Skriva recensioner </p>
            <p> - Skriva artiklar </p>
          </div>
        </div>
        
        <div class="nf_minibox nf_red">
          <h4>Observera</h4>
          <div class="nf_inside">
            <p>E-postadresser från sidor som <em>mailinator.com</em> och liknande är spärrade i våran databas.</p>
          </div>
        </div>
          
      </div>
    
    </form>  
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->