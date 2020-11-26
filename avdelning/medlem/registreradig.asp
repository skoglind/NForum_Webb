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
  page_WhereAmI = "&gt; <a href='registreradig.asp' title='G� till &quot;Registrera dig&quot; ...'>Registrera dig</a> "
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
          <p><strong>Vill du bli medlem p� N-Forum.se?</strong></p>
          <p>D� fyller du i foruml�ret nedan, att bli medlem p� N-Forum.se �r helt <strong>GRATIS</strong>.</p>
        </div>
        
        <% If errCode > 0 Then %>
          <div class="nf_msg nf_red">
            <% If errCode = 1 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Ogiltigt anv�ndarnamn.</p><% End If %>
            <% If errCode = 2 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Anv�ndarnamnet �r upptaget.</p><% End If %>
            <% If errCode = 3 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Ogiltig e-postadress.</p><% End If %>
            <% If errCode = 4 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Det finns redan en anv�ndare registrerad med den angivna e-postadressen.</p><% End If %>
            <% If errCode = 5 Then %><p><strong>Registreringen misslyckades!</strong></p><p>E-postadresserna st�mde inte �verrens.</p><% End If %>
            <% If errCode = 6 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Du har inte godk�nt reglerna.</p><% End If %>
            <% If errCode = 7 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Du fyllde i fel svar p� fr�gan.</p><% End If %>
            <% If errCode = 666 Then %><p><strong>Registreringen misslyckades!</strong></p><p>Nyregistreringar �r nerst�ngda av systemadministrat�ren, om detta kommer best� under en l�ngre tid kommer vi med mer information snart.</p><% End If %>
          </div>
        <% End If %>
        
        <div class="nf_form">
        
          <div class="nf_falt"><label>Anv�ndarnamn:</label> <input type="text" name="anvnamn" value="<% = text_aNamn %>" maxlength=30 style="width: 436px;"></div>
          
          <div class="nf_separator"></div>
          
          <div class="nf_falt"><label>E-postadress</label> <input type="text" name="epost1" value="<% = text_Epost1 %>" maxlength=255 style="width: 436px;"></div>
          <div class="nf_falt"><label>Bekr�fta e-postadress:</label> <input type="text" name="epost2" value="<% = text_Epost2 %>" maxlength=255 style="width: 436px;"></div>
          
          <div class="nf_separator"></div>
          
          <div class="nf_falt"><label>Fyll i svaret p�: <span style="color: #080; font: bold 16px Arial;"><% = Server.URLEncode(lVal1) %> + <% = Server.URLEncode(lVal2) %></span> ?</label> <input type="text" name="math" value="<% = text_Titel %>" maxlength=10 style="width: 436px;"></div>
          
          <div class="nf_separator"></div>
          
          <div class="nf_falt"><input type="checkbox" name="avtal" value="YES"> <span style="float: left; font: 12px Arial; margin: 4px 0 0 0;">Ja, jag godk�nner <a href="information.asp" target="_blank">reglerna</a> och lovar att f�lja dem!</span></div>
          
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
            <p>Alla f�lt �r obligatoriska att fylla i, n�r du trycker p� [<em>Registrera mig</em>] kommer du f� ett e-postbrev f�r att bekr�fta din e-postadress.</p>
            <p>N�r du v�l �r medlem har du m�jlighet att komplettera dina uppgifter.</p>
          </div>
        </div>
        
        <div class="nf_minibox nf_blue">
          <h4>F�rdelar som medlem</h4>
          <div class="nf_inside">
            <p> - Egen profil </p>
            <p> - Lista dina spel </p>
            <p> - �ndra sidinst�llningar </p>
            <p> - Skriva inl�gg i forumet </p>
            <p> - Betygs�tta spel </p>
            <p> - Se h�guppl�st boxart </p>
            <p> - L�gga upp k�p & s�lj annonser </p>
            <p> - Kommentera annonser </p>
            <p> - Ladda upp bilder i din profil </p>
            <p> - Skriva recensioner </p>
            <p> - Skriva artiklar </p>
          </div>
        </div>
        
        <div class="nf_minibox nf_red">
          <h4>Observera</h4>
          <div class="nf_inside">
            <p>E-postadresser fr�n sidor som <em>mailinator.com</em> och liknande �r sp�rrade i v�ran databas.</p>
          </div>
        </div>
          
      </div>
    
    </form>  
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->