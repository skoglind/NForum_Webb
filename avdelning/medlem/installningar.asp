<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  ' ## H�mta all data ##
  lID = CLng(CONST_USERID)
  sP  = GetQ("p", "ABC", 100)
  
  errCode = GetQ("fail","123",0)
  
  RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(lID), False
  
    If rsDB(1).EOF Then Response.Redirect("/")
    
    text_ID               = CLng(rsDB(1)("aID"))
    
    text_Namn             = sEncode(rsDB(1)("aNamn"))
    text_Plats            = sEncode(rsDB(1)("aPlats"))
    text_Hemsida          = sEncode(rsDB(1)("aHemsida"))
    text_MSN              = sEncode(rsDB(1)("aMSN"))
    text_ICQ              = sEncode(rsDB(1)("aICQ"))
    text_Signatur         = sEncode(rsDB(1)("aSignatur"))
    
    text_PM               = sEncode(rsDB(1)("aPM"))
    
    'text_PosterPerSida    = CLng(rsDB(1)("aID"))
    text_PMPerSida        = CLng(rsDB(1)("aIn_PM"))
    text_Position         = CLng(rsDB(1)("aIn_LoginPos"))
    text_Font             = CLng(rsDB(1)("aIn_Fontsize"))
    text_FontFam          = CLng(rsDB(1)("aIn_Fontfamily"))
    text_EpostVidPM       = rsDB(1)("aEpostPM")
    text_AktiveraPM       = rsDB(1)("aAktiveraPM")
    text_QuickList        = CLng(rsDB(1)("aIn_QuickList"))
    
    text_TradarPerSida    = CLng(rsDB(1)("aIn_Tradar"))
    text_InlaggPerSida    = CLng(rsDB(1)("aIn_Inlagg"))
    text_VisaAvatarer     = rsDB(1)("aIn_Avatarer")
    text_VisaSignaturer   = rsDB(1)("aIn_Signaturer")
    
    text_Epost            = sEncode(rsDB(1)("aEpost"))
    
    text_Avatar           = rsDB(1)("aAvatar")
  
  RS_Close 1
  
  Select Case LCase(sP)
    Case "meddelande"
      showPage = "message"
      showTitel = "Personligt meddelande"
    Case "sidan"
      showPage = "site"
      showTitel = "Sidinst�llningar"
    Case "forum"
      showPage = "forum"
      showTitel = "Foruminst�llningar"
    Case "epost"
      showPage = "email"
      showTitel = "Byt e-postadress"
    Case "losenord"
      showPage = "password"
      showTitel = "Byt l�senord"
    Case "avatar"
      showPage = "avatar"
      showTitel = "Ladda upp avatar"
    Case Else
      showPage = "personal"
      showTitel = "Personliga inst�llningar"
  End Select
  
  Call stop_Rec2Session("settings")
%>

<%
  ' ## Globala variabler ##
  page_Title    = showTitel & " - Inst�llningar - Medlem"
  page_Header   = "Inst�llningar"
  page_WhereAmI = "&gt; <a href='installningar.asp' title='G� till &quot;Inst�llningar&quot; ...'>Inst�llningar</a> "
  page_SelMenu  = "user"
  page_Slide    = "medlem"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_u.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1><% = showTitel %></h1>
    </div>
  
      <div class="nf_datablock nf_size_twothird">
      
        <% Select Case showPage %>
          <% Case "personal" ' PERSONLIGA INST�LLNINGAR %>
    
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>�ndringarna �r sparade!</p></div>
    
            <form method="POST" action="_action/savesettings.asp?p=personlig">        
              <div class="nf_form">
                <div class="nf_falt"><label>Namn:</label> <input type="text" name="namn" value="<% = text_Namn %>" maxlength=50 style="width: 436px;"></div>
                <div class="nf_falt"><label>Plats:</label> <input type="text" name="plats" value="<% = text_Plats %>" maxlength=50 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt"><label>Hemsida:</label> <input type="text" name="hemsida" value="<% = text_Hemsida %>" maxlength=255 style="width: 436px;"></div>
                <div class="nf_falt"><label>MSN:</label> <input type="text" name="MSN" value="<% = text_MSN %>" maxlength=255 style="width: 436px;"></div>
                <div class="nf_falt"><label>ICQ:</label> <input type="text" name="ICQ" value="<% = text_ICQ %>" maxlength=255 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt"><label>Signatur:</label> <input type="text" name="signatur" value="<% = text_Signatur %>" maxlength=255 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt nf_buttons">
                  <input type="submit" value="Spara">
                </div>
              </div>
            </form>
          <% Case "message" ' PERSONLIGT MEDDELANDE %>
    
            <div class="nf_msg"><p>Du kan h�r anv�nda dig av samma BB-Koder som forumet till�ter. Du m�ste begr�nsa din text till att best� av maximalt 20.000 tecken.</p></div>
            
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>�ndringarna �r sparade!</p></div>
    
            <form method="POST" action="_action/savesettings.asp?p=meddelande">        
              <div class="nf_form">

                <div class="nf_falt nf_buttonbar">
                  <input onclick="addText('aTextM','b');" type="button" value="B" style="width: 25px; font-weight: bold;">
                  <input onclick="addText('aTextM','i');" type="button" value="I" style="width: 25px; font-style: italic;">
                  <input onclick="addText('aTextM','u');" type="button" value="U" style="width: 25px; font-decoration: underline;">
                  <input onclick="addText('aTextM','s');" type="button" value="S" style="width: 25px; text-decoration: line-through;">
                  <div class="nf_buttonsplit">|</div>
                  <input onclick="addText('aTextM','url');" type="button" value="URL" style="width: 40px;">
                  <input onclick="addText('aTextM','img');" type="button" value="IMG" style="width: 40px;">
                  <div class="nf_buttonsplit">|</div>
                  <input onclick="addText('aTextM','spoiler');" type="button" value="Spoiler" style="width: 56px;">
                  <input onclick="addText('aTextM','indent');" type="button" value="Indenterad" style="width: 80px;">
                  <input onclick="addText('aTextM','code');" type="button" value="Monospace" style="width: 80px;">
                </div>
                
                <div class="nf_falt">
                  <textarea name="PM" id="aTextM" style="height: 500px; width: 576px" maxlength="20000" onkeyup="return ismaxlength(this)"><% = text_PM %></textarea>
                </div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt nf_buttons">
                  <input type="submit" value="Spara">
                </div>
              </div>
            </form>
          <% Case "site" ' SIDINST�LLNINGAR %>
    
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>�ndringarna �r sparade!</p></div>
    
            <form method="POST" action="_action/savesettings.asp?p=sidan">        
              <div class="nf_form">
                <div class="nf_falt"><label>PM per sida:</label>
                  <select name="PMSida" style="width: 441px;">
                    <% For zx = 10 To 50 Step 5 %>
                      <option value="<% = zx %>" <% If text_PMPerSida = zx Then Response.Write(" selected") %>><% = zx %> PM/Sida</option>
                    <% Next %>
                  </select>
                </div>
                <div class="nf_falt"><label>Position efter inloggning:</label>
                  <select name="Position" style="width: 441px;">
                    <option value=5 <% If text_Position = 5 Then Response.Write(" selected") %>> Min profil </option>
                    <!--<option value=4 <% If text_Position = 4 Then Response.Write(" selected") %>> Min blogg </option>-->
                    <option value=3 <% If text_Position = 3 Then Response.Write(" selected") %>> Startsidan </option>
                    <option value=2 <% If text_Position = 2 Then Response.Write(" selected") %>> Forumindex </option>
                    <option value=1 <% If text_Position = 1 Then Response.Write(" selected") %>> Forumet, Alla Forum </option>
                    <option value=0 <% If text_Position = 0 Then Response.Write(" selected") %>> �terg� till sidan du var p� </option>
                  </select>
                </div>
                <div class="nf_falt"><label>Snabblistning:</label> 
                  <select name="Quick" style="width: 441px;">
                    <option value=0 <% If text_QuickList = 0 Then Response.Write(" selected") %>> Inaktiverad </option>
                    <option value=1 <% If text_QuickList = 1 Then Response.Write(" selected") %>> Aktiverad (Listas med m�jlighet till �ndringar)  </option>
                    <option value=2 <% If text_QuickList = 2 Then Response.Write(" selected") %>> Aktiverad (Listas direkt) </option>
                  </select>
                </div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt"><label>Skicka e-post vid nytt PM:</label> <input type="checkbox" class="chk" name="EpostPM" value="YES" <% If text_EpostVidPM Then Response.Write(" checked") %>></div>
                <div class="nf_falt"><label>Aktivera PM:</label> <input type="checkbox" class="chk" name="AktPM" value="YES" <% If text_AktiveraPM Then Response.Write(" checked") %>></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt nf_buttons">
                  <input type="submit" value="Spara">
                </div>
              </div>
            </form>
          <% Case "forum" ' FORUMINST�LLNINGAR %>
    
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>�ndringarna �r sparade!</p></div>
    
            <form method="POST" action="_action/savesettings.asp?p=forum">        
              <div class="nf_form">
                <div class="nf_falt"><label>Tr�dar per sida:</label>
                  <select name="TradarSida" style="width: 441px;">
                    <% For zx = 10 To 50 Step 5 %>
                      <option value="<% = zx %>" <% If text_TradarPerSida = zx Then Response.Write(" selected") %>><% = zx %> Tr�dar/Sida</option>
                    <% Next %>
                  </select>
                </div>
                <div class="nf_falt"><label>Inl�gg per sida:</label>
                  <select name="InlaggSida" style="width: 441px;">
                    <% For zx = 10 To 40 Step 5 %>
                      <option value="<% = zx %>" <% If text_InlaggPerSida = zx Then Response.Write(" selected") %>><% = zx %> Inl�gg/Sida</option>
                    <% Next %>
                  </select>
                </div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt"><label>Visa avatarer:</label> <input type="checkbox" class="chk" value="YES" name="VisaAvatar" <% If text_VisaAvatarer Then Response.Write(" checked") %>></div>
                <div class="nf_falt"><label>Visa signaturer:</label> <input type="checkbox" class="chk" value="YES" name="VisaSignatur" <% If text_VisaSignaturer Then Response.Write(" checked") %>></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt nf_buttons">
                  <input type="submit" value="Spara">
                </div>
              </div>
            </form>
          <% Case "email" ' BYT E-POSTADRESS %>
    
            <div class="nf_msg"><p>Du m�ste ange ditt l�senord och sedan klicka p� verifieringsl�nken som skickas ut till din nya e-postadress f�r att kunna byta.</p></div>
            
            <% If errCode > 0 Then %>
              <div class="nf_msg nf_red">
                <% If errCode = 1 Then %><p><strong>E-postadressen byttes inte!</strong></p><p>Ditt l�senord st�mmer inte.</p><% End If %>
                <% If errCode = 2 Then %><p><strong>E-postadressen byttes inte!</strong></p><p>Ogiltigt e-postadress.</p><% End If %>
                <% If errCode = 3 Then %><p><strong>E-postadressen byttes inte!</strong></p><p>E-postadresserna st�mmer inte �verrens.</p><% End If %>
                <% If errCode = 4 Then %><p><strong>E-postadressen byttes inte!</strong></p><p>E-postadressern �r upptagen.</p><% End If %>
              </div>
            <% End If %>
            
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>Verifieringsmail skickat!</p></div>
    
            <form method="POST" action="_action/savesettings.asp?p=epost">        
              <div class="nf_form">
                <div class="nf_falt"><label>Nuvarande e-postadress:</label> <input type="text" disabled value="<% = text_Epost %>" style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt"><label>Ny e-postadress:</label> <input type="text" name="epost1" maxlength=255 style="width: 436px;"></div>
                <div class="nf_falt"><label>Bekr�fta ny e-postadress:</label> <input type="text" name="epost2" maxlength=255 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt"><label>L�senord:</label> <input type="password" name="passwd" maxlength=255 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt nf_buttons">
                  <input type="submit" value="Spara">
                </div>
              </div>
            </form>
          <% Case "password" ' BYT L�SENORD %>
    
            <div class="nf_msg"><p>Ditt gamla l�senord m�ste anges och det nya l�senordet m�ste best� av minst <strong>7</strong> (sju) tecken.</p></div>
            
            <% If errCode > 0 Then %>
              <div class="nf_msg nf_red">
                <% If errCode = 1 Then %><p<strong>L�senordet byttes inte!</strong></p><p>Ditt gamla l�senord st�mmer inte.</p><% End If %>
                <% If errCode = 2 Then %><p><strong>L�senordet byttes inte!</strong></p><p>Ogiltigt l�senord, det m�ste best� av minst 7 (sju) tecken.</p><% End If %>
                <% If errCode = 3 Then %><p><strong>L�senordet byttes inte!</strong></p><p>L�senorden st�mmer inte �verrens.</p><% End If %>
              </div>
            <% End If %>
            
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>Ditt l�senord �r nu bytt!</p></div>
    
            <form method="POST" action="_action/savesettings.asp?p=losenord">        
              <div class="nf_form">
                <div class="nf_falt"><label>Nytt l�senord:</label> <input type="password" name="pass1" maxlength=255 style="width: 436px;"></div>
                <div class="nf_falt"><label>Bekr�fta nytt l�senord:</label> <input type="password" name="pass2" maxlength=255 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt"><label>L�senord:</label> <input type="password" name="oldpass" maxlength=255 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt nf_buttons">
                  <input type="submit" value="Spara">
                </div>
              </div>
            </form>
          <% Case "avatar" ' LADDA UPP AVATAR %>
    
            <div class="nf_msg"><p>Bilden f�r inte vara st�rre �n <strong>50kB</strong> och kommer skalas om till <strong>100x100 pixlar</strong> om den inte redan har den storleken. Formatet p� bilden som laddas upp m�ste vara <strong>.png, .jpg, .bmp</strong> eller <strong>.gif</strong>.</p></div>
            
            <% If errCode > 0 Then %>
              <div class="nf_msg nf_red">
                <% If errCode = 1 Then %><p><strong>Avataren byttes inte!</strong></p><p>Du har inte valt n�gon fil.</p><% End If %>
                <% If errCode = 2 Then %><p><strong>Avataren byttes inte!</strong></p><p>Filen �r f�r stor.</p><% End If %>
                <% If errCode = 3 Then %><p><strong>Avataren byttes inte!</strong></p><p>Den var inte av formatet (jpg,bmp,png,gif).</p><% End If %>
              </div>
            <% End If %>
            
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>�tg�rden �r utf�rd</p></div>
    
            <form method="POST" action="_action/uploadavatar.asp" enctype="multipart/form-data">        
              <div class="nf_form">
                
                <% If text_Avatar Then %>
                  <div class="nf_falt" style="text-align: center;">
                    <img src="<% = config_Avatar %>u<% = Right("000000" & CONST_USERID, 6) %>.jpg" style="border: solid 1px #CCC; width: 100px; height: 100px;">
                  </div>
                
                  <div class="nf_separator"></div>
                <% End If %>
                
                <div class="nf_falt"><label>Bildfil:</label> <input type="file" name="avatar" size=68></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt nf_buttons">
                  <input type="submit" value="Ladda upp">
                  <input type="button" value="Ta bort" style="color: #A00;" onclick="if(confirm('Vill du radera din avatar?')){location.href='_action/savesettings.asp?p=deleteavatar';}" <% If Not text_Avatar Then Response.Write(" disabled") %>>
                </div>
              </div>
            </form>
        <% End Select %>

      </div>
      
      <div class="nf_datablock nf_size_onethird">
        <div class="nf_minibox">
          <h4>Inst�llningar</h4>
          <div class="nf_inside">
            <p><img src="<% = config_GFXLocation %>icons/menu/v_list.gif"> <a href="installningar.asp?p=personlig">Personliga inst�llningar</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/page.gif"> <a href="installningar.asp?p=meddelande">Personligt meddelande</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/avatar.gif"> <a href="installningar.asp?p=avatar">Avatar</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/v_list.gif"> <a href="installningar.asp?p=sidan">Sidinst�llningar</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/v_list.gif"> <a href="installningar.asp?p=forum">Foruminst�llningar</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/brev.gif"> <a href="installningar.asp?p=epost">Byt e-postadress</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/star.gif"> <a href="installningar.asp?p=losenord">Byt l�senord</a></p>
          </div>
        </div>
        
        <div class="nf_minibox nf_blue">
          <h4>Information</h4>
          <div class="nf_inside">
            <p>Komplettera g�rna dina uppgifter f�r att andra medlemmar ska veta vem du �r.</p>
            <p><strong>Gl�m inte att spara n�r du �ndrat n�got.</strong></p>
          </div>
        </div>
      </div>
      
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