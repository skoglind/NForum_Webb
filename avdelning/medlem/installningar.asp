<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  ' ## Hämta all data ##
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
      showTitel = "Sidinställningar"
    Case "forum"
      showPage = "forum"
      showTitel = "Foruminställningar"
    Case "epost"
      showPage = "email"
      showTitel = "Byt e-postadress"
    Case "losenord"
      showPage = "password"
      showTitel = "Byt lösenord"
    Case "avatar"
      showPage = "avatar"
      showTitel = "Ladda upp avatar"
    Case Else
      showPage = "personal"
      showTitel = "Personliga inställningar"
  End Select
  
  Call stop_Rec2Session("settings")
%>

<%
  ' ## Globala variabler ##
  page_Title    = showTitel & " - Inställningar - Medlem"
  page_Header   = "Inställningar"
  page_WhereAmI = "&gt; <a href='installningar.asp' title='Gå till &quot;Inställningar&quot; ...'>Inställningar</a> "
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
          <% Case "personal" ' PERSONLIGA INSTÄLLNINGAR %>
    
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>Ändringarna är sparade!</p></div>
    
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
    
            <div class="nf_msg"><p>Du kan här använda dig av samma BB-Koder som forumet tillåter. Du måste begränsa din text till att bestå av maximalt 20.000 tecken.</p></div>
            
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>Ändringarna är sparade!</p></div>
    
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
          <% Case "site" ' SIDINSTÄLLNINGAR %>
    
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>Ändringarna är sparade!</p></div>
    
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
                    <option value=0 <% If text_Position = 0 Then Response.Write(" selected") %>> Återgå till sidan du var på </option>
                  </select>
                </div>
                <div class="nf_falt"><label>Snabblistning:</label> 
                  <select name="Quick" style="width: 441px;">
                    <option value=0 <% If text_QuickList = 0 Then Response.Write(" selected") %>> Inaktiverad </option>
                    <option value=1 <% If text_QuickList = 1 Then Response.Write(" selected") %>> Aktiverad (Listas med möjlighet till ändringar)  </option>
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
          <% Case "forum" ' FORUMINSTÄLLNINGAR %>
    
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>Ändringarna är sparade!</p></div>
    
            <form method="POST" action="_action/savesettings.asp?p=forum">        
              <div class="nf_form">
                <div class="nf_falt"><label>Trådar per sida:</label>
                  <select name="TradarSida" style="width: 441px;">
                    <% For zx = 10 To 50 Step 5 %>
                      <option value="<% = zx %>" <% If text_TradarPerSida = zx Then Response.Write(" selected") %>><% = zx %> Trådar/Sida</option>
                    <% Next %>
                  </select>
                </div>
                <div class="nf_falt"><label>Inlägg per sida:</label>
                  <select name="InlaggSida" style="width: 441px;">
                    <% For zx = 10 To 40 Step 5 %>
                      <option value="<% = zx %>" <% If text_InlaggPerSida = zx Then Response.Write(" selected") %>><% = zx %> Inlägg/Sida</option>
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
    
            <div class="nf_msg"><p>Du måste ange ditt lösenord och sedan klicka på verifieringslänken som skickas ut till din nya e-postadress för att kunna byta.</p></div>
            
            <% If errCode > 0 Then %>
              <div class="nf_msg nf_red">
                <% If errCode = 1 Then %><p><strong>E-postadressen byttes inte!</strong></p><p>Ditt lösenord stämmer inte.</p><% End If %>
                <% If errCode = 2 Then %><p><strong>E-postadressen byttes inte!</strong></p><p>Ogiltigt e-postadress.</p><% End If %>
                <% If errCode = 3 Then %><p><strong>E-postadressen byttes inte!</strong></p><p>E-postadresserna stämmer inte överrens.</p><% End If %>
                <% If errCode = 4 Then %><p><strong>E-postadressen byttes inte!</strong></p><p>E-postadressern är upptagen.</p><% End If %>
              </div>
            <% End If %>
            
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>Verifieringsmail skickat!</p></div>
    
            <form method="POST" action="_action/savesettings.asp?p=epost">        
              <div class="nf_form">
                <div class="nf_falt"><label>Nuvarande e-postadress:</label> <input type="text" disabled value="<% = text_Epost %>" style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt"><label>Ny e-postadress:</label> <input type="text" name="epost1" maxlength=255 style="width: 436px;"></div>
                <div class="nf_falt"><label>Bekräfta ny e-postadress:</label> <input type="text" name="epost2" maxlength=255 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt"><label>Lösenord:</label> <input type="password" name="passwd" maxlength=255 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt nf_buttons">
                  <input type="submit" value="Spara">
                </div>
              </div>
            </form>
          <% Case "password" ' BYT LÖSENORD %>
    
            <div class="nf_msg"><p>Ditt gamla lösenord måste anges och det nya lösenordet måste bestå av minst <strong>7</strong> (sju) tecken.</p></div>
            
            <% If errCode > 0 Then %>
              <div class="nf_msg nf_red">
                <% If errCode = 1 Then %><p<strong>Lösenordet byttes inte!</strong></p><p>Ditt gamla lösenord stämmer inte.</p><% End If %>
                <% If errCode = 2 Then %><p><strong>Lösenordet byttes inte!</strong></p><p>Ogiltigt lösenord, det måste bestå av minst 7 (sju) tecken.</p><% End If %>
                <% If errCode = 3 Then %><p><strong>Lösenordet byttes inte!</strong></p><p>Lösenorden stämmer inte överrens.</p><% End If %>
              </div>
            <% End If %>
            
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>Ditt lösenord är nu bytt!</p></div>
    
            <form method="POST" action="_action/savesettings.asp?p=losenord">        
              <div class="nf_form">
                <div class="nf_falt"><label>Nytt lösenord:</label> <input type="password" name="pass1" maxlength=255 style="width: 436px;"></div>
                <div class="nf_falt"><label>Bekräfta nytt lösenord:</label> <input type="password" name="pass2" maxlength=255 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt"><label>Lösenord:</label> <input type="password" name="oldpass" maxlength=255 style="width: 436px;"></div>
                
                <div class="nf_separator"></div>
                
                <div class="nf_falt nf_buttons">
                  <input type="submit" value="Spara">
                </div>
              </div>
            </form>
          <% Case "avatar" ' LADDA UPP AVATAR %>
    
            <div class="nf_msg"><p>Bilden får inte vara större än <strong>50kB</strong> och kommer skalas om till <strong>100x100 pixlar</strong> om den inte redan har den storleken. Formatet på bilden som laddas upp måste vara <strong>.png, .jpg, .bmp</strong> eller <strong>.gif</strong>.</p></div>
            
            <% If errCode > 0 Then %>
              <div class="nf_msg nf_red">
                <% If errCode = 1 Then %><p><strong>Avataren byttes inte!</strong></p><p>Du har inte valt någon fil.</p><% End If %>
                <% If errCode = 2 Then %><p><strong>Avataren byttes inte!</strong></p><p>Filen är för stor.</p><% End If %>
                <% If errCode = 3 Then %><p><strong>Avataren byttes inte!</strong></p><p>Den var inte av formatet (jpg,bmp,png,gif).</p><% End If %>
              </div>
            <% End If %>
            
            <div id="savemess" class="nf_infomsg" style="display: none;"><p>Åtgärden är utförd</p></div>
    
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
          <h4>Inställningar</h4>
          <div class="nf_inside">
            <p><img src="<% = config_GFXLocation %>icons/menu/v_list.gif"> <a href="installningar.asp?p=personlig">Personliga inställningar</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/page.gif"> <a href="installningar.asp?p=meddelande">Personligt meddelande</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/avatar.gif"> <a href="installningar.asp?p=avatar">Avatar</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/v_list.gif"> <a href="installningar.asp?p=sidan">Sidinställningar</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/v_list.gif"> <a href="installningar.asp?p=forum">Foruminställningar</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/brev.gif"> <a href="installningar.asp?p=epost">Byt e-postadress</a></p>
            <p><img src="<% = config_GFXLocation %>icons/menu/star.gif"> <a href="installningar.asp?p=losenord">Byt lösenord</a></p>
          </div>
        </div>
        
        <div class="nf_minibox nf_blue">
          <h4>Information</h4>
          <div class="nf_inside">
            <p>Komplettera gärna dina uppgifter för att andra medlemmar ska veta vem du är.</p>
            <p><strong>Glöm inte att spara när du ändrat något.</strong></p>
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