<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect("default.asp") %>
<% If config_LockDown_Forum Then Response.Redirect("default.asp") %>

<%
  ' ## H�mta all data ##
  lID  = GetQ("e", "123", 0)
  lAID = GetQ("a", "123", 0)
  lCID = GetQ("c", "123", 0)
  
  lFora = GetQ("f", "123", 0)
  sURL  = GetQ("url", "ABC", 255)
  
  errCode = GetQ("fail","123",0)
  
  If Session.Value("record_post") Then  ' ## POSTBACK
    text_Amne             = sEncode(Session.Value("post_fAmne"))
    text_TextM            = sEncode(Session.Value("post_fMsg"))
    
    text_AutoSmil         = Session.Value("post_fAutoSmil")
    text_AutoUrl          = Session.Value("post_fAutoUrl")
    
    m_lID                 = Session.Value("post_tradID")
    m_lIDSvar             = Session.Value("post_tradID_Svar")

    ' ## KOLLA TR�DTYP
      If m_lID = 0 And m_lIDSvar = 0  Then lPostStat = 0    ' En helt ny tr�d
      If m_lID = 0 And m_lIDSvar <> 0 Then lPostStat = 1    ' Ett nytt inl�gg
      If m_lID <> 0 And m_lIDSvar = 0 Then lPostStat = 2    ' En tr�d/inl�gg redigeras
      
      If lPostStat = 2 Then
        lTradTyp    = IsMainThread(m_lID)
        If lTradTyp Then lPostStat = 2 Else lPostStat = 3
      End If
    ' ##
    
    Select Case lPostStat
      Case 0  ' En helt ny tr�d
        sTitle                = "Ny tr�d"
        text_MainThread       = True
        text_IsTrad           = True
        text_ID               = m_lID
        text_Forum            = Session.Value("post_kategori")
        text_Locked           = Session.Value("post_fLocked")
        text_Klistrad         = Session.Value("post_fKlistrad")
        text_Dold             = Session.Value("post_fDold")
        
        'GetRightsForum forumID ' Fixa
        'If Not sec_Trad_Skapa Then Response.Redirect("forum.asp?e=" & forumID)
      Case 1  ' Ett nytt inl�gg
        sTitle                = "Nytt inl�gg"
        text_MainThread       = False
        text_IsTrad           = False
        text_ID               = m_lIDSvar
        
        GetRights text_ID ' H�mta fram r�ttigheterna
        If Not sec_Inlagg_Skapa Then Response.Redirect("trad.asp?e=" & text_ID)
      Case 2  ' En tr�d redigeras
        sTitle                = "Redigera tr�d"
        text_MainThread       = True
        text_IsTrad           = True
        text_ID               = m_lID
        text_Forum            = Session.Value("post_kategori")
        text_Locked           = Session.Value("post_fLocked")
        text_Klistrad         = Session.Value("post_fKlistrad")
        text_Dold             = Session.Value("post_fDold")
        
        GetRights text_ID ' H�mta fram r�ttigheterna
        If Not sec_Trad_Hantera Then Response.Redirect("trad.asp?e=" & text_ID)
      Case 3  ' Ett inl�gg redigeras
        sTitle                = "Redigera inl�gg"
        text_MainThread       = False
        text_IsTrad           = True
        text_ID               = m_lID
        
        GetRights text_ID ' H�mta fram r�ttigheterna
        GetUserStatsFromPost text_ID
        
        If Not sec_Hantera(lPost_UserAdmin, lPost_UserID) Then Response.Redirect("trad.asp?e=" & lPost_TradID)
    End Select
  ElseIf lCID > 0 Then  ' ## NYTT INL�GG med CITAT
    RS_Open 1, "SELECT *, fsBB_Anv.aAnvNamn AS aAnvNamn " & _
               "FROM fsBB_Tradar " & _
               "LEFT JOIN fsBB_Anv ON tAnv_Skapad = fsBB_Anv.aID " & _
               "LEFT JOIN fsBB_Forum ON tForum = fsBB_Forum.fID " & _
               "WHERE tID = " & CLng(lCID), False
    
      If Not rsDB(1).EOF Then
        sTitle = "Nytt inl�gg"
    
        If rsDB(1)("tStatus_Undertrad") = 0 Then
          text_ID               = rsDB(1)("tID")
        Else
          text_ID               = rsDB(1)("tStatus_Undertrad")
        End If
    
        If InStr(rsDB(1)("fSec_View"), ";" & CLng(CONST_TITEL) & ";") Or rsDB(1)("fSec_View") = "0" Then 
          text_Amne             = CutText("Sv: " & Trim(sEncode(rsDB(1)("tAmne"))), 100)
          text_TextM            = "[quote][i]Ursprungligen inskrivet av [url=/avdelning/medlem/?m=" & rsDB(1)("aAnvNamn") & "]" & rsDB(1)("aAnvNamn") & "[/url][/i]" & vbcrlf & "[b]" & sEncode(rsDB(1)("tTextM")) & "[/b][/quote]" & vbcrlf
        Else
          text_Amne             = ""
          text_TextM            = ""
        End If
        
        text_AutoUrl          = True
        text_AutoSmil         = True
        
        text_IsTrad = False
      
      Else
        newPost = True
      End if
    
    RS_Close 1
  ElseIf lAID > 0 Then  ' ## NYTT INL�GG
    RS_Open 1, "SELECT * FROM fsBB_Tradar WHERE tID = " & CLng(lAID) & " AND tStatus_Trad = 1", False
    
      If Not rsDB(1).EOF Then
        sTitle = "Nytt inl�gg"
    
        text_ID               = CLng(rsDB(1)("tID"))
        text_Amne             = CutText("Sv: " & Trim(sEncode(rsDB(1)("tAmne"))), 100)
        
        text_AutoUrl          = True
        text_AutoSmil         = True
        
        text_IsTrad = False
      
        GetRights text_ID ' H�mta fram r�ttigheterna
        If Not sec_Inlagg_Skapa Then Response.Redirect("trad.asp?e=" & text_ID)
      Else
        newPost = True
      End if
    
    RS_Close 1
  Else              ' ## Redigera
    RS_Open 1, "SELECT * FROM fsBB_Tradar WHERE tID = " & CLng(lID), False

      If Not rsDB(1).EOF Then     
        text_ID               = CLng(rsDB(1)("tID"))
        text_Forum            = CLng(rsDB(1)("tForum"))
        text_UnderTrad        = CLng(rsDB(1)("tStatus_UnderTrad"))
        text_Amne             = Trim(sEncode(rsDB(1)("tAmne")))
        text_TextM            = Trim(sEncode(rsDB(1)("tTextM")))
        
        text_UserID           = CLng(rsDB(1)("tAnv_Skapad"))
        
        text_Locked           = rsDB(1)("tStatus_Last")
        text_Klistrad         = rsDB(1)("tInst_Klistrad")
        text_Dold             = rsDB(1)("tStatus_Dold")
        
        text_AutoSmil         = rsDB(1)("tInst_Smilies")
        text_AutoUrl          = rsDB(1)("tInst_AutoLankar")
        
        text_MainThread       = rsDB(1)("tStatus_Trad")
        
        If text_MainThread Then
          text_IsTrad = True
          sTitle = "Redigera tr�d"
          
          GetRights text_ID ' H�mta fram r�ttigheterna
          If Not sec_Trad_Hantera Then Response.Redirect("trad.asp?e=" & text_ID)
        Else
          text_IsTrad = True
          sTitle = "Redigera inl�gg"
          
          GetRights text_ID ' H�mta fram r�ttigheterna
          GetUserStatsFromPost text_ID
          If Not sec_Hantera(lPost_UserAdmin, lPost_UserID) Then Response.Redirect("trad.asp?e=" & lPost_TradID)
        End If
      
      Else
        newPost = True
      End if
    
    RS_Close 1
  End If
  
  If newPost Then
    sTitle = "Ny tr�d"
  
    text_AutoUrl          = True
    text_AutoSmil         = True
    
    text_MainThread       = True
    text_IsTrad           = True
    
    text_Forum            = CLng(lFora)
    
    text_ID               = 0
  End If
  
  If text_MainThread Then
    RS_Open 1, "SELECT fID, fName, fSplitterBefore, fSec_Mod FROM fsBB_Forum WHERE fGroup = 0 AND (fSec_NewThread = '0' OR fSec_NewThread LIKE '%;" & SEC_TITEL & ";%') ORDER BY fNoAllView ASC, fSortNr ASC", False
    
      If rsDB(1).EOF Then
        any_Forum = False
      Else
        any_Forum = True
        list_Forum = rsDB(1).GetRows
      End If
    
    RS_Close 1
  End If
  
  ' ## F�RHANDSGRANSKNING
  bGranska      = False
  'text_Granska  = BBCode(text_TextM, True)
  
  Call stop_Rec2Session("post")
%>

<%
  ' ## Globala variabler ##
  page_Title    = sTitle & " - Forum"
  page_Header   = sTitle
  page_WhereAmI = "&gt; <a href='ny_trad.asp?f=" & text_Forum & "' title='G� till &quot;" & sTitle & "&quot; ...'>" & sTitle & "</a> "
  page_SelMenu  = "forum"
  page_Slide    = "forum"
  Remove_Distans= True
  
  iFilter       = text_Forum
  
  page_description    = "Skapa eller redigera en tr�d eller ett inl�gg i forumet p� N-Forum.se, Nintendo Forum."
  page_keywords       = "ny tr�d, "
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
   
    <!-- ## SKAPA TR�DEN ## -->
    <div id="edit_preview" style="display: block;">
      
      <div class="nf_datablock nf_size_full">
        <h1><% = sTitle %></h1><a name="EDIT"> </a>

        <% If errCode > 0 Then %>
          <div class="nf_msg_full nf_red">
            <% If errCode = 1 Then %><p><strong>Posten sparades inte!</strong></p><p>Du har inte angett n�got �mne.</p><% End If %>
            <% If errCode = 2 Then %><p><strong>Posten sparades inte!</strong></p><p>Du har inte angett n�got meddelande.</p><% End If %>
            <% If errCode = 3 Then %><p><strong>Posten sparades inte!</strong></p><p>Du saknar beh�righet (A=<% = GetQ("A") %>).</p><% End If %>
            <% If errCode = 4 Then %><p><strong>Posten sparades inte!</strong></p><p>Du har inte angett vilket forum tr�den ska ligga i.</p><% End If %>
          </div>
        <% End If %>
      </div>
      
      <div class="nf_datablock nf_size_full">
        <div class="nf_forum">
          <form method="POST" action="_action/post.asp" id="frmPost">
            <div class="nf_post_msg">
              <% If text_MainThread Then %>
                <script type="text/javascript">
                  var aFora;
                  aFora = ";-2;";
          
                  function ModStat(vValue) {
                    if (aFora.indexOf(";" + vValue + ";") > 0) {
                      show("admPanel");
                      show("admPanel2");
                    } else {
                      hide("admPanel");
                      hide("admPanel2");
                    }
                  }
                </script>
              
                  <label for="kat">Forum:</label>
                  
                  <select name="kategori" id="kat" onchange="ModStat(this.value);">
                    <option value=0 style="padding: 1px 0 1px 0; font-weight: bold; color: #CCC;"> - V�lj forum - </option>
                    <option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option>
                    <% For zx = 0 To UBound(list_Forum, 2) %>
                      <% If InStr(list_Forum(3, zx),";" & CONST_TITEL & ";") Then %><script type="text/javascript">aFora = aFora + "<% = list_Forum(0, zx) %>;";</script><% End If %>
                      <% If list_Forum(2, zx) Then %><option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option><% End If %>
                      <option value=<% = list_Forum(0, zx) %> style="padding: 1px 0 1px 10px;" <% If CLng(text_Forum) = CLng(list_Forum(0, zx)) Then Response.Write(" selected") %>> <% = sEncode(list_Forum(1, zx)) %> </option>
                    <% Next %>
                  </select>
              <% End If %>
            
              <label for="fAmne">�mne:</label> <input class="text" type="text" name="fAmne" id="fAmne" value="<% = text_Amne %>" maxlength=100 onkeyup="document.getElementById('amne_preview').innerHTML=this.value;">
              
              <label for="fAutoSmil">Automatiska smilies:</label> <input class="checkbox" type="checkbox" value="<% If text_AutoSmil Then Response.Write("YES") Else Response.Write("NO") %>" onchange="if(this.checked){this.value='YES'}else{this.value='NO'}" name="fAutoSmil" id="fAutoSmil" <% If text_AutoSmil Then Response.Write(" checked") %>>
              <label for="fAutoUrl">Automatiska l�nkar:</label> <input class="checkbox" type="checkbox" value="<% If text_AutoUrl Then Response.Write("YES") Else Response.Write("NO") %>" onchange="if(this.checked){this.value='YES'}else{this.value='NO'}" name="fAutoUrl" id="fAutoUrl" <% If text_AutoUrl Then Response.Write(" checked") %>>
              
              <% If text_MainThread Then %>
                <div id="admPanel">
                  <label for="fLocked"> L�st</label> <input class="checkbox" type="checkbox" value="YES" name="fLocked" id="fLocked" <% If text_Locked Then Response.Write(" checked") %>>
                  <label for="fKlistrad"> Klistrad</label> <input class="checkbox" type="checkbox" value="YES" name="fKlistrad" id="fKlistrad" <% If text_Klistrad Then Response.Write(" checked") %>>
                  <label for="fDold"> Egna forumet</label> <input class="checkbox" type="checkbox" value="YES" name="fDold" id="fDold" <% If text_Dold Then Response.Write(" checked") %>>
                </div>
              <% End If %>
              
              <div class="nf_buttonbar">
                <input onclick="addText('aTextM','b');" type="button" value="B" style="width: 25px; font-weight: bold;">
                <input onclick="addText('aTextM','i');" type="button" value="I" style="width: 25px; font-style: italic;">
                <input onclick="addText('aTextM','u');" type="button" value="U" style="width: 25px; font-decoration: underline;">
                <input onclick="addText('aTextM','s');" type="button" value="S" style="width: 25px; text-decoration: line-through;">
                <div class="nf_buttonsplit"></div>
                <input onclick="addText('aTextM','url');" type="button" value="URL" style="width: 40px;">
                <input onclick="addText('aTextM','img');" type="button" value="IMG" style="width: 40px;">
                <!--
                <div class="nf_buttonsplit"></div>
                <input onclick="addText('aTextM','spoiler');" type="button" value="Spoiler" style="width: 56px;">
                <input onclick="addText('aTextM','indent');" type="button" value="Indenterad" style="width: 80px;">
                <input onclick="addText('aTextM','code');" type="button" value="Monospace" style="width: 80px;">
                -->
              </div>
              
              <textarea name="fMsg" id="aTextM" maxlength="20000" style="height: 600px;" onkeyup="if(this.value==''){document.getElementById('btPreview').disabled=true;}else{document.getElementById('btPreview').disabled=false;}closePreview('fMsg',false);"><% = text_TextM %></textarea>
            </div>
            <div class="nf_post_btn">
              <input type="submit" style="font-weight: bold;" value="Posta">
              <input type="button" <% If Len(text_TextM) = 0 Then %>disabled<% End If %> id="btPreview" value="F�rhandsgranska" onclick="doPreview('aTextM','' + document.getElementById('fAutoUrl').value + '','' + document.getElementById('fAutoSmil').value + '');">
              <!-- <input type="button" value="Avbryt" style="color: #A00;" onclick="doActionWithPrompt('<% = Server.HTMLEncode(sURL) %>','Vill du avbyta och �terg�?');"> -->
              
              <div id="admPanel2" style="float: left; border-top: solid 1px #CCC; margin-top: 10px; padding-top: 10px;">
                <!--
                <% If text_MainThread Then %>
                  <input type="button" onclick="showFrameBox('trad_settings.asp?e=<% = lID %>&do=fuse','Sl� ihop med annan tr�d');" value="Sl� ihop">
                  <input type="button" onclick="showFrameBox('trad_settings.asp?e=<% = lID %>&do=owner','Ta �ver �garskap av denna tr�d');" value="Ta �ver �garskap">
                <% Else %>
                  <input type="button" onclick="showFrameBox('trad_settings.asp?e=<% = lID %>&do=break','Bryt ut inl�gg till egen tr�d');" value="Bryt ut inl�gg">
                  <input type="button" onclick="showFrameBox('trad_settings.asp?e=<% = lID %>&do=move','Flytta inl�gg till annan tr�d');" value="Byt huvudtr�d">
                <% End If %>
                -->
              </div>
            </div>
            <% If text_IsTrad Then %>
              <input type="hidden" name="tradID" value=<% = text_ID %>>
            <% Else %>
              <input type="hidden" name="tradID_Svar" value=<% = text_ID %>>
            <% End If %>
            <input type="hidden" name="url" value="<% = Server.HTMLEncode(sURL) %>">
          </form>
        </div>
      </div>

      <% If text_MainThread Then %>
        <script type="text/javascript">
          ModStat(document.getElementById("kat").value);
        </script>
      <% End If %>
    </div>
    <!-- ## /SKAPA TR�DEN ## -->
  
    <!-- ## F�RHANDSGRANSKA TR�DEN ## -->
    <div id="post_preview" style="display: none;">
      
      <div class="nf_datablock nf_size_full">
        <h1>F�rhandsgranskning</h1> <a name="PREVIEW"> </a>
      </div>
    
      <div class="nf_datablock nf_size_full">
        <div class="nf_msg_full nf_green" id="warn_preview" style="display: none;">
          <p><img src="<% = config_GFXLocation %>preview_arrow.png" style="float: left; margin-right: 8px;"><strong>Din f�rhandsgranskning, klicka p� [�ndra] f�r att forts�tta redigera din post.</strong></p>
        </div>
        
        <div class="nf_forum">
           
          <div class="nf_trad">
            <div class="nf_trad_content">
              <p class="nf_trad_title" id="amne_preview"><% = text_Amne %></p>
              <p id="post_preview_text"></p>
            </div>
            
            <div class="nf_trad_data">
              <div class="nf_post_btn">
                <input type="button" style="font-weight: bold;" value="Posta" onclick="document.getElementById('frmPost').submit();">
                <input type="button" value="�ndra" onclick="closePreview('aTextM',true);">
              </div>
            </div>
          </div>
        </div>
      
      </div>
    </div>
    <!-- ## F�RHANDSGRANSKA TR�DEN ## -->
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->