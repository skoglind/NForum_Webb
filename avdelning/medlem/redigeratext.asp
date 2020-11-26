<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  ' ## Hämta all data ##
  lID           = GetQ("e", "123", 0)
  
  sTextAvd      = UCase(GetQ("avd","ABC",10))
  
  filter_list   = GetQ("list","ABC",15)
  filter_page   = CLng(GetQ("page","123",0))

  errCode = GetQ("fail","123",0)
  
  Select Case sTextAvd
    Case "REC"
      RS_Open 1, "SELECT * FROM cms_Recensioner WHERE (rStatus = 1 Or rStatus = 3) And rSkapadAv = " & CLng(CONST_USERID) & " And rID = " & CLng(lID), False
        If Not rsDB(1).EOF Then
          text_TextM  = rsDB(1)("rText")
          text_Status = rsDB(1)("rStatus")
          text_Notes  = rsDB(1)("rNotes")
        Else
          Response.Redirect("minatexter.asp?list=" & filter_list & "&page=" & filter_page)
        End If 
      RS_Close 1
      
      sTitel = "Recension"
    Case "ART"
      RS_Open 1, "SELECT * FROM cms_Artiklar WHERE (aaStatus = 1 Or aaStatus = 3) And aaSkapadAv = " & CLng(CONST_USERID) & " And aaID = " & CLng(lID), False
        If Not rsDB(1).EOF Then
          text_TextM  = rsDB(1)("aaText")
          text_Status = rsDB(1)("aaStatus")
          text_Notes  = rsDB(1)("aaNotes")
        Else
          Response.Redirect("minatexter.asp?list=" & filter_list & "&page=" & filter_page)
        End If 
      RS_Close 1
      
      sTitel = "Artikel"
    Case Else
      Response.Redirect("minatexter.asp")
  End Select
    
  If Session.Value("record_txt") Then  ' ## POSTBACK
    text_TextM            = sEncode(Session.Value("txt_TextM"))
  End If
  
  Call stop_Rec2Session("txt")
%>

<%
  ' ## Globala variabler ##
  page_Title    = "Redigera " & sTitel & " - Medlem"
  page_Header   = "Redigera text"
  page_WhereAmI = "&gt; <a href='default.asp?m=" & lMedlem & "' title='Gå till &quot;Hem&quot; ...'>Profil</a> " & _
                  "&gt; Redigera " & sTitel
  page_SelMenu  = "user"
  page_Slide    = "medlem"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_u.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1>Redigera <% = sTitel %></h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">

      <% If text_Status = 3 Then %>
        <div class="nf_msg nf_red">
          <p><strong>Denna text har nekats publicering!</strong></p>
          <p><% = BBCode(text_Notes, False) %></p>
        </div>
      <% End If %>
      
      <% If errCode > 0 Then %>
        <div class="nf_msg nf_red">
          <% If errCode = 1 Then %><p><strong>Texten sparades inte!</strong></p><p>Du har inte angett en nog lång text, den ska innehålla något också.</p><% End If %>
        </div>
      <% End If %>
      
      <form method="POST" action="_action/savetext.asp">
      
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
            <textarea name="TextM" id="aTextM" style="height: 500px; width: 576px" maxlength="20000" onkeyup="return ismaxlength(this)"><% = text_TextM %></textarea>
          </div>
          
          <div class="nf_falt nf_buttons">
            <input type="hidden" name="e" value="<% = CLng(lID) %>">
          
            <input type="hidden" name="list" value="<% = filter_list %>">
            <input type="hidden" name="page" value="<% = filter_page %>">
            <input type="hidden" name="avd" value="<% = sTextAvd %>">
          
            <input type="submit" value="Spara">
            <input type="button" value="Avbryt" onclick="if(confirm('Vill du avbryta redigeringen, allt som inte har sparats kommer försvinna?')){location.href='minatexter.asp?list=<% = filter_list %>&amp;page=<% = filter_page %>'}">
          </div>
          
        </div>
      
      </form>
    
    </div>
    
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_minibox nf_blue">
        <h4>Information</h4>
        <div class="nf_inside">
          <% If CONST_PUBLISH Then %>
            <p>Texter som du skickar in kommer synas direkt på sidan, om du missbrukar detta kommer du bli av med den möjligheten.</p>
            <p><strong>Korrekturläs texten innan du skickar in den!</strong></p>
          <% Else %>
            <p>Observera att din inskickade text inte kommer synas på en gång då den ska godkännas av oss först. Du kommer få ett PM av oss när den har blivit godkänd.</p>
            <p><strong>Korrekturläs texten innan du skickar in den!</strong> Vi godkänner inte oläsliga texter.</p>
          <% End If %>
        </div>
      </div>
    </div>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->