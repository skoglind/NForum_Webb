<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  lID = GetQ("a", "123", 0)
  lMedlem = GetQ("m","ABC",50)

  errCode = GetQ("fail","123",0)
 
  If Session.Value("record_pm") Then
    text_Till   = sEncode(Session.Value("pm_pTill"))
    text_Amne   = sEncode(Session.Value("pm_pAmne"))
    text_TextM  = sEncode(Session.Value("pm_pMsg"))
    
    If Session.Value("pm_pAnswer") Then lockedTill = True
    
    Session.Value("record_pm") = False
  ElseIf lID <> 0 Then
    RS_Open 1, "SELECT *, anvF.aAnvNamn AS pAnvFran FROM fsBB_PM " & _
             "LEFT JOIN fsBB_Anv AS anvF ON pFran = anvF.aID " & _
             "WHERE pID = " & CLng(lID) & " AND pTill = " & CLng(CONST_USERID), True
  
      If Not rsDB(1).EOF Then
        text_Till       = rsDB(1)("pAnvFran")
        text_Amne       = "Sv: " & sEncode(CutText(rsDB(1)("pAmne"),45))
        
        text_TextM      = vbCrlf & vbCrlf & vbCrlf & "[b] -- Meddelande nedan av [url=/avdelning/medlem/?m=" & text_Till & "]" & text_Till & "[/url] (" & rsDB(1)("pDatum") & ") -- [/b]" & vbCrlf & sEncode(rsDB(1)("pPM"))
        
        lockedTill      = True
      End If
      
    RS_Close 1
  Else
      If Len(lMedlem) > 0 Then
        text_Till         = sEncode(lMedlem)
        lockedTill        = True
      End If
  End If
 
  Call stop_Rec2Session("pm")
%>

<%
  ' ## Globala variabler ##
  page_Title    = "Skriv ett PM - PM - Medlem"
  page_Header   = "Skriv ett PM"
  page_WhereAmI = "&gt; PM &gt; <a href='skrivpm.asp' title='Gå till &quot;Skriv ett PM&quot; ...'>Skriv ett PM</a> "
  page_SelMenu  = "user"
  page_Slide    = "medlem"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_u.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1>Skriv ett PM</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">

      <% If errCode > 0 Then %>
        <div class="nf_msg nf_red">
          <% If errCode = 1 Then %><p><strong>PM skickades inte!</strong></p><p>Du har angett ett felaktigt användarnamn.</p><% End If %>
          <% If errCode = 2 Then %><p><strong>PM skickades inte!</strong></p><p>Du kan inte skicka PM till dig själv.</p><% End If %>
          <% If errCode = 3 Then %><p><strong>PM skickades inte!</strong></p><p>Du har inte angett något ämne.</p><% End If %>
          <% If errCode = 4 Then %><p><strong>PM skickades inte!</strong></p><p>Du har inte skrivit någon text.</p><% End If %>
          <% If errCode = 5 Then %><p><strong>PM skickades inte!</strong></p><p>Du kan inte skicka PM till denna användare.</p><% End If %>
          <% If errCode = 6 Then %><p><strong>PM skickades inte!</strong></p><p>Du kan inte skicka PM då du har inaktiverat funktionen.</p><% End If %>
        </div>
      <% End If %>
      
      <form method="POST" action="_action/sendpm.asp">
      
        <div class="nf_form">
        
          <div class="nf_falt"><label>Till:</label> <input type="text" name="pTill" value="<% = text_Till %>" maxlength=100 style="width: 436px;"  <% If lockedTill Then Response.Write(" disabled") %>></div>
          <div class="nf_falt"><label>Ämne:</label> <input type="text" name="pAmne" value="<% = text_Amne %>" maxlength=100 style="width: 436px;"></div>
        
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
            <textarea name="pMsg" id="aTextM" style="height: 260px; width: 576px" maxlength="20000" onkeyup="return ismaxlength(this)"><% = text_TextM %></textarea>
          </div>
          
          <div class="nf_falt nf_buttons">
            <% If lockedTill Then %>
              <input type="hidden" name="pTill" value="<% = text_Till %>">
              <input type="hidden" name="pAnswer" value="YES">
            <% End If %>
          
            <input type="submit" value="Skicka PM">
          </div>
          
        </div>
      
      </form>
    
    </div>
    
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_minibox">
        <h4>PM - Personliga meddelanden</h4>
        <div class="nf_inside">
          <!--#INCLUDE FILE="_sidebar_pm.asp"-->
        </div>
      </div>
    </div>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->