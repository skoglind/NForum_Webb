<!--#INCLUDE FILE="../__INC/includes.asp"-->

  <%
    lRegKod   = GetQ("kod","ABC",13)
    
    retData = getRegKod(lRegKod)
    retData = Split(retData,"-")
    
    retData(0) = CLng(retData(0))
    retData(1) = CLng(retData(1))
    
    If CLng(retData(1)) > 0 Then
      ' Träff
      
      RS_Open 1, "SELECT * FROM cms_SpelTitlar LEFT JOIN cms_Spel ON sID = tSpelID WHERE tID = " & CLng(retData(1)), False
        If rsDB(1).EOF Then
          bTraff = False
        Else
          bTraff = True
          
          lTitelID  = rsDB(1)("tID")
          sTitel    = rsDB(1)("tTitel")
          sRegKod   = rsDB(1)("tRegionsKod")
          sKonsol   = lstKonsolSuperShort(rsDB(1)("sKonsol"))
          
          lImgID = 0
          If CLng(rsDB(1)("tBoxart_Kassett")) > 0 Then lImgID = CLng(rsDB(1)("tBoxart_Kassett"))
          If CLng(rsDB(1)("tBoxart_Manual")) > 0 Then lImgID = CLng(rsDB(1)("tBoxart_Manual"))
          If CLng(rsDB(1)("tBoxart_BoxFram")) > 0 Then lImgID = CLng(rsDB(1)("tBoxart_BoxFram"))
          
        End If
      RS_Close 1
    Else
      ' Ingen träff
      
      bTraff = False
    End if
  
    If retData(0) = 0 Then bSure = False Else bSure = True
  %>
  
  <% If bTraff Then %>  
    <p class="nf_center"><em>&Auml;r det h&auml;r spelet du s&ouml;ker?</em></p>
    
    <% If lImgID > 0 Then %>
      <p class="nf_center"><a href="/avdelning/spel/spel_visa_info.asp?e=<% = lTitelID %>"><img src="<% = config_ImageLocation %>?e=<% = lImgID %>&amp;h=150&amp;w=150" alt="<% = Server.HTMLEncode(sTitel) %>" style="margin: 0 55px 0 56px; width: 150px; height: 150px;"></a></p>
    <% End If %>
    
    <p class="nf_center"><a href="/avdelning/spel/spel_visa_info.asp?e=<% = lTitelID %>"><% = Server.HTMLEncode(sTitel) %></a> (<% = sKonsol %>)</p>
    <p class="nf_gray nf_center nf_bold"><% = sRegKod %></p>
  <% Else %>
    <p class="nf_center"><em>Hittade inget. F&ouml;r f&aring; tecken kanske...</em></p>
  <% End If %>

<!--#INCLUDE FILE="../__INC/includes_end.asp"-->