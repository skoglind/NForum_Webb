<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)

  ' ## Hämta all data ##
  lID = GetQ("e", "123", 0)
  RS_Open 1, "SELECT *, anvF.aAnvNamn AS pAnvFran, anvT.aAnvNamn AS pAnvTill FROM fsBB_PM " & _
             "LEFT JOIN fsBB_Anv AS anvF ON pFran = anvF.aID " & _
             "LEFT JOIN fsBB_Anv AS anvT ON pTill = anvT.aID " & _
             "WHERE pID = " & CLng(lID) & " AND (pTill = " & CLng(CONST_USERID) & " OR pFran = " & CLng(CONST_USERID) & ")", True
  
    If rsDB(1).EOF Then Response.Redirect("inkorg.asp")
    
    text_ID         = CLng(rsDB(1)("pID"))
    text_Amne       = sEncode(CutText(rsDB(1)("pAmne"),45))
    text_Text       = BBCode(rsDB(1)("pPM"), True)
    text_Datum      = DatumReplace(rsDB(1)("pDatum"))
    
    text_TillID     = rsDB(1)("pTill")
    text_FranID     = rsDB(1)("pFran")
    
    text_Till       = rsDB(1)("pAnvTill")
    text_Fran       = rsDB(1)("pAnvFran")
    
    activatedPM     = GetSendPM(pFran)
    
    If rsDB(1)("pTill") = CLng(CONST_USERID) And rsDB(1)("pLast") = False Then 
      rsDB(1)("pLast") = True
      rsDB(1).Update
    End If
  
  RS_Close 1
%>

<%
  ' ## Globala variabler ##
  page_Title    = text_Amne & " - PM - Medlem"
  page_Header   = text_Amne
  page_WhereAmI = "&gt; PM &gt; Visa PM"
  page_SelMenu  = "user"
  page_Slide    = "medlem"
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_u.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1><% = text_Amne %></h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">

      <div class="nf_msg">
        <p><strong>Datum:</strong> <% = text_Datum %></p>
        <% If text_TillID = CLng(CONST_USERID) Then %><p><strong>Från:</strong> <a href="/avdelning/medlem/?m=<% = text_Fran %>"><% = text_Fran %></a></p<% End If %>
        <% If text_FranID = CLng(CONST_USERID) Then %><p><strong>Till:</strong> <a href="/avdelning/medlem/?m=<% = text_Till %>"><% = text_Till %></a></p><% End If %>
      </div>
      
      <div class="nf_text">
        <p><% = text_Text %></p>
      </div>
      
      <div class="nf_form">
        <div class="nf_falt nf_buttons">
          <% If text_TillID = CLng(CONST_USERID) Then %>
            <input type="button" value="Besvara PM" onclick="location.href='skrivpm.asp?a=<% = text_ID %>';" <% If Not activatedPM Or Not CONST_PM Then Response.Write(" disabled") %>>
          <% End If %>
          <input type="button" value="Radera" style="color: #A00; font-weight: normal !important;" onclick="doActionWithPrompt('_action/deletepm.asp?e=<% = text_ID %>','Vill du radera detta PM?');">
        </div>
      </div>
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