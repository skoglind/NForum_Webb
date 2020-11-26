<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  lMedlem = GetQ("m","ABC",50)
  If Trim(lMedlem) = Empty Then lMedlem = CONST_USERNAME

  If Not dbUserExists(lMedlem) Then Response.Redirect("/")
  anvID = GetIDFromUsername(lMedlem) 
  
  If config_LockDown_Feedback Then Response.Redirect("../default.asp?m=" & lMedlem)
  
  RS_Open 1, "SELECT fbID, fbTextM, fbAnv, fbDatum, fbRaderadAv, fbMedlem, " & _
             "fsBB_Anv.aAnvNamn, fsBB_Anv.aID, fsBB_Anv.aAvatar, fsBB_Anv.aPlats, fsBB_Anv.aTimeStamp, fsBB_Anv.aAktiveraPM " & _
             "FROM cms_Feedback " & _
             "LEFT JOIN fsBB_Anv ON cms_Feedback.fbAnv = aID " & _
             "WHERE fbMedlem = " & CLng(anvID) & " " & _
             "ORDER BY fbDatum ASC", False
   
     If rsDB(1).EOF Then
      any_Comments = False
    Else
      any_Comments = True
      list_Comments = rsDB(1).GetRows
    End If
   
  RS_Close 1
  
  canEdit = False
  If CLng(CONST_USERID) = CLng(text_AvID) Then canEdit = True
%>

<%
  ' ## Globala variabler ##
  page_Title    = lMedlem & " - Omdömen som säljare - Medlem"
  page_Header   = lMedlem & "s omdömen som säljare"
  page_WhereAmI = "&gt; <a href='default.asp?m=" & lMedlem & "' title='Gå till &quot;Hem&quot; ...'>Profil</a> " & _
                  "&gt; Omdömen om säljare"
  page_SelMenu  = "user"
  page_Slide    = "medlem"
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <% If CONST_LOGIN And CLng(anvID) = CLng(CONST_USERID) Then %>
    <!--#INCLUDE FILE="__menu_u.asp"-->
  <% Else %>
    <!--#INCLUDE FILE="__menu_other.asp"-->
  <% End If %>
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full">
      <h1>Omdömen</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      <ul class="nf_list">
        
        <li class="nf_listsplit"> Omdömen </li>
        
        <% If any_Comments Then %>
        
          <% For zx = 0 To UBound(list_Comments, 2) %>
            <li> 
              <div class="nf_header">
                <span class="nf_big">#<% = zx + 1 %></span>
              </div>
              <div class="nf_data">
                <span class="nf_medium nf_gray nf_bold">
                  <a href="/avdelning/medlem/?m=<% = list_Comments(6, zx) %>"><% = list_Comments(6, zx) %></a> / <% = DatumReplace(list_Comments(3, zx)) %>
                  <% If (CLng(CONST_USERID) = CLng(list_Comments(2, zx)) Or HasAcc(CONST_CMS_RIGHTS,"CMS700")) And CLng(list_Comments(4, zx)) = 0 Then %><img src="<% = config_GFXLocation %>icons/del.png" onclick="doActionWithPrompt('_action/deletecomment.asp?e=<% = list_Comments(0, zx) %>&amp;m=<% = sEncode(lMedlem) %>','Vill du ta bort omdömet?');" style="float: right; cursor: pointer;" title="Ta bort omdömet" alt="Radera"><% End If %>
                </span>
                
                <% If CLng(list_Comments(4, zx)) = 0 Then %>
                  <p><% = BBCode(list_Comments(1, zx), True) %></p>
                <% Else %>
                  <p style="font-size: 10px !important;font-style: italic !important; color: #A00 !important;">Omdömet är borttagen av <strong><% If CLng(list_Comments(4, zx)) = CLng(list_Comments(2, zx)) Then %>användaren<% Else %>administratören<% End If %></strong>!</p>
                  
                  <% If HasAcc(CONST_CMS_RIGHTS,"CMS700") Then %>
                    <p style="font-size: 10px !important; color: #CCC !important;"><% = BBCode(list_Comments(1, zx), True) %></p>
                  <% End If %>
                <% End If %>
                
                <% If CLng(list_Comments(4, zx)) = 0 And list_Comments(11, zx) And CONST_LOGIN Then %><span class="nf_small nf_bold">» <a href="/avdelning/medlem/skrivpm.asp?m=<% = list_Comments(6, zx) %>">Skicka PM</a></span><% End If %>
              </div>
            </li>
          <% Next %>
        
        <% End If %>
        
      </ul>
      
      <% If Not any_Comments Then %>
        <div class="nf_msg">
          <p> Det finns inga omdömen. </p>
        </div>
      <% End If %>
      
      <% If CONST_LOGIN Then %>
        <form method="POST" action="_action/postcomment.asp">
          <div class="nf_form">
  
            <div class="nf_falt"><textarea name="aMsg" style="height: 100px; width: 576px"></textarea></div>
            
            <div class="nf_falt nf_buttons">
              <input type="hidden" name="e" value="<% = lID %>">
              <input type="hidden" name="m" value="<% = sEncode(lMedlem) %>">
              <input type="submit" style="font-weight: bold;" value="Posta">
            </div>
  
          </div>
        </form>
      <% Else %>
        <div class="nf_msg">
          <p> Du måste <a href="<% = config_NotLoggedIn %>">logga in</a> för att kunna lämna ett omdöme. </p>
        </div>
      <% End If %>
    </div>
    
    <div class="nf_datablock nf_size_onethird">
    
      <div class="nf_minibox nf_blue">
        <h4>Omdömen</h4>
        <div class="nf_inside">
          <p>Här lämnar du omdömen när du köpt eller sålt något av användaren.</p>
          <p><strong>Detta är INTE tänkt för övrigt chattande.</strong></p>
        </div>
      </div>
    
    </div>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->