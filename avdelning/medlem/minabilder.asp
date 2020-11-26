<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  ' ## Hämta all data ##
  lMedlem = GetQ("m","ABC",50)
  If Trim(lMedlem) = Empty Then lMedlem = CONST_USERNAME

  If Not dbUserExists(lMedlem) Then Response.Redirect("/")
  anvID = GetIDFromUsername(lMedlem) 
  
  errCode = GetQ("fail","123",0)
  
  RS_Open 1, "SELECT abID, abTitel, abUppladdadDatum FROM cms_AnvBilder WHERE abUppladdadAv = " & CLng(anvID), False
                 
    If rsDB(1).EOF Then
      any_Bild = False
    Else
      any_Bild = True
      list_Bild = rsDB(1).GetRows()
      antal_bilder = UBound(list_Bild, 2) + 1
    End If
  
  RS_Close 1
  
  If CLng(anvID) <> CONST_USERID Then canEdit = False Else canEdit = True
  If HasAcc(CONST_CMS_RIGHTS,"CMS202") Then canEdit = True
  If CLng(antal_bilder) >= config_UserMaxImages Then canUpload = False Else canUpload = True
  If config_UserImagesDays > CONST_DAYSMEMBER Then canUpload = False
  
  If antal_bilder < 1 Then antal_bilder = 0
%>

<%
  ' ## Globala variabler ##
  page_Title    = lMedlem & " - Bilder - Medlem"
  page_Header   = lMedlem & "s bilder"
  page_WhereAmI = "&gt; <a href='default.asp?m=" & lMedlem & "' title='Gå till &quot;Hem&quot; ...'>Profil</a> " & _
                  "&gt; Bilder"
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
      <h1><% = lMedlem %>s bilder (<% = antal_bilder %> st)</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      <% If anvID = CONST_USERID Then %>
        <div class="nf_msg"><p>Bilden får inte vara större än <strong>500kB</strong>. Formatet på bilden som laddas upp måste vara <strong>.png, .jpg, .bmp</strong> eller <strong>.gif</strong>.</p></div>
        
        <% If errCode > 0 Then %>
          <div class="nf_msg nf_red">
            <% If errCode = 1 Then %><p><strong>Bilden laddades inte upp!</strong><p><p>Du har inte valt någon fil.</p><% End If %>
            <% If errCode = 2 Then %><p><strong>Bilden laddades inte upp!</strong><p><p>Filen är för stor.</p><% End If %>
            <% If errCode = 3 Then %><p><strong>Bilden laddades inte upp!</strong><p><p>Den var inte av formatet (jpg,bmp,png,gif).</p><% End If %>
            <% If errCode = 4 Then %><p><strong>Bilden laddades inte upp!</strong><p><p>Du har redan <strong><% = config_UserMaxImages %></strong> bilder.</p><% End If %>
            <% If errCode = 5 Then %><p><strong>Bilden laddades inte upp!</strong><p><p>Du har inte varit medlem i <strong><% = config_UserImagesDays %></strong> dagar.</p><% End If %>
          </div>
        <% End If %>
    
        <div id="savemess" class="nf_infomsg" style="display: none;"><p>Åtgärden är utförd</p></div>
    
        <form method="POST" action="_action/uploadbild.asp" enctype="multipart/form-data">        
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
            </div>
          </div>
        </form>
        
        <% If Session.Value("form_saved") Then %>
          <script type="text/javascript">
            show("savemess");
            setTimeout("hide('savemess');", 2500);
          </script>
          <% Session.Value("form_saved") = False %>
        <% End If %>
      <% End If %>
      
      <% If any_Bild Then %>
        <div class="nf_images">
          <% For zx = 0 To UBound(list_Bild, 2) %>
            <% If canEdit Then %>
              <div class="nf_editrow">
                <a href="/userimage.asp?e=<% = list_Bild(0, zx) %>" rel="lightbox[minabilder]" target="_blank"><img src="/userimage.asp?e=<% = list_Bild(0, zx) %>" title="<% = sEncode(list_Bild(1, zx)) %> / Uppladdad: <% = DatumReplace(list_Bild(2, zx)) %>" alt="<% = sEncode(list_Bild(1, zx)) %>"></a>
                <img class="nf_imgbutton" src="<% = config_GFXLocation %>icons/del.png" title="Ta bort" onclick="doActionWithPrompt('_action/deletebild.asp?e=<% = list_Bild(0, zx) %>','Vill du ta bort bilden?');">
                <p><strong>Filnamn:</strong> <% = sEncode(list_Bild(1, zx)) %></p>
                <p><strong>Uppladdad:</strong> <% = DatumReplace(list_Bild(2, zx)) %></p>
              </div>
            <% Else %>
              <a href="/userimage.asp?e=<% = list_Bild(0, zx) %>" rel="lightbox[minabilder]" target="_blank"><img src="/userimage.asp?e=<% = list_Bild(0, zx) %>" title="<% = sEncode(list_Bild(1, zx)) %> / Uppladdad: <% = DatumReplace(list_Bild(2, zx)) %>" alt="<% = sEncode(list_Bild(1, zx)) %>"></a>
            <% End If %>
          <% Next %>
        </div>
      <% Else %>
        <div class="nf_msg">
          <p>Det finns inga bilder att visa.</p>
        </div>
      <% End If %>
    </div>
    
    <div class="nf_datablock nf_size_onethird">     
      <div class="nf_minibox nf_blue">
        <h4>Information</h4>
        <div class="nf_inside">
          <p>Du kan endast ha <strong><% = config_UserMaxImages %></strong> bilder åt gången.</p>
          <p>Du måste ha varit medlem i minst <strong><% = config_UserImagesDays %></strong> dagar för att kunna ladda upp bilder.</p>
        </div>
      </div>
      
      <div class="nf_minibox nf_red">
        <h4>Observera</h4>
        <div class="nf_inside">
          <p>Ladda inte upp några olämpliga eller olagliga bilder.</p>
          <p>Detta kan leda till att din användare blir avstängd.</p>
        </div>
      </div>

    </div>

  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->