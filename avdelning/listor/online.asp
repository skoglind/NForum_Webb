<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  RS_Open 1, "SELECT aAnvNamn, aTimeStamp, fsBB_Titlar.ttText AS aTitelText, aInloggadSenast, aEgenTitel, aAvatar, aID FROM fsBB_Anv " & _
             "LEFT JOIN fsBB_Titlar ON aTitelID = fsBB_Titlar.ttID " & _
             "WHERE aTimeStamp > '" & DateAdd("n", -5, Now) & "' AND aBlockadTill < '" & Date & "' AND aAktiverad = 1 ORDER BY aInloggadSenast ASC", False
  
    If rsDB(1).EOF Then
      any_On = False
    Else
      any_On = True
      list_On = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  If any_On Then
    For zx = 0 to UBound(list_On, 2)
      subUsr = subUsr & list_On(6, zx) & ","
    Next
    
    subUsr = Left(subUsr, Len(subUsr) - 1)
  Else
    subUsr = 0
  End if
  
  RS_Open 1, "SELECT aAnvNamn, aTimeStamp, fsBB_Titlar.ttText AS aTitelText, aInloggadSenast, aEgenTitel, aAvatar, aID FROM fsBB_Anv " & _
             "LEFT JOIN fsBB_Titlar ON aTitelID = fsBB_Titlar.ttID " & _
             "WHERE aID NOT IN (" & subUsr & ") AND aInloggadSenast > '" & DateAdd("h", -48, Now) & "' AND aBlockadTill < '" & Date & "' AND aAktiverad = 1 ORDER BY aInloggadSenast DESC", False
  
    If rsDB(1).EOF Then
      any_On48 = False
    Else
      any_On48 = True
      list_On48 = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  ' ### Fler senast registrerade
  RS_Open 1, "SELECT TOP 10 aAnvNamn, aTimeStamp, fsBB_Titlar.ttText AS aTitelText, aMedlemSedan, aEgenTitel, aAvatar, aID FROM fsBB_Anv " & _
             "LEFT JOIN fsBB_Titlar ON aTitelID = fsBB_Titlar.ttID " & _
             "WHERE aBlockadTill < '" & Date & "' AND aAktiverad = 1 ORDER BY aMedlemSedan DESC", False
  
    If rsDB(1).EOF Then
      any_Reg = False
    Else
      any_Reg = True
      list_Reg = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1
  
  antalMemb = con.ExeCute("SELECT COUNT(aID) FROM fsBB_Anv WHERE aBlockadTill < '" & Date & "' AND aAktiverad = 1")(0)
%>

<%
  ' ## Globala variabler ##
  page_Title    = "Medlemsstatistik"
  page_Header   = "Medlemmar online just nu"
  page_WhereAmI = "&gt; <a href='online.asp' title='Gå till &quot;Online&quot; ...'>Online</a> "
  page_SelMenu  = "user"
  page_Slide    = "medlem"
  
  page_description  = "Medlemmar som är online just nu på på N-Forum.se, Nintendo Forum. Senast registrerade medlemmarna och de senast inloggade."
  page_keywords     = "online ,"
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1>Medlemsstatistik</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      
      <% If any_On Then %>
        <ul class="nf_list">
          <li class="nf_listsplit"> Inloggad just nu </li>
        
          <%
            For zx = 0 To UBound(list_On, 2)
              %>
                <li>
                  <div class="nf_icon">
                    <% If list_On(5, zx) Then %>
                      <img src="<% = config_Avatar & "u" & Right("000000" & list_On(6, zx), 6) & ".jpg" %>" alt="Avatar">
                    <% End If %>
                  </div>
                  <div class="nf_data">
                    <h3><a href="/avdelning/medlem/?m=<% = list_On(0, zx) %>"><% = sEncode(list_On(0, zx)) %></a></h3>
                    <span class="nf_medium nf_gray"><% If Len(list_On(4, zx)) > 0 Then %><% = list_On(4, zx) %><% Else %><% = list_On(2, zx) %><% End If %></span> 
                    <span class="nf_bold">Online i <% = MinutesToSplit(DateDiff("n", list_On(3, zx), list_On(1, zx))) %></span>
                  </div>
                </li>
              <%
            Next
          %>
        </ul>
      <% Else %>
        <div class="nf_msg"><p>Det är ingen online på N-Forum.se just nu.</p></div>
      <% End If %>
      
      <% If any_On48 Then %>
        <ul class="nf_list">
          <li class="nf_listsplit"> Inloggad de senaste 2 dagarna </li>
        
          <%
            For zx = 0 To UBound(list_On48, 2)
              %>
                <li>
                  <div class="nf_icon">
                    <% If list_On48(5, zx) Then %>
                      <img src="<% = config_Avatar & "u" & Right("000000" & list_On48(6, zx), 6) & ".jpg" %>" alt="Avatar">
                    <% End If %>
                  </div>
                  <div class="nf_data">
                    <h3><a href="/avdelning/medlem/?m=<% = list_On48(0, zx) %>"><% = sEncode(list_On48(0, zx)) %></a></h3>
                    <span class="nf_medium nf_gray"><% If Len(list_On48(4, zx)) > 0 Then %><% = list_On48(4, zx) %><% Else %><% = list_On48(2, zx) %><% End If %></span> 
                    <span class="nf_bold">Inloggad: <% = DatumReplace(list_On48(3, zx)) %></span>
                  </div>
                </li>
              <%
            Next
          %>
        </ul>
      <% End If %>
    </div>
    
    <div class="nf_datablock nf_size_onethird">
      
      <div class="nf_minibox nf_blue">
        <h4>Information</h4>
        <div class="nf_inside">
          <% numOnSite = CLng(Application("nfOnline")) - CLng(pagingNumOfPosts) %>
          <% If numOnSite < 0 Then numOnSite = 0 %>
          <p>Antal besökare som inte är inloggade på sidan <strong><% = numOnSite %> st</strong></p>
          
          <p>Vi har i dagsläget <strong><% = CLng(antalMemb) %> st</strong> medlemmar.</p>
        </div>
      </div>
      
      <!-- ## SENASTE REGISTRERADE ## -->
      <% If any_Reg Then %>
        <div class="nf_minibox nf_red">
          <h4>Senaste registrerade medlemmarna</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_Reg, 2) %>
                <li onclick="location.href='/avdelning/medlem/?m=<% = sEncode(list_Reg(0,zx)) %>';"><a href="/avdelning/medlem/?m=<% = sEncode(list_Reg(0,zx)) %>" title="<% = sEncode(list_Reg(0,zx)) %>"><% = sEncode(CutText(list_Reg(0,zx), 32)) %></a> Registrerad: <% = DatumReplace(list_Reg(3,zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/listor/sokmedlem.asp">Sök medlem</a></p>
          </div>
        </div>
      <% End If %>
      <!-- ## /SENASTE REGISTRERADE ## -->
      
    </div>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->