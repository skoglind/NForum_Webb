<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  
  bActivationFailed = False
  
  sUser = Trim(GetQ("ua","ABC",30))
  sKey  = Trim(GetQ("x","ABC",50))
  
  If Len(sUser) < 1             Then bActivationFailed = True
  If Len(sKey) < 1              Then bActivationFailed = True
  If MakeLegal(sUser) <> sUser  Then bActivationFailed = True
  If MakeLegal(sKey) <> sKey    Then bActivationFailed = True
  
  If Not bActivationFailed Then
    sUser = MakeLegal(sUser)
    sKey  = MakeLegal(sKey)
    
    RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aAnvNamn = '" & sUser & "' AND aAktiveradEpost = '" & sKey & "' AND Len(aNyEpost) > 5", True
    
      If Not rsDB(1).EOF Then
        
        rsDB(1)("aEpost")           = rsDB(1)("aNyEpost")
        rsDB(1)("aAktiveradEpost")  = ""
        rsDB(1)("aNyEpost")         = ""
        rsDB(1).Update
      Else
        bActivationFailed = True
      End If
    
    RS_Close 1
  End If
%>

<%

  ' ## Globala variabler ##
  page_Title    = "Aktivera - Medlem"
  page_Header   = "Aktivera din e-postadress"
  page_WhereAmI = "&gt; <a href='installningar.asp' title='Gå till &quot;Inställningar&quot; ...'>Inställningar</a> &gt; E-postadressverifiering"
  page_SelMenu  = "user"
  page_Slide    = "innanlogin"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu_u.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <% If Not bActivationFailed Then %>
        <h1>Din e-postadress är nu ändrad!</h1>
      <% Else %>
        <h1>Verifieringen misslyckades!</h1>
      <% End If %>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
    
      <% If Not bActivationFailed Then %>
        <div class="nf_msg">
          <p>Din e-postadress har nu blivit ändrad på N-Forum.se.</p>
        </div>
      <% Else %>
        <div class="nf_msg">
          <p><strong>Din e-postadress har inte blivit bytt. Det kan bero på följande orsaker:</strong></p>
          <p>- E-postadressen är redan aktiverad</p>
          <p>- Aktiveringsnyckeln är felaktig</p>
          <p>Om du vet att nyckeln är korrekt och att e-postadressen inte är aktiverad kontakta då <a href="mailto:info@n-forum.se">info@n-forum.se</a> för att få hjälp.</p>
        </div>
      <% End If %>
      
    </div>
    
    <div class="nf_datablock nf_size_onethird">

    </div>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->