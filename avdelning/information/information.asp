<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  RS_Open 1, "SELECT aAnvNamn, aTimeStamp, fsBB_Titlar.ttText AS aTitelText, aInloggadSenast, aEgenTitel, aAvatar, aID FROM fsBB_Anv " & _
             "LEFT JOIN fsBB_Titlar ON aTitelID = fsBB_Titlar.ttID " & _
             "WHERE aBlockadTill < '" & Date & "' AND aTitelID IN(5,16) ORDER BY aTitelID DESC", False
  
    If rsDB(1).EOF Then
      any_Red = False
    Else
      any_Red = True
      list_Red = rsDB(1).GetRows
    End If
  
  RS_Close 1
%>

<%

  ' ## Globala variabler ##
  page_Title    = "Information - Om N-Forum.se"
  page_Header   = "Om N-Forum.se"
  page_WhereAmI = "&gt; Om N-Forum.se "
  page_SelMenu  = "home"
  page_Slide    = "forum"
  
  page_description  = "Information om N-Forum.se, Nintendo Forum. Vilka deltar i redaktionen och vad g�ller vid l�nkbyte/l�nkv�nner."
  page_keywords     = "information, "
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full">
      <h1>Om N-Forum.se</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">

      <div class="nf_text">
        <p>N-Forum.se �r en sida av fans om konsol- och speltillverkaren <a href="http://www.nintendo.se" target="_blank" rel="nofollow">Nintendo</a> och ska <strong>inte</strong> misstas f�r att vara officiell. </p>
        <p>Allt b�rjade som <a href="http://web.archive.org/web/20050810084352/http://www.gameboy.nu/" target="_blank" rel="nofollow">Gameboy.nu</a> sommaren 2005 d� <em><a href="/avdelning/medlem/?m=folkow" target="_blank">folkow</a></em> p� <a href="http://www.listmygames.se" target="_blank" rel="nofollow">LMG</a> s�kte efter en sida som listade alla sl�ppta GameBoy-spel. Jag, <em><a href="/avdelning/medlem/?m=skogga" target="_blank">skogga</a></em>, kollade d� upp om det fanns och kunde inte hitta n�gon vilket ledde till att jag sj�lv la upp en sida som var t�nkt att lista alla de spelen. </p>
        <p>Sagt och gjort, efter bara n�gra veckor hoppade <em><a href="/avdelning/medlem/?m=stuff_larsson" target="_blank">stuff_larsson</a></em> p� t�get och b�rjade hj�lpa till att lista spel till sidan. Allt eftersom hoppade fler p� att hj�lpa till med sidan. </p>
        <p>Efter cirka ett halv�r n�r julen n�rmade sig s� ins�g jag att sida beh�vde uppgraderas f�r att bli mer l�tthanterad och i samband med detta valde jag att alla Nintendos konsoler skulle representeras. Det var b�rjan till N-Forum.se. </p> 
        <p>Utvecklingen av N-Forum.se gick tr�gt och hade passerat flera olika designer innan jag fastslog denna design som du ser nu. </p>
        <p>Vill du hj�lpa oss att fylla upp denna databas med speldata, nyheter, recensioner eller artiklar?, h�r av dig. </p>
      </div>
      
      <% If any_Red Then %>
        <ul class="nf_list">
          <li class="nf_listsplit"> N-Forum.se's redaktion </li>
          <%
            For zx = 0 To UBound(list_Red, 2)
              %>
                <li>
                  <div class="nf_icon">
                    <% If list_Red(5, zx) Then %>
                      <img src="<% = config_Avatar & "u" & Right("000000" & list_Red(6, zx), 6) & ".jpg" %>" alt="Avatar">
                    <% End If %>
                  </div>
                  <div class="nf_data">
                    <h3><a href="/avdelning/medlem/?m=<% = list_Red(0, zx) %>"><% = sEncode(list_Red(0, zx)) %></a></h3>
                    <span class="nf_medium nf_gray"><% = list_Red(2, zx) %></span>
                    <span class="nf_medium"><% = sEncode(list_Red(0, zx)) %>&#64;n-forum.se</span> 
                  </div>
                </li>
              <%
            Next
          %>
        </ul>
      <% End If %>
      
      <div class="nf_msg">
        <p><strong>Annonsera p� N-Forum.se</strong></p>
        <p>Vi kommer ALDRIG l�gga upp annonser p� denna sida som g�r till dobbel(poker osv...), porr eller andra ol�mpliga sidor. S� bem�da er inte ens med att kontakta oss.</p>
        <p>Detta �r en tv-spelssida och eventuella annonser ska d�rmed passa det omr�det.</p>
        <p>Vill ni annonsera p� N-Forum.se kan ni kontakta oss p� <a href="mailto:info@n-forum.se">info@n-forum.se</a> och diskutera det med oss, men observera att vi inte kommer fylla sidan med annonser i on�dan.<p>
      </div>
      
      <div class="nf_msg">
        <p><strong>L�nkbyten / L�nkv�nner</strong></p>
        <p>Vill du g�ra ett l�nkbyte med v�ran sida, vad kul. S� l�nge den ber�r omr�det TV-Spel. Din tillbakal�nk kommer hamna p� v�r f�rsta sida.</p>
        <p>H�r i s�dana fall bara av dig till <a href="mailto:info@n-forum.se">info@n-forum.se</a> s� ska det g� att l�sa. Vi kommer dock inte byta l�nk med vem som helst utan det avg�r vi fr�n fall till fall.</p>
        <p>Under tiden kan du f�rbereda din webbsida med en l�nk fr�n "<a href="/avdelning/information/linkback.asp">L�nk hit</a>".</p>
      </div>
      
      <div class="nf_msg">
        <p><strong>�vrig information</strong></p>
        <p>N-Forum.se �r en inofficiell fansida om Nintendo och ska INTE kopplas ihop med <a href="http://www.bergsala.se">Bergsala</a>/<a href="http://www.nintendo.se">Nintendo</a> som bolag. </p>
        <p>R�ttigheterna till spelen, karakt�rerna och de avbildade spelfodralen �r f�rbeh�llna respektive bolag. Inga intr�ng avsedda. </p>
        <p>N-Forum.se drivs helt p� ideell basis av fans. Vill ni kontakta oss n�r ni oss p� <a href="mailto:info@n-forum.se">info@n-forum.se</a>. </p>
      </div>
    </div>
    
    <div class="nf_datablock nf_size_onethird">

    </div> 
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->