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
  
  page_description  = "Information om N-Forum.se, Nintendo Forum. Vilka deltar i redaktionen och vad gäller vid länkbyte/länkvänner."
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
        <p>N-Forum.se är en sida av fans om konsol- och speltillverkaren <a href="http://www.nintendo.se" target="_blank" rel="nofollow">Nintendo</a> och ska <strong>inte</strong> misstas för att vara officiell. </p>
        <p>Allt började som <a href="http://web.archive.org/web/20050810084352/http://www.gameboy.nu/" target="_blank" rel="nofollow">Gameboy.nu</a> sommaren 2005 då <em><a href="/avdelning/medlem/?m=folkow" target="_blank">folkow</a></em> på <a href="http://www.listmygames.se" target="_blank" rel="nofollow">LMG</a> sökte efter en sida som listade alla släppta GameBoy-spel. Jag, <em><a href="/avdelning/medlem/?m=skogga" target="_blank">skogga</a></em>, kollade då upp om det fanns och kunde inte hitta någon vilket ledde till att jag själv la upp en sida som var tänkt att lista alla de spelen. </p>
        <p>Sagt och gjort, efter bara några veckor hoppade <em><a href="/avdelning/medlem/?m=stuff_larsson" target="_blank">stuff_larsson</a></em> på tåget och började hjälpa till att lista spel till sidan. Allt eftersom hoppade fler på att hjälpa till med sidan. </p>
        <p>Efter cirka ett halvår när julen närmade sig så insåg jag att sida behövde uppgraderas för att bli mer lätthanterad och i samband med detta valde jag att alla Nintendos konsoler skulle representeras. Det var början till N-Forum.se. </p> 
        <p>Utvecklingen av N-Forum.se gick trögt och hade passerat flera olika designer innan jag fastslog denna design som du ser nu. </p>
        <p>Vill du hjälpa oss att fylla upp denna databas med speldata, nyheter, recensioner eller artiklar?, hör av dig. </p>
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
        <p><strong>Annonsera på N-Forum.se</strong></p>
        <p>Vi kommer ALDRIG lägga upp annonser på denna sida som går till dobbel(poker osv...), porr eller andra olämpliga sidor. Så bemöda er inte ens med att kontakta oss.</p>
        <p>Detta är en tv-spelssida och eventuella annonser ska därmed passa det området.</p>
        <p>Vill ni annonsera på N-Forum.se kan ni kontakta oss på <a href="mailto:info@n-forum.se">info@n-forum.se</a> och diskutera det med oss, men observera att vi inte kommer fylla sidan med annonser i onödan.<p>
      </div>
      
      <div class="nf_msg">
        <p><strong>Länkbyten / Länkvänner</strong></p>
        <p>Vill du göra ett länkbyte med våran sida, vad kul. Så länge den berör området TV-Spel. Din tillbakalänk kommer hamna på vår första sida.</p>
        <p>Hör i sådana fall bara av dig till <a href="mailto:info@n-forum.se">info@n-forum.se</a> så ska det gå att lösa. Vi kommer dock inte byta länk med vem som helst utan det avgör vi från fall till fall.</p>
        <p>Under tiden kan du förbereda din webbsida med en länk från "<a href="/avdelning/information/linkback.asp">Länk hit</a>".</p>
      </div>
      
      <div class="nf_msg">
        <p><strong>Övrig information</strong></p>
        <p>N-Forum.se är en inofficiell fansida om Nintendo och ska INTE kopplas ihop med <a href="http://www.bergsala.se">Bergsala</a>/<a href="http://www.nintendo.se">Nintendo</a> som bolag. </p>
        <p>Rättigheterna till spelen, karaktärerna och de avbildade spelfodralen är förbehållna respektive bolag. Inga intrång avsedda. </p>
        <p>N-Forum.se drivs helt på ideell basis av fans. Vill ni kontakta oss når ni oss på <a href="mailto:info@n-forum.se">info@n-forum.se</a>. </p>
      </div>
    </div>
    
    <div class="nf_datablock nf_size_onethird">

    </div> 
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->