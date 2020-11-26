<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<%
If Session.Value("SET_AllowTimer") Then timerStart = Timer
%>

<%
Randomize
rndID = CLng(Rnd*99999999) + 1

If Len(page_Title) > 1 Then
  raw_Title = page_Title & "."
  page_Title = page_Title & " / " & page_Name
Else
  page_Title = page_Name
End if
%>

<html>
  <head>
    <title> <% = page_Title %>  </title>
    <meta http-equiv="content-type" content="text/html; CHARSET=ISO-8859-1">
    <meta http-equiv="content-language" content="sv">
    
    <META http-equiv="PICS-Label" content='(PICS-1.1 "http://www.classify.org/safesurf/" l gen t for "http://www.n-forum.se/nintendo/" r (SS~~000 1))'>
     
    <% If Len(page_description) > 0 Then %>
      <meta name="Description" content="<% = page_description %>">
    <% Else %>
      <meta name="Description" content="Lista över alla Nintendos spel, konsoler och tillbehör med boxart. Forum och möjlighet att lista dina spel. Köp & Sälj dina spel i annonsavdelningen.">
    <% End if %>
    <meta name="Keywords" content="<% = page_keywords %>nintendo, forum, spel, databas, gameboy">
    
    <link rel="stylesheet" type="text/css" href="/__CSS/nforum_new.css">
    
    <link rel="stylesheet" href="/__CSS/lightbox.css" type="text/css" media="screen">
    
    <!--[if IE 6]>
      <link rel="stylesheet" type="text/css" href="/__CSS/internet_explorer_6.css">
      <script src="/__JS/png/DD_belatedPNG.js"></script>
      <script>
          DD_belatedPNG.fix('.png_bg, img.png_img');
      </script>
    <![endif]-->
    
    <!--[if IE 7]>
      <link rel="stylesheet" type="text/css" href="/__CSS/internet_explorer_7.css">
    <![endif]-->
    
    <script type="text/javascript" src="/__JS/nforum.js"></script>
    
    <script type="text/javascript" src="/__JS/prototype.js"></script>
    <script type="text/javascript" src="/__JS/scriptaculous.js?load=effects,builder"></script>
    <script type="text/javascript" src="/__JS/lightbox.js"></script>

  </head>
  <body>
  
    <div id="NFORUM">
    <div id="YAH_Inner">
      
      <div class="nf_userbar">
        <% If CONST_LOGIN Then %>
          <ul>
            <li> <a href="/avdelning/medlem/">Profil (<% = CONST_USERNAME %>)</a></li>
            <li> <p>|</p> </li>
            <li> <a href="/avdelning/medlem/inkorg.asp" id="anPM">PM (<% = antalNyaPM %>)</a> </li>
            <li> <p>|</p> </li>
            <li> <a href="/avdelning/medlem/minainlagg.asp">Foruminlägg</a> </li>
            <li> <p>|</p> </li>
            <li> <a href="/avdelning/medlem/minaspel.asp">Spellista</a> </li>
            <li> <p>|</p> </li>
            <li> <a href="/avdelning/annonser/minaannonser.asp">Annonser</a> </li>
          </ul>
          <ul style="float: right;">
            <li> <a href="/avdelning/listor/online.asp" id="anOn">Online (<% = antalOnline %>)</a> </li>
            <li> <p>|</p> </li>
            <li> <a href="/avdelning/medlem/installningar.asp">Inställningar</a> </li>
            <li> <p>|</p> </li>
            <li> <a href="/_action/do_logout.asp">Logga ut</a> </li>
          </ul>
        <% Else %>
          <div class="nf_loginbar">
            <form method="POST" action="/_action/do_login.asp">
              <div>
                <input class="text" type="text" name="r" id="anvandarnamn" value="Användarnamn" onfocus="clearField(this,'Användarnamn');" onblur="retypeField(this,'Användarnamn');">
                <input class="text" type="password" name="g" id="losenord" value="Lösenord" onfocus="clearField(this,'Lösenord');" onblur="retypeField(this,'Lösenord');">
                <input class="chk" type="checkbox" name="s" value="YES" id="top_remme"> <label for="top_remme">Kom ihåg!</label>
                <input type="hidden" name="postback" value="<% = ActivePage %>">
                <input class="btn" type="submit" value="Logga in">
              </div>
            </form>
          </div>
          <ul style="float: right;">
            <li> <a href="/avdelning/medlem/registreradig.asp">Bli medlem GRATIS!</a> </li>
            <li> <p>|</p> </li>
            <li> <a href="/avdelning/medlem/glomtlosen.asp">Glömt lösenordet?</a> </li>
            <li> <p>|</p> </li>
            <li> <a href="/avdelning/medlem/loggain.asp">Logga in</a> </li>
          </ul>
        <% End If %>
      </div>
      
      <div class="nf_logo">
        <a class="nf_logoA" href="http://www.n-forum.se/nintendo/" title="N-Forum.se - Nintendo Forum"></a>
      
        <div class="nf_searchbox">
          <% remQ = GetQ("q","ABC",0) %>
          <% remQ = MakeLegal(remQ) %>
          <% remQ = sEncode(remQ) %>
               
          <form onsubmit="do_search(document.getElementById('search_section').value, document.getElementById('search_query').value,<% = config_MinSearch %>); return false; " action="">
            <div>
              <select id="search_section">
                <option disabled style="font-weight: bold;">Hela sidan</option>
                <option <% If page_Slide = "forum" Then Response.Write(" selected") %> value="forum/sokforum.asp?q=">&nbsp;&nbsp; Forum</option>
                <% If CONST_LOGIN Then %><option <% If page_Slide = "medlem" Then Response.Write(" selected") %> value="listor/sokmedlem.asp?q=">&nbsp;&nbsp; Medlem</option><% End If %>
                <option <% If page_Slide = "annonser" Then Response.Write(" selected") %> value="annonser/?q=">&nbsp;&nbsp; Annonser</option>
                <option disabled style="font-weight: bold;">&nbsp;&nbsp; Databasen</option>
                <option <% If page_Slide = "spel" Then Response.Write(" selected") %> value="spel/sokspel.asp?q=">&nbsp;&nbsp;&nbsp;&nbsp; Spel</option>
                <option <% If page_Slide = "konsoler" Then Response.Write(" selected") %> value="konsol/sokkonsol.asp?q=">&nbsp;&nbsp;&nbsp;&nbsp; Konsoler</option>
                <option <% If page_Slide = "tillbehor" Then Response.Write(" selected") %> value="tillbehor/soktillbehor.asp?q=">&nbsp;&nbsp;&nbsp;&nbsp; Tillbehör</option>
                <option disabled style="font-weight: bold;">&nbsp;&nbsp; Texter</option>
                <option <% If page_Slide = "nyheter" Then Response.Write(" selected") %> value="nyheter/?q=">&nbsp;&nbsp;&nbsp;&nbsp; Nyheter</option>
                <option <% If page_Slide = "recensioner" Then Response.Write(" selected") %> value="recensioner/?q=">&nbsp;&nbsp;&nbsp;&nbsp; Recensioner</option>
                <option <% If page_Slide = "artiklar" Then Response.Write(" selected") %> value="artiklar/?q=">&nbsp;&nbsp;&nbsp;&nbsp; Artiklar</option>
              </select>
              
              <input type="text" class="sbox" id="search_query" maxlength=100 value="<% = remQ %>">
              <input type="submit" value="Sök" class="sbut">
            </div>
          </form>

        </div>
      </div>
      
      <div class="nf_menu">
        <ul>
          <li <% If page_SelMenu = "home"       Then Response.Write(" class='c'") %>> <a href="/nintendo/" title="Gå till första sidan">Första Sidan</a> </li>
          <li <% If page_SelMenu = "forum"      Then Response.Write(" class='c'") %>> <a href="/avdelning/forum/default.asp" title="Gå till forumet">Forum</a> </li>
          <li <% If page_SelMenu = "user"       Then Response.Write(" class='c'") %>> <a href="/avdelning/medlem/default.asp" title="Gå till medlemssidan">Medlem</a> </li>
          <!-- <li <% If page_SelMenu = "blog"       Then Response.Write(" class='c'") %>> <a href="/avdelning/blog/default.asp" title="Gå till bloggarna">Bloggar</a> </li> -->
          <li <% If page_SelMenu = "texter"     Then Response.Write(" class='c'") %>> <a href="/avdelning/texter/default.asp" title="Gå till texter">Texter</a> </li>
          <li <% If page_SelMenu = "databas"    Then Response.Write(" class='c'") %>> <a href="/avdelning/databas/default.asp" title="Gå till databas">Databas</a> </li>
          <li <% If page_SelMenu = "buy"        Then Response.Write(" class='c'") %>> <a href="/avdelning/annonser/default.asp" title="Gå till annonser">Annonser</a> </li>
        </ul>
      </div>
      
      <div class="nf_submenu">