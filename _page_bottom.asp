          <% If Not LCase(page_SelMenu) = "forum" Then %>
            <div class="content" style="margin-top: 15px; min-height: 0px; border-top: dotted 2px #DDD;">
              <div class="nf_datablock nf_size_full">
                <h1 style="text-align: center;">Nintendo Forum / <a href="http://www.n-forum.se">N-Forum.se</a></h1>
                <h2 style="text-align: center; margin-bottom: 10px;">Nintendo när det är som bäst.</h2>
              </div>
            
              <div class="nf_datablock nf_size_full">
                <div class="nf_msg_full">
                  <p style="text-align: center;"><strong>Vi rekommenderar dig att besöka:</strong></p>
                  <p style="text-align: center;">
                    <a href="http://www.nesdb.se" target="_blank" title="Nintendo Entertainment System Database"><img src="http://grafik.n-forum.se/link/nesdb.gif" alt="NESdb"></a>
                    <a href="http://www.mega-man.se" target="_blank" title="Mega Man Evolution"><img src="http://grafik.n-forum.se/link/banner_blue.gif" alt="Mega-Man.se"></a>
                  </p>
                </div>
              </div>
          
            </div>
          <% End If %>
           
        </div>
      </div>
      
      <div id="bottom" class="png_bg">
        <div class="left">
          <strong><a href="http://www.n-forum.se">N-Forum.se</a></strong> är en inofficiell <strong>fansida</strong> om <strong>Nintendo</strong> och ska <strong>INTE</strong> kopplas ihop med <strong><a href="http://www.bergsala.se" target="_blank" rel="nofollow">Bergsala</a>/<a href="http://www.nintendo.se" target="_blank" rel="nofollow">Nintendo</a></strong> som bolag.<br>
          Rättigheterna till spelen, karaktärerna och de avbildade spelfodralen är förbehållna respektive bolag. Inga intrång avsedda.<br>
          N-Forum.se drivs helt på ideell basis av fans. Vill ni kontakta oss når ni oss på <a href="mailto:info@n-forum.se">info@n-forum.se</a>.
        </div>
      
        <div class="right">
          <ul>
            <li> <a href="/avdelning/information/cookies.asp" title="">Cookies</a> </li>
            <li> <a href="/avdelning/information/faq.asp" title="">F.A.Q.</a> </li>
            <li> <a href="/avdelning/information/linkback.asp" title="">Länka hit</a> </li>
            <li> <a href="/avdelning/information/information.asp" title="">Information</a> </li>
          </ul>
        </div>
      </div>
    </div>
    </div>
  
    <div id="popBox" class="popBox">
      <div class="popBox_Inner" id="popBox_Inner"> </div>
      <div class="popBox_Buttons">
        <input id="popBox_BT" type="button" value="Spara" disabled  onclick="OK_PopBox();">
        <input type="button" value="Avbryt" style="color: #A00;" onclick="ClosePopBox();">
      </div>
      <iframe class="popBox_Frame" name="popBox_Frame" id="popBox_Frame" src="/__AJAX/popbox/_action/hold.asp"></iframe>
    </div>
    
    <div id="jsFrameBox" class="jsFrameBox">
      <div class="FrameBox_Title" id="FrameBox_Title">Tjipp</div>
      <iframe id="FrameBox_Frame"></iframe>
      <div class="FrameBox_Buttons">
        <input type="button" onclick="submitFrameBox();" value="Verkställ" style="font-weight: bold;">
        <input type="button" onclick="document.getElementById('jsFrameBox').style.display='none';" value="Avbryt" style="color: #C00;">
      </div>
    </div>
    
    <div id="enlargescreen">
      <img src="<% = config_GFXLocation %>img/loading.png" id="enlargescreen_image" alt="Laddar" title="">
    </div>
    
    <% If CONST_LOGIN Then %><script type="text/javascript">KeepOnline();</script><% End If %>
  
    <%
      If Session.Value("SET_AllowTimer") Then 
        timerStop = Timer
        allTime   = FormatNumber(timerStop - timerStart, 6)
        
        Response.Write "<div class='timefly'>T: " & allTime & "</div>"
      End If
    %>
    
    <!-- Start of StatCounter Code -->
    <script type="text/javascript">
    var sc_project=6296423; 
    var sc_invisible=1; 
    var sc_security="1386f5f4"; 
    </script>
    
    <script type="text/javascript"
    src="http://www.statcounter.com/counter/counter.js"></script><noscript><div
    class="statcounter"><a title="hit counter joomla"
    href="http://statcounter.com/joomla/" target="_blank"><img
    class="statcounter"
    src="http://c.statcounter.com/6296423/0/1386f5f4/1/"
    alt="hit counter joomla" ></a></div></noscript>
    <!-- End of StatCounter Code -->
    
    <script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
    </script>
    <script type="text/javascript">
    _uacct = "UA-911061-2";
    urchinTracker();
    </script>
  </body>
</html>