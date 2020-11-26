<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%

  If Not CONST_LOGIN Then Response.Redirect("hold.asp")
  
  lID       = GetF("e","123",0)   ' Titelns ID
  lPostID   = GetF("id","123",0)  ' ID numret på samlingsposten
  sType     = GetF("tp","ABC",10) ' Typ av objekt
  sDoIt     = LCase(GetF("do","ABC",10)) ' Vad gör efter utförd åtgärd (edit,list,new)
  
  bBox      = GetF("box","CHK",0)
  bObjekt   = GetF("objekt","CHK",0)
  bManual   = GetF("manual","CHK",0)
  bExtra    = GetF("extra","CHK",0)
    
  sOvrigt   = GetF("ovrigt","ABC",255)
  
  Select Case sType
    Case "game"
      SQLSamling    = "SELECT * FROM cms_Bind_Anv_Spel WHERE biAnv = " & CLng(CONST_USERID) & " AND biID = " & CLng(lPostID)
      SQLTitel      = "SELECT * FROM cms_SpelTitlar LEFT JOIN cms_Spel ON cms_Spel.sID = tSpelID WHERE tID = " & CLng(lID)
      sql_ID        = "tSpelID"
      sql_Objekt    = "biMedia"
      sql_Bind      = "biSpelID"
      sql_Konsol    = "sKonsol"
      sql_Extra     = True
    Case "console"
      SQLSamling    = "SELECT * FROM cms_Bind_Anv_Konsol WHERE biAnv = " & CLng(CONST_USERID) & " AND biID = " & CLng(lPostID)
      SQLTitel      = "SELECT * FROM cms_KonsolTitlar LEFT JOIN cms_Konsol ON cms_Konsol.kID = tKonsolID WHERE tID = " & CLng(lID)
      sql_ID        = "tKonsolID"
      sql_Objekt    = "biKonsol"
      sql_Bind      = "biKonsolID"
      sql_Konsol    = "kKonsol"
    Case "addon"
      SQLSamling    = "SELECT * FROM cms_Bind_Anv_Tillbehor WHERE biAnv = " & CLng(CONST_USERID) & " AND biID = " & CLng(lPostID)
      SQLTitel      = "SELECT * FROM cms_TillbehorTitlar LEFT JOIN cms_Tillbehor ON cms_Tillbehor.iID = tTillbehorID WHERE tID = " & CLng(lID)
      sql_ID        = "tTillbehorID"
      sql_Objekt    = "biTillbehor"
      sql_Bind      = "biTillbehorID"
      sql_Konsol    = "iKonsol"
    Case Else
      Response.Redirect("hold.asp")
  End Select
  
  RS_Open 1, SQLTitel, False ' Själva titeln
    If rsDB(1).EOF Then
      Response.Redirect("hold.asp")
    Else
      text_BindID     = rsDB(1)(sql_ID)
      text_TitelID    = rsDB(1)("tID")
      text_TitelName  = rsDB(1)("tTitel")
      If sql_Extra Then text_TitelExtra = rsDB(1)("tExtra")
      text_Region     = rsDB(1)("tRegion")
      text_Konsol     = rsDB(1)(sql_Konsol)
    End If
  RS_Close 1
  
  RS_Open 1, SQLSamling, True ' Posten i samlingsdatabasen
    If rsDB(1).EOF Then
      rsDB(1).AddNew
      rsDB(1)("biAnv")          = CLng(CONST_USERID)
      rsDB(1)(sql_Bind)         = CLng(text_BindID)
      rsDB(1)("biDatumSparad")  = Now 
    End If
    
    rsDB(1)("biTitelID")      = CLng(text_TitelID)
    
    rsDB(1)("biBox")          = bBox
    rsDB(1)(sql_Objekt)       = bObjekt
    rsDB(1)("biManual")       = bManual
    rsDB(1)("biExtra")        = bExtra
    
    rsDB(1)("biOvrigt")       = sOvrigt
    
    ' ## COLLECT DATA ##
    
      ssRegion          = text_Region
      ssGameName        = text_TitelName
      
      If Len(text_TitelExtra) > 1 Then
        ssGameNameCut     = CutText(text_TitelName, 65) & "</a>" & "<span>" & text_TitelExtra & "</span>"
      Else
        ssGameNameCut     = CutText(text_TitelName, 65) & "</a>"
      End If
      
      If bBox Then ssCBox = "blank" Else ssCBox = ""
      If bObjekt Then ssCMedia = "blank" Else ssCMedia = ""
      If bManual Then ssCManual = "blank" Else ssCManual = ""
      If bExtra Then ssCExtra = "blank" Else ssCExtra = ""
    
    ' ##################
      
    rsDB(1).Update
    
    ssDataID = rsDB(1)("biID") 
    
  RS_Close 1

%>

<script type="text/javascript">
  <% Select Case sDoIt %>
    <% Case "edit" %>
      parent.rh_updateRow("titleListed_Clone", "titleListed_Row_" + <% = ssDataID %>, "titleListed_Row_", <% = ssDataID %>, "LI","REGION==<% = ssRegion %>;;GAMEID==<% = lID %>;;GAME==<% = sEncode(ssGameName) %>;;CUTGAME==<% = ssGameNameCut %>;;POSTID==<% = ssDataID %>;;CBOX==<% = ssCBox %>;;CMEDIA==<% = ssCMedia %>;;CMANUAL==<% = ssCManual %>;;CEXTRA==<% = ssCExtra %>;;KONSOL==<% = text_Konsol %>");
      parent.ClosePopBox();
    <% Case "new" %>
      parent.rh_cloneRow("titleListed_Clone", "titleListed_List", "titleListed_Row_", <% = ssDataID %>, "LI","REGION==<% = ssRegion %>;;GAMEID==<% = lID %>;;GAME==<% = sEncode(ssGameName) %>;;CUTGAME==<% = ssGameNameCut %>;;POSTID==<% = ssDataID %>;;CBOX==<% = ssCBox %>;;CMEDIA==<% = ssCMedia %>;;CMANUAL==<% = ssCManual %>;;CEXTRA==<% = ssCExtra %>");
      
      if(parent.CountItemList("titleListed_List") > 0) {
        parent.document.getElementById("titleListed_List").style.display = "block";
        parent.document.getElementById("titleListed_Mess").style.display = "none";
      }
      
      parent.ClosePopBox();
    <% Case "list" %>
      parent.SavedCollection(<% = lID %>);
  <% End Select %>
  
  location.href = "/__AJAX/popbox/_action/hold.asp";
</script>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->