<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%

  If Not CONST_LOGIN Then Response.Redirect("_login.asp")

  lID       = GetQ("e","123",0)   ' Titelns ID
  lPostID   = GetQ("id","123",0)  ' ID numret på samlingsposten
  sType     = GetQ("tp","ABC",10) ' Typ av objekt
  sDoIt     = GetQ("do","ABC",10) ' Vad gör efter utförd åtgärd (edit,list,new)
  
  Select Case sType
    Case "game"
      text_Titel    = "Spelsamling"
      SQLSamling    = "SELECT * FROM cms_Bind_Anv_Spel WHERE biAnv = " & CLng(CONST_USERID) & " AND biID = " & CLng(lPostID)
      SQLTitel      = "SELECT * FROM cms_SpelTitlar WHERE tID = " & CLng(lID)
      SQLRegioner   = "SELECT tID, tRegion, tTitel, rShort FROM cms_SpelTitlar LEFT JOIN cms_Region ON cms_Region.rID = tRegion WHERE tSpelID IN(SELECT tSpelID FROM cms_SpelTitlar WHERE tID = " & CLng(lID) & ") ORDER BY tRelease ASC"
      sql_ID        = "tSpelID"
      sql_Objekt    = "biMedia"
      name_Objekt   = "Media"
    Case "console"
      text_Titel    = "Konsolsamling"
      SQLSamling    = "SELECT * FROM cms_Bind_Anv_Konsol WHERE biAnv = " & CLng(CONST_USERID) & " AND biID = " & CLng(lPostID)
      SQLTitel      = "SELECT * FROM cms_KonsolTitlar WHERE tID = " & CLng(lID)
      SQLRegioner   = "SELECT tID, tRegion, tTitel, rShort FROM cms_KonsolTitlar LEFT JOIN cms_Region ON cms_Region.rID = tRegion WHERE tKonsolID IN(SELECT tKonsolID FROM cms_KonsolTitlar WHERE tID = " & CLng(lID) & ") ORDER BY tRelease ASC"
      sql_ID        = "tKonsolID"
      sql_Objekt    = "biKonsol"
      name_Objekt   = "Konsol"
    Case "addon"
      text_Titel    = "Tillbehörssamling"
      SQLSamling    = "SELECT * FROM cms_Bind_Anv_Tillbehor WHERE biAnv = " & CLng(CONST_USERID) & " AND biID = " & CLng(lPostID)
      SQLTitel      = "SELECT * FROM cms_TillbehorTitlar WHERE tID = " & CLng(lID)
      SQLRegioner   = "SELECT tID, tRegion, tTitel, rShort FROM cms_TillbehorTitlar LEFT JOIN cms_Region ON cms_Region.rID = tRegion WHERE tTillbehorID IN(SELECT tTillbehorID FROM cms_TillbehorTitlar WHERE tID = " & CLng(lID) & ") ORDER BY tRelease ASC"
      sql_ID        = "tTillbehorID"
      sql_Objekt    = "biTillbehor"
      name_Objekt   = "Tillbehör"
    Case Else
      Response.Redirect("_err.asp")
  End Select
  
  RS_Open 1, SQLTitel, False ' Själva titeln
    If rsDB(1).EOF Then
      Response.Redirect("_err.asp")
    Else
      text_ObjektName = rsDB(1)("tTitel")
      text_BindID     = rsDB(1)(sql_ID)
      text_TitelID    = rsDB(1)("tID")
    End If
  RS_Close 1
  
  RS_Open 1, SQLSamling, False ' Posten i samlingsdatabasen
    If rsDB(1).EOF Then
      text_ID       = 0
      text_Box      = False
      text_Objekt   = True
      text_Manual   = False
      text_Extra    = False
      text_SelReg   = lID
    Else
      text_ID       = CLng(rsDB(1)("biID"))
      text_Box      = rsDB(1)("biBox")
      text_Objekt   = rsDB(1)(sql_Objekt)
      text_Manual   = rsDB(1)("biManual")
      text_Extra    = rsDB(1)("biExtra")
      text_SelReg   = rsDB(1)("biTitelID")
      
      text_Ovrigt   = sEncode(rsDB(1)("biOvrigt"))
    End If
  RS_Close 1
  
  RS_Open 1, SQLRegioner, False ' Alla olika regioner/titlar för objektet
    If rsDB(1).EOF Then
      any_Region = False
    Else
      any_Region = True
      list_Regions = rsDB(1).GetRows
    End If
  RS_Close 1

%>

<h3><% = sEncode(text_Titel) %></h3>
<h4><% = sEncode(text_ObjektName) %></h4>

<div class="popBox_Inner_Split"></div>

<form method="POST" id="popForm" target="popBox_Frame" action="/__AJAX/popbox/_action/savecollection.asp">
  <label>Titel</label>
  <select name="e">
    <% For zx = 0 To UBound(list_Regions, 2) %>
      <option value="<% = list_Regions(0, zx) %>" <% If CLng(text_SelReg) = CLng(list_Regions(0, zx)) Then Response.Write(" selected") %>> <% = sEncode(list_Regions(3, zx)) %> | <% = sEncode(list_Regions(2, zx)) %></option>
    <% Next %>
  </select>
  
  <div class="popBox_Inner_Split"></div>
  
  <label>Box</label><input type="checkbox" class="chk" value="YES" name="box" <% If text_Box Then Response.Write(" checked") %>>
  <label><% = sEncode(name_Objekt) %></label><input type="checkbox" class="chk" value="YES" name="objekt" <% If text_Objekt Then Response.Write(" checked") %>>
  <label>Manual</label><input type="checkbox" class="chk" value="YES" name="manual" <% If text_Manual Then Response.Write(" checked") %>>
  <label>Extras</label><input type="checkbox" class="chk" value="YES" name="extra" <% If text_Extra Then Response.Write(" checked") %>>
  
  <div class="popBox_Inner_Split"></div>
  
  <label>&Ouml;vrigt</label>
  <input class="text" type="text" name="ovrigt" value="<% = sEncode(text_Ovrigt) %>" maxlength=255>
  
  <input type="hidden" name="id" value="<% = text_ID %>">
  <input type="hidden" name="tp" value="<% = sType %>">
  <input type="hidden" name="do" value="<% = sDoIt %>">
</form>

<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->