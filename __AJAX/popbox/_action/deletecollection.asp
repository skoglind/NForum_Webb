<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%

  lID     = GetQ("e","123",0)
  sType   = GetQ("tp","ABC",10) ' Typ av objekt

  Select Case sType
    Case "game"
      SQLTable      = "cms_Bind_Anv_Spel"
    Case "console"
      SQLTable      = "cms_Bind_Anv_Konsol"
    Case "addon"
      SQLTable      = "cms_Bind_Anv_Tillbehor"
    Case Else
      bErr = True
  End Select

  If Not CONST_LOGIN Then bErr = True
  
  If Not bErr Then
    RS_Open 1, "SELECT * FROM " & SQLTable & " WHERE biAnv = " & CLng(CONST_USERID) & " AND biID = " & CLng(lID), True
      
      If Not rsDB(1).EOF Then 
        lTitelID = rsDB(1)("biTitelID")
        rsDB(1).Delete
      End If
      
    RS_Close 1
    
    Response.Write("1")
  Else
    Response.Write("0") 
  End If

%>