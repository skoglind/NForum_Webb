<!--#INCLUDE FILE="../__INC/includes.asp"-->

  <%
    If CONST_LOGIN AND CONST_SET_QUICK = 2 Then
    
      bDoListing = False
    
      lAnvID    = CLng(CONST_USERID)
      lTitelID  = GetF("e","123",0)
      
      RS_Open 1, "SELECT * FROM cms_Speltitlar WHERE tID = " & lTitelID, False
        If Not rsDB(1).EOF Then
          bDoListing = True
          
          lSpelID = CLng(rsDB(1)("tSpelID"))
        End If
      RS_Close 1
      
      RS_Open 1, "SELECT * FROM cms_Bind_Anv_Spel WHERE biAnv = " & lAnvID & " AND biTitelID = " & lTitelID, True
        If rsDB(1).EOF AND bDoListing Then
          rsDB(1).AddNew
            
            rsDB(1)("biAnv")          = lAnvID
            rsDB(1)("biSpelID")       = lSpelID
            rsDB(1)("biDatumSparad")  = Now
            
            rsDB(1)("biTitelID")      = lTitelID
      
            rsDB(1)("biBox")          = False
            rsDB(1)("biMedia")        = True
            rsDB(1)("biManual")       = False
            rsDB(1)("biExtra")        = False
            
            rsDB(1)("biBox_Grade")    = 0
            rsDB(1)("biMedia_Grade")  = 0
            rsDB(1)("biManual_Grade") = 0
            rsDB(1)("biExtra_Grade")  = 0
            
            rsDB(1)("biOvrigt")       = "*Q-Listad*"
          
          rsDB(1).Update
        End if
      RS_Close 1
    
    End If
  %>

<!--#INCLUDE FILE="../__INC/includes_end.asp"-->