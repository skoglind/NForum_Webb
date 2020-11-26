<!--#INCLUDE FILE="../__INC/includes.asp"-->

  <%
  
    lastAttempt = Session.Value("Return_Last")
    If lastAttempt = Empty Then
      lastAttempt = CDate(Now - #00:00:35#)
    End If
  
    If CONST_LOGIN Then
      lLogOut = Con.ExeCute("SELECT aLOCK FROM fsBB_Anv WHERE aID = " & CLng(CONST_USERID))(0)

      If lLogOut = True Then
        Con.ExeCute("UPDATE fsBB_Anv SET aLOCK = 0, aTimeStamp = '" & DateAdd("n", -6, Now) & "' WHERE aID = " & CLng(CONST_USERID))
        Session.Abandon
        Response.Write("L:0:0")
        Response.End
      End If
    End If
    
    If Now < CDate(lastAttempt + #00:00:30#) Then
      lNoOn = CLng(Session.Value("Return_Online"))
      lNoPM = CLng(Session.Value("Return_PM"))
    
      Response.Write("O:" & lNoOn & ":" & lNoPM)
    Else
      If CONST_LOGIN Then
        Con.ExeCute("UPDATE fsBB_Anv SET aTimeStamp = '" & Now & "' WHERE aID = " & CLng(CONST_USERID))
        lNoPM   = Con.ExeCute("SELECT COUNT(*) FROM fsBB_PM WHERE pLast = 0 AND pTill = " & CLng(CONST_USERID))(0)
        lNoOn   = Con.ExeCute("SELECT COUNT(*) FROM fsBB_Anv WHERE aTimeStamp > '" & DateAdd("n", -5, Now) & "' AND aBlockadTill < '" & Date & "' AND aAktiverad = 1")(0)
        
        Session.Value("Return_Last")    = Now
        Session.Value("Return_Online")  = CLng(lNoOn)
        Session.Value("Return_PM")      = CLng(lNoPM)
        
        Response.Write("N:" & lNoOn & ":" & lNoPM)
      Else
        Response.Write("F:0:0")
      End If
    End If
  %>

<!--#INCLUDE FILE="../__INC/includes_end.asp"-->