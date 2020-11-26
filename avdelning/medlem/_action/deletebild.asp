<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%

    If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
    
    lID  = GetQ("e","ABC",100)

    If HasAcc(CONST_CMS_RIGHTS,"CMS202") Then
      sSQL = "SELECT * FROM cms_AnvBilder WHERE abID = " & CLng(lID)
    Else
      sSQL = "SELECT * FROM cms_AnvBilder WHERE abID = " & CLng(lID) & " AND abUppladdadAv = " & CLng(CONST_USERID)
    End If
    
    RS_Open 1, sSQL, True
      If Not rsDB(1).EOF Then
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
          sFile = "userimg" & Right("000000" & lID, 6) & ".jpg"
          sFullFile = config_UserImageFolder & sFile
          If fso.FileExists(sFullFile) Then fso.DeleteFile sFullFile, True
        Set fso = Nothing
        
        rsDB(1).Delete
      End If
    RS_Close 1
    
    Call SayMe("Raderad","Din <strong>bild</strong> har nu raderats!", "/avdelning/medlem/minabilder.asp")
  
  %>
        
<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->