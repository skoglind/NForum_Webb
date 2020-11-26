<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%
  If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
  
  lAntalBilder = Con.ExeCute("SELECT COUNT(abID) FROM cms_AnvBilder WHERE abUppladdadAv = " & CLng(CONST_USERID))(0)
  If lAntalBilder >= config_UserMaxImages Then Response.Redirect("../minabilder.asp?fail=4")
  
  If config_UserImagesDays > CONST_DAYSMEMBER Then Response.Redirect("../minabilder.asp?fail=5")
  
  On Error Resume Next
  Set Upload = Server.CreateObject("aspSmartUpload.SmartUpload")
    Upload.MaxFileSize      = 512000
    Upload.Upload
  
    ' #### UPPLADDNINGSPROCEDUR ####
    
      If Err.Number = -2147220399 Then bErr = True : lErr = 2
      If Not bErr Then
        Set File = Upload.Files.Item(1)
          If Not File.IsMissing Then
            Select Case LCase(File.Fileext)
              Case "jpg", "jpeg", "png", "bmp", "gif"
                sFilename   = "userimg" & Right("000000" & CLng(CONST_USERID), 6) & "_" & Timer & "." & File.Fileext
                File.SaveAs config_UpTemp & sFilename
                
                sOldFileName = File.FileName
                sOldFileType = File.FileExt
              Case Else
                bErr = True : lErr = 3
            End Select 
          Else
            bErr = True : lErr = 1
          End If
        Set File = Nothing
      End If
      
    ' ##############################
  
  Set Upload = Nothing
  On Error Goto 0
  
  If bErr Then
    Response.Redirect("../minabilder.asp?fail=" & lErr)
  End If
  
  RS_Open 1, "SELECT * FROM cms_AnvBilder WHERE 1 = 2", True
    rsDB(1).AddNew
      rsDB(1)("abTitel")          = sOldFileName
      rsDB(1)("abOriginalNamn")   = sOldFileName
      rsDB(1)("abTyp")            = sOldFileType
      rsDB(1)("abUppladdadAv")    = CONST_USERID
      rsDB(1)("abUppladdadDatum") = Now
      rsDB(1)("abInSizes")        = ","
    rsDB(1).Update
    
    lNewID = rsDB(1)("abID")
  RS_Close 1
  
  sFile     = Server.MapPath(config_UpTemp) & "\" & sFilename
  sFileSave = config_UserImageFolder & "userimg" & Right("000000" & CLng(lNewID), 6) & ".jpg"
  
  ' #### BEHANDLA BILDEN ####
  
    Set Jpeg = Server.CreateObject("Persits.Jpeg")
      Jpeg.Open sFile
      Jpeg.PreserveAspectRatio  = True
      Jpeg.Quality              = 80
      Jpeg.Interpolation        = 10
      jpeg.Canvas.Brush.Color   = &HFFFFFF
      
      If Jpeg.OriginalWidth > Jpeg.OriginalHeight Then
        If Jpeg.OriginalWidth > 500 Then
          Jpeg.Width  = 500
          Jpeg.Crop 0, -((500 - Jpeg.Height) / 2), 500, (((500 - Jpeg.Height) / 2) + Jpeg.Height)
        End If
      Else
        If Jpeg.OriginalHeight > 500 Then
          Jpeg.Height  = 500
          Jpeg.Crop -((500 - Jpeg.Width) / 2), 0, (((500 - Jpeg.Width) / 2) + Jpeg.Width), 500
        End If
      End If

      Jpeg.Save sFileSave
    Set Jpeg = Nothing
    
  ' #########################
  
  Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sFile) Then fso.DeleteFile sFile, True
  Set fso = Nothing
  
  'RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(CONST_USERID), True
  '  rsDB(1)("aAvatar") = True
  '  rsDB(1).Update
  'RS_Close 1
  
  Session.Value("form_saved") = True
  Call SayMe("Uppladdad","Din <strong>bild</strong> har nu laddats upp!", "/avdelning/medlem/minabilder.asp")

%>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->