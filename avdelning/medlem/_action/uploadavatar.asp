<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%
  If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn)
  
  On Error Resume Next
  Set Upload = Server.CreateObject("aspSmartUpload.SmartUpload")
    Upload.MaxFileSize      = 51200
    Upload.Upload
  
    ' #### UPPLADDNINGSPROCEDUR ####
    
      If Err.Number = -2147220399 Then bErr = True : lErr = 2
      If Not bErr Then
        Set File = Upload.Files.Item(1)
          If Not File.IsMissing Then
            Select Case LCase(File.Fileext)
              Case "jpg", "jpeg", "png", "bmp", "gif"
                sFilename   = "u" & Right("000000" & CLng(CONST_USERID), 6) & "." & File.Fileext
                File.SaveAs config_UpTemp & sFilename
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
    Response.Redirect("../installningar.asp?p=avatar&fail=" & lErr)
  End If
  
  sFile     = Server.MapPath(config_UpTemp) & "\" & sFilename
  sFileSave = Server.MapPath(config_Avatar) & "\" & "u" & Right("000000" & CLng(CONST_USERID), 6) & ".jpg"
  
  ' #### BEHANDLA BILDEN ####
  
    Set Jpeg = Server.CreateObject("Persits.Jpeg")
      Jpeg.Open sFile
      Jpeg.PreserveAspectRatio  = True
      Jpeg.Quality              = 80
      Jpeg.Interpolation        = 2
      
      If Jpeg.OriginalWidth > Jpeg.OriginalHeight Then
        If Jpeg.OriginalWidth > 100 Then
          Jpeg.Width  = 100
          Jpeg.Crop 0, -((100 - Jpeg.Height) / 2), 100, (((100 - Jpeg.Height) / 2) + Jpeg.Height)
        End If
      Else
        If Jpeg.OriginalHeight > 100 Then
          Jpeg.Height  = 100
          Jpeg.Crop -((100 - Jpeg.Width) / 2), 0, (((100 - Jpeg.Width) / 2) + Jpeg.Width), 100
        End If
      End If

      Jpeg.Save sFileSave
    Set Jpeg = Nothing
    
  ' #########################
  
  Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sFile) Then fso.DeleteFile sFile, True
  Set fso = Nothing
  
  RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aID = " & CLng(CONST_USERID), True
    rsDB(1)("aAvatar") = True
    rsDB(1).Update
  RS_Close 1
  
  Session.Value("form_saved") = True
  Call SayMe("Uppladdad","Din <strong>avatar</strong> har nu laddats upp!", "/avdelning/medlem/installningar.asp?p=avatar")

%>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->