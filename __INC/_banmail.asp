<%
Function MailIsValid(sSearch)
  Dim sM, bStatus, lExist
  sM = Trim(sSearch)
  bStatus = False
  
  If Len(sM) < 6 Then 
    bStatus = True
  Else
    If CLng(InStr(1, sM, "@", vbTextCompare)) = 0 Then bStatus = True
    If CLng(InStr(1, sM, ".", vbTextCompare)) = 0 Then bStatus = True
    If Not bStatus Then If MailIsInBanList(sM) Then bStatus = True
  End If
  
  MailIsValid = bStatus
End Function

Function MailIsInBanList(sSearch)
  Dim sM, bStatus, dat_file, sMAILS, sNoder, sNod
  sM = LCase(Trim(sSearch))
  bStatus = False
  
  ' #### LÄS IN DAT-FILEN ####
  dat_file = "_mail.dat"
  Set fso = Server.Createobject("scripting.FileSystemObject")
    Set dat_get = fso.OpenTextFile(Server.MapPath("/__INC/dat/" & dat_file))
      sMAILS = dat_get.ReadAll
      dat_get.Close
  Set fso = Nothing
  ' ##########################
  
  ' #### SÖK IGENOM DAT-FILEN EFTER TRÄFFAR ####
  sNoder = Split(sMAILS, Chr(13))
  For Each sNod In sNoder
    sNod = Replace(LCase(Trim(sNod & " ")), Chr(10), "")
    If Len(sNod) > 4 Then
      If Right(sM, Len(sNod)) = sNod Then
        bStatus = True
        Exit For
      End If
    End If
  Next
  ' ############################################
  
  MailIsInBanList = bStatus
End Function
%>