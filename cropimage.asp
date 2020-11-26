<% measureStart = Timer %>

<!--#INCLUDE FILE="__INC/includes.asp"-->

<%
Response.Buffer = True
Response.Expires = 0

If config_LockDown_Bilder Then
  Response.Write("Bildvisning temporärt nedstängt av N-Forum.se Admin!")
  Response.End
End If

lID    = GetQ("e","123",0)
bCropP = GetQ("top","CHK",0)

  Con_Open

  sErrorFile = Server.MapPath("/design/noimg.png")
  
    sOriginal = ImgOriginal(lID)
      
    If sOriginal = "NO_IMG" Then
      cSTREAMFILE = sErrorFile
    Else  
      cSTREAMFILE = config_ImageFolder & sOriginal
    End If
    
  Con_Close
  
  Response.Clear
  
  Set Jpeg = Server.CreateObject("Persits.Jpeg")
    Jpeg.Open(cSTREAMFILE)
    Jpeg.PreserveAspectRatio = True
    Jpeg.PNGOutput = True
    
    If Jpeg.Width <> 321 Then Jpeg.Width = 321
    If Jpeg.Height < 200 Then Jpeg.Height = 200
    
    If Not bCropP Then
      boxHeight = CLng((Jpeg.Height - 200))
      
      upHeight = CLng(boxHeight / 2)
      dwHeight = CLng(Jpeg.Height - upHeight) + 1
    Else
      upHeight = 0
      dwHeight = 200
    End IF
    
    Jpeg.Crop 10, upHeight, Jpeg.Width - 10, dwHeight
    
    Jpeg.SendBinary
  Set Jpeg = Nothing
%>

<!--#INCLUDE FILE="__INC/includes_end.asp"-->