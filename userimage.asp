<!--#INCLUDE FILE="__INC/includes.asp"-->

<%
Response.Buffer = True
Response.Expires = 0

If config_LockDown_Bilder Then
  Response.Write("Bildvisning temporärt nedstängt av N-Forum.se Admin!")
  Response.End
End If

lID = GetQ("e","123",0)

  Con_Open

  sErrorFile = Server.MapPath("../gfx/noimg.png")
  
  sFile = "userimg" & Right("000000" & lID, 6) & ".jpg"
  sFullFile = config_UserImageFolder & sFile
  
  Set fso = Server.CreateObject("Scripting.FileSystemObject")
    bHasOriginal  = True
    If Not fso.FileExists(sFullFile) Then bHasOriginal = False
  Set fso = Nothing
  
  If bHasOriginal AND CONST_LOGIN Then 
    cSTREAMFILE = sFullFile
  Else
    cSTREAMFILE = sErrorFile
  End If
    
  Con_Close
  
  Response.Clear
  
  If sEr = "no" Then If cSTREAMFILE = sErrorFile Then Response.End

  'Const adTypeBinary = 1 
  'cCONTENTTYPE = "image/jpeg"
  
  'Response.Contenttype = cCONTENTTYPE
    
  'Set Stream = server.CreateObject("ADODB.Stream") 
  'Stream.Type = adTypeBinary 
  'Stream.Open 
  'Stream.LoadFromFile cSTREAMFILE
  'While Not Stream.EOS 
  '  Response.BinaryWrite Stream.Read(1024 * 64) 
  'Wend 
  'Stream.Close 
  'Set Stream = Nothing 
    
  'Response.Flush 
  'Response.End
  
  Set Jpeg = Server.CreateObject("Persits.Jpeg")
    Jpeg.Open(cSTREAMFILE)
    
  '  measureStop = Timer
  '  measure = FormatNumber(measureStop - measureStart, 3)
    
  '  Select Case Trim(LCase(sF))
  '    Case "grayscale"
  '      Jpeg.Grayscale 1
  '    Case "timer"
  '      Jpeg.Canvas.Font.Color  = &H000000
  '      Jpeg.Canvas.Font.Size   = 12
  '      Jpeg.Canvas.PrintTextEx measure & " seconds", 2, 12, "c:\Windows\Fonts\Arial.ttf"
  '    Case "sharpen"
  '      Jpeg.Interpolation = 2
  '  End Select
    
    'If lW > 600 Then
    '  Jpeg.Canvas.DrawPNG Jpeg.Width - 340, Jpeg.Height - 115, Server.MapPath("../bilder") & "\stamp_large.png"
    'ElseIf lW > 150 Then
    '  Jpeg.Canvas.DrawPNG Jpeg.Width - 140, Jpeg.Height - 44, Server.MapPath("../bilder") & "\stamp.png"
    'End If
    
  Jpeg.SendBinary
  Set Jpeg = Nothing
%>

<!--#INCLUDE FILE="__INC/includes_end.asp"-->