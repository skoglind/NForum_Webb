<%
Function lstKategori(lNo)
  iBound = 18

  Select Case CLng(lNo)
    Case  1   : sData = "Nintendo allmänt"
    Case  2   : sData = "Bärbart"
    Case  3   : sData = "Stationärt"
    Case  4   : sData = "Gameboy (Original)"
    Case  5   : sData = "Gameboy Color"
    Case  6   : sData = "Gameboy Advance"
    Case  7   : sData = "Nintendo DS"
    Case  8   : sData = "Nintendo (Original)"
    Case  9   : sData = "Super Nintendo"
    Case 10   : sData = "Nintendo 64"
    Case 11   : sData = "Nintendo Gamcube"
    Case 12   : sData = "Nintendo Wii"
    Case 13   : sData = "Virtual Boy"
    Case 14   : sData = "Game & Watch"
    Case 15   : sData = "N-Forum.se/Gameboy.nu"
    Case 16   : sData = "Övrigt"
    Case 17   : sData = "Color TV-Game"
    Case 18   : sData = "Nintendo 3DS"
    Case Else : sData = CLng(iBound)
  End Select

  lstKategori = sData
End Function

Function lstKonsol(lNo)
  iBound = 13

  Select Case CLng(lNo)
    Case  1   : sData = "Nintendo Gameboy"
    Case  2   : sData = "Nintendo Gameboy Color"
    Case  3   : sData = "Nintendo Gameboy Advance"
    Case  4   : sData = "Nintendo DS"
    Case  5   : sData = "Nintendo (8-bit)"
    Case  6   : sData = "Super Nintendo"
    Case  7   : sData = "Nintendo 64"
    Case  8   : sData = "Nintendo Gamecube"
    Case  9   : sData = "Nintendo Wii"
    Case 10   : sData = "Nintendo Virtual Boy"
    Case 11   : sData = "Nintendo Game & Watch"
    Case 12   : sData = "Color TV-Game"
    Case 13   : sData = "Nintendo 3DS"
    Case Else : sData = CLng(iBound)
  End Select

  lstKonsol = sData
End Function

Function lstKonsolShort(lNo)
  iBound = 13

  Select Case CLng(lNo)
    Case  1   : sData = "Gameboy"
    Case  2   : sData = "Gameboy Color"
    Case  3   : sData = "Gameboy Advance"
    Case  4   : sData = "Nintendo DS"
    Case  5   : sData = "Nintendo (8-bit)"
    Case  6   : sData = "Super Nintendo"
    Case  7   : sData = "Nintendo 64"
    Case  8   : sData = "Gamecube"
    Case  9   : sData = "Nintendo Wii"
    Case 10   : sData = "Virtual Boy"
    Case 11   : sData = "Game & Watch"
    Case 12   : sData = "Color TV-Game"
    Case 13   : sData = "Nintendo 3DS"
    Case Else : sData = CLng(iBound)
  End Select

  lstKonsolShort = sData
End Function

Function lstKonsolSuperShort(lNo)
  iBound = 13

  Select Case CLng(lNo)
    Case  1   : sData = "GB"
    Case  2   : sData = "GBC"
    Case  3   : sData = "GBA"
    Case  4   : sData = "NDS"
    Case  5   : sData = "NES"
    Case  6   : sData = "SNES"
    Case  7   : sData = "N64"
    Case  8   : sData = "GC"
    Case  9   : sData = "WII"
    Case 10   : sData = "VB"
    Case 11   : sData = "G&W"
    Case 12   : sData = "CTVG"
    Case 13   : sData = "3DS"
    Case Else : sData = CLng(iBound)
  End Select

  lstKonsolSuperShort = sData
End Function

Function lstRegion(lNo)
  iBound = 11

  Select Case CLng(lNo)
    Case  1   : sData = "Europa"
    Case  2   : sData = "USA"
    Case  3   : sData = "Japan"
    Case  4   : sData = "Skandinavien"
    Case  5   : sData = "Australien"
    Case  6   : sData = "Sverige"
    Case  7   : sData = "Tyskland"
    Case  8   : sData = "Kanada"
    Case  9   : sData = "Frankrike"
    Case 10   : sData = "Spanien"
    Case 11   : sData = "Storbritannien"
    Case Else : sData = CLng(iBound)
  End Select

  lstRegion = sData
End Function

Function lstFont(lNo)

  Select Case CLng(lNo)
    Case  1   : sData = "Arial"
    Case  2   : sData = "Verdana"
    Case  3   : sData = "Times New Roman"
    Case  4   : sData = "Courier New"
    Case Else : sData = "Arial"
  End Select

  lstFont = sData
End Function

Function lstKSTyp(lNo)
  iBound = 5

  Select Case CLng(lNo)
    Case  1   : sData = "Säljes"
    Case  2   : sData = "Köpes"
    Case  3   : sData = "Bytes"
    Case  4   : sData = "Auktion"
    Case  5   : sData = "Skänkes"
    Case Else : sData = CLng(iBound)
  End Select

  lstKSTyp = sData
End Function

Function lstImgSize(lNo)
  iBound = 18

  Select Case CLng(lNo)
    Case  1   : sData = "28,28"
    Case  2   : sData = "50,45"
    Case  3   : sData = "50,50"
    Case  4   : sData = "80,80"
    Case  5   : sData = "100,100"
    Case  6   : sData = "150,150"
    Case  7   : sData = "320,240"
    Case  8   : sData = "640,480"
    Case  9   : sData = "800,600"
    Case 10   : sData = "NO_LOGIN_1024,768"
    Case 11   : sData = "NO_LOGIN_1280,1024"
    Case 12   : sData = "200,150"
    Case 13   : sData = "300,300"
    Case 14   : sData = "400,400"
    Case 15   : sData = "500,500"
    Case 16   : sData = "180,120"
    Case 17   : sData = "301,200"
    Case 18   : sData = "23,23"
    Case Else : sData = CLng(iBound)
  End Select

  lstImgSize = sData
End Function

Function lstKSKategori(lNo)
  iBound = 5

  Select Case CLng(lNo)
    Case  0   : sData = "Allt"
    Case  1   : sData = "Spel"
    Case  2   : sData = "Konsoler"
    Case  3   : sData = "Tillbehör"
    Case  4   : sData = "Tidningar"
    Case  5   : sData = "Övrigt"
    Case Else : sData = CLng(iBound)
  End Select

  lstKSKategori = sData
End Function

Function lstKSUnderKategori(lNo)
  iBound = 1

  Select Case CLng(lNo)
    Case  0   : sData = "Allt"
    Case  1   : sData = "Övrigt"
    Case Else : sData = CLng(iBound)
  End Select

  lstKSUnderKategori = sData
End Function

Function lstAvdelning(lNo)
  iBound = 2

  Select Case CLng(lNo)
    Case  0   : sData = "nyheter;nyheter_visa;cms_Nyheter;nID"
    Case  1   : sData = "recensioner;recension_visa;cms_Recensioner;rID"
    Case  2   : sData = "artiklar;artikel_visa;cms_Artiklar;aaID"
    Case Else : sData = CLng(iBound)
  End Select

  lstAvdelning = sData
End Function
%>