<%
  ' #### DATABAS
    config_ConnectionString         = "Provider=SQLOLEDB;Data Source=creeper\SQLExpress2008;Initial Catalog=db_NForum;User Id=*****;Password=*****"
    config_ImageLocation            = "/image.asp"
    config_CropImageLocation        = "/cropimage.asp"
    config_UserImageLocation        = "/userimage.asp"
    config_GFXLocation              = "http://grafik.n-forum.se/"
    
  ' #### LOCKDOWN
    config_LockDown_All             = False
    config_LockDown_Forum           = False
    config_LockDown_Registrering    = False
    config_LockDown_Bilder          = False
    config_LockDown_Kommentarer     = False
    config_LockDown_Feedback        = False
  
  ' #### STANDARD 
    config_UpTemp                   = "/upload/temp/"
    config_Avatar                   = "/upload/avatar/"
    config_NotLoggedIn              = "/avdelning/medlem/loggain.asp"
    config_SystemUser               = 401
    config_MinSearch                = 3
    config_ImageFolder              = "C:\WebbRoot\N-Forum.se\bilder\" 
    config_UserImageFolder          = "C:\WebbRoot\N-Forum.se\anvbilder\" 
    
  ' #### SÄKERHET
    config_Hash_Salt_1              = "***"
    config_Hash_Salt_2              = "***"
  
  ' #### PAGING
    config_MaxAntalPosterPerSida    = 20
    config_MaxAntalSamlingPerSida   = 25
    config_MaxAntalTradarPerSida    = 25
    config_MaxAntalInlaggPerSida    = 10

  ' #### USERS
    config_UserTitle                = 8
    config_UserMaxImages            = 15
    config_UserImagesDays           = 7
    config_WelcomePMFrom            = 1
    
    config_WelcomePMTitle           = "Välkommen som medlem till N-Forum.se!"
    config_WelcomePM                = "[b]Hej och välkommen till N-Forum.se![/b]" & vbCrlf & vbCrlf & "Här kommer lite vägledande information för dig som är ny på [b]N-Forum.se[/b]." & vbCrlf & vbCrlf & "Som medlem har du nu möjlighet att lista dina spel och diskutera med likasinnade i forumet." & vbCrlf & "- [url=/avdelning/forum/]Forumet[/url]" & vbCrlf & "- [url=/avdelning/spel/]Spel[/url] / [url=/avdelning/konsol/]Konsoler[/url] / [url=/avdelning/tillbehor/]Tillbehör[/url]" & vbCrlf & vbCrlf & "Du kan nu ändra dina inställningar och komplettera din profil med information om dig själv." & vbCrlf & "- [url=/avdelning/medlem/]Din profil[/url]" & vbCrlf & "- [url=/avdelning/medlem/installningar.asp]Inställningar[/url]" & vbCrlf & vbCrlf & "Om du har något spelrelaterat du vill sälja eller köpa kan du använda dig av våran annonsavdelning." & vbCrlf & "- [url=/avdelning/annonser/]Annonser[/url]" & vbCrlf & vbCrlf & "Det finns en del regler som gäller på sidan, dessa kan du läsa om här." & vbCrlf & "- [url=/avdelning/medlem/information.asp]Regler[/url]" & vbCrlf & vbCrlf & "Vi hoppas att du kommer trivas!" & vbCrlf & vbCrlf & "[b]Mvh[/b]" & vbCrlf & "[i]N-Forum.se Red[/i]"
    
    config_AdDays                   = 14
    
  ' #### FORUM
    config_UseTrash                 = True
    config_Trashbin                 = 32
    config_RemoOlasta               = 30
    
  ' #### DATA
    config_StandardSize             = 11
    config_StandardFont             = 1
    
  ' #### Sidorna
    page_SubDomain                  = "www."
  
    page_NForum                     = page_SubDomain & "n-forum.se"
    page_GWDB                       = page_SubDomain & "gwdb.se"
    page_GBDB                       = page_SubDomain & "gbdb.se"
    page_GBCDB                      = page_SubDomain & "gbcdb.se"
    page_GBADB                      = page_SubDomain & "gbadb.se"
    page_DSDB                       = page_SubDomain & "dsdb.se"
    page_VBDB                       = page_SubDomain & "vbdb.se"
    page_SNESDB                     = page_SubDomain & "snesdb.se"
    page_N64DB                      = page_SubDomain & "n64db.se"
    page_GCDB                       = page_SubDomain & "gcdb.se"
    page_WIIDB                      = page_SubDomain & "wiidb.se"
    
  ' #### KÄNN AV VILKEN SIDA MAN ÄR PÅ
    test_Page = LCase(Trim(Request.ServerVariables("SERVER_NAME")))
    
  ' #### EPOSTKOMPONENTER
  '    ASPEMAIL        » Persits Software    http://www.persits.com
  '    ASPSMARTMAIL    » ASP Smart           http://www.aspsmart.com
  '    CDONTS          » Microsoft           http://www.microsoft.com
  '    CDOSYS          » Microsoft           http://www.microsoft.com (Win 2003 och uppåt)
  '    JMAIL           » Dimac Development   http://tech.dimac.net
  '    ASPMAIL         » ServerObjects       http://www.serverobjects.com
  
    KOMP_MAIL         = "JMAIL"
    MAIL_NOREPLY      = "noreply@n-forum.se" 
    MAIL_SMTP         = "127.0.0.1"
    MAIL_NAME         = "N-Forum.se"
    
    config_MinSearch = config_MinSearch - 1
  
    test_Page = "n-forum"

    page_Console  = "1,2,3,4,5,6,7,8,9,10,11,12,13"
    page_Category = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18"
    page_Name     = "Nintendo Forum - [N-Forum.se]"
    page_Address  = page_NForum
    
    page_Slogan = "Allt handlar bara om Nintendo"
    
    If config_LockDown_All Then
      Response.Write "<h2>N-Forum.se tillfälligt nerstängd!</h1>"
      Response.Write "<p>N-Forum.se är nerstängd av systemadministratören, om detta kommer bestå under en längre tid kommer vi med mer information snart.</p>" 
      Response.Write "<p>Kontaka oss på <a href='mailto:info@n-forum.se'>info@n-forum.se</a> vid frågor.</p>"
      Response.Write "<p>//SysAdmin</p>"
      Response.End
    End If
%>