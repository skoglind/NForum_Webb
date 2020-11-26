<%
  Function SetConstants()
    CONST_LOGIN           = Session.Value("NFORUM_Login")
    CONST_USERID          = Session.Value("NFORUM_ID")
    CONST_TITEL           = Session.Value("NFORUM_TitelID")
    CONST_ADMIN           = Session.Value("NFORUM_Admin")
    CONST_CMS             = Session.Value("NFORUM_CMS")
    CONST_CMS_RIGHTS      = Session.Value("NFORUM_CMS_RIGHTS")
    CONST_USERNAME        = Session.Value("NFORUM_AnvNamn")
    CONST_PUBLISH         = Session.Value("NFORUM_Publish")
    CONST_DAYSMEMBER      = Session.Value("NFORUM_DaysMember")
    
    CONST_PM              = Session.Value("NFORUM_PM")
    
    CONST_SET_PMSIDA      = Session.Value("SET_PmSida")
    CONST_SET_TRADARSIDA  = Session.Value("SET_TradarSida") 
    CONST_SET_INLAGGSIDA  = Session.Value("SET_InlaggSida")
    
    CONST_SET_SIZE        = Session.Value("SET_FontSize")
    CONST_SET_FONT        = Session.Value("SET_FontFam")
    
    CONST_SET_QUICK       = Session.Value("SET_Quick")
    
    CONST_SET_AVATAR      = Session.Value("SET_ShowAvatar")
    CONST_SET_SIGN        = Session.Value("SET_ShowSign")
    
    SEC_TITEL             = Session.Value("NFORUM_TitelID")
    
    If CONST_SET_PMSIDA < 10      Then CONST_SET_PMSIDA = config_MaxAntalPosterPerSida
    If CONST_SET_TRADARSIDA < 10  Then CONST_SET_TRADARSIDA = config_MaxAntalTradarPerSida
    If CONST_SET_INLAGGSIDA < 10  Then CONST_SET_INLAGGSIDA = config_MaxAntalInlaggPerSida
    
    If CONST_SET_SIZE < 8         Then CONST_SET_SIZE = config_StandardSize
    If CONST_SET_FONT < 1         Then CONST_SET_FONT = config_StandardFont
  End Function

  CONST_LOGIN           = Session.Value("NFORUM_Login")
  CONST_USERID          = Session.Value("NFORUM_ID")
  CONST_TITEL           = Session.Value("NFORUM_TitelID")
  CONST_ADMIN           = Session.Value("NFORUM_Admin")
  CONST_CMS             = Session.Value("NFORUM_CMS")
  CONST_CMS_RIGHTS      = Session.Value("NFORUM_CMS_RIGHTS")
  CONST_USERNAME        = Session.Value("NFORUM_AnvNamn")
  CONST_PUBLISH         = Session.Value("NFORUM_Publish")
  CONST_DAYSMEMBER      = Session.Value("NFORUM_DaysMember")
  
  CONST_PM              = Session.Value("NFORUM_PM")
  
  CONST_SET_PMSIDA      = Session.Value("SET_PmSida")
  CONST_SET_TRADARSIDA  = Session.Value("SET_TradarSida") 
  CONST_SET_INLAGGSIDA  = Session.Value("SET_InlaggSida")
  
  CONST_SET_SIZE        = Session.Value("SET_FontSize")
  CONST_SET_FONT        = Session.Value("SET_FontFam")
  
  If CONST_LOGIN Then
    CONST_SET_AVATAR      = Session.Value("SET_ShowAvatar")
    CONST_SET_SIGN        = Session.Value("SET_ShowSign")
    CONST_SET_QUICK       = Session.Value("SET_Quick")
  Else
    CONST_SET_AVATAR      = True
    CONST_SET_SIGN        = True
    CONST_SET_QUICK       = 0
    CONST_DAYSMEMBER      = 0
  End if
  
  SEC_TITEL             = Session.Value("NFORUM_TitelID")
  
  If CONST_SET_PMSIDA < 10      Then CONST_SET_PMSIDA = config_MaxAntalPosterPerSida
  If CONST_SET_TRADARSIDA < 10  Then CONST_SET_TRADARSIDA = config_MaxAntalTradarPerSida
  If CONST_SET_INLAGGSIDA < 10  Then CONST_SET_INLAGGSIDA = config_MaxAntalInlaggPerSida
  
  If CONST_SET_SIZE < 8         Then CONST_SET_SIZE = config_StandardSize
  If CONST_SET_FONT < 1         Then CONST_SET_FONT = config_StandardFont
%>