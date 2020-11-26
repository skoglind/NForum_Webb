<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

<%

  If Not CONST_LOGIN Then Response.Redirect("hold.asp")
  
  lID       = GetF("e","123",0)   ' Titelns ID
  sTextM    = GetF("xTextM","ABC",500)
  
  If Len(Trim(sTextM)) > 1 Then
      
    RS_Open 1, "SELECT * " & _
           "FROM fsBB_Tradar " & _
           "LEFT JOIN fsBB_Forum ON fID = tForum " & _
           "WHERE tID = " & CLng(lID), False
  
      If rsDB(1).EOF Then Response.Redirect("default.asp")
      
      If rsDB(1)("tStatus_Trad") Then
        text_TradID = rsDB(1)("tID")
      Else
        text_TradID = rsDB(1)("tStatus_UnderTrad")
      End If
      
      GetRights text_TradID ' Hämta fram rättigheterna
      If Not sec_Trad_Visa Then
        txt_Status  = "Posten anmäldes inte!"
        txt_Done    = 0
      Else
        text_ID     = CLng(rsDB(1)("tID"))
      
        RS_Close 1
      
        RS_Open 1, "SELECT * FROM fsBB_Anmal WHERE 1 = 2", True
          
          rsDB(1).AddNew
          
            rsDB(1)("anTradID")     = CLng(text_ID)
            rsDB(1)("anAnv")        = CONST_USERID
            rsDB(1)("anDatum")      = Now
            rsDB(1)("anTextM")      = sTextM
          
          rsDB(1).Update
        
        RS_Close 1
        
        txt_Status  = "Posten är nu anmäld!"
        txt_Done    = 1
      End If
  Else
    txt_Status  = "Posten anmäldes inte!"
    txt_Done    = 0
  End If
  
%>

<script type="text/javascript">
  parent.DoneReportPost("<% = txt_Status %>",<% = txt_Done %>);
  location.href = "/__AJAX/popbox/_action/hold.asp";
</script>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->