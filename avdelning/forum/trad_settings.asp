<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
lID  = GetQ("e", "123", 0)
lDO  = GetQ("do", "ABC", 6)

Select Case LCase(lDO)
  Case "fuse"   : pageIs = "FUSE"
  Case "owner"  : pageIs = "OWNER"
  Case "break"  : pageIs = "BREAK"
  Case "move"   : pageIs = "MOVE"
  Case Else
    hidForm = True
End Select

If CONST_LOGIN Then

  RS_Open 1, "SELECT * " & _
             "FROM fsBB_Tradar " & _
             "WHERE tID = " & CLng(lID) & " AND tStatus_Raderad = 0", False
  
    If Not rsDB(1).EOF Then
      GetRights lID ' H�mta fram r�ttigheterna
      
      If Not sec_Trad_Admin Then Response.Redirect("trad_settings.asp")
    End If
    
  RS_Close 1

Else
  hidForm = True
End If

%>

<% If Not hidForm Then %>
<html>
  <head>
    <title> </title>
    <meta http-equiv="content-type" content="text/html; CHARSET=ISO-8859-1">
    <meta http-equiv="content-language" content="sv">
    
    <style type="text/css">
      body {
        padding:          10px;
        margin:           0;
        font:             12px Arial, Verdana, Sans-Serif;
      }
     
      #doform {
        background-color: #FFF;
        display:          none;
      }
      
      #doform input {
        font:             12px Verdana;
        width:            60px;
      }
    </style>
    
  </head>
  <body>
    
    <div id="doform">
      
     <form id="FrameBox_Form" method="POST">
       <% Select Case pageIs %>
       <% Case "FUSE" %>
         <p>Tr�den kommer att sl�s ihop med tr�den du anger nedan, tr�den du sl�r ihop denna med kommer bli huvudtr�d.</p>
         <p><strong>Tr�dens ID-nummer:</strong> <input type="text" name="newid" value=0 maxlength=100> <input style="width: 25px;" type="button" value="?" title="Kolla upp om tr�den finns." disabled></p>
         <p><strong>Observera att �ndringar nyligen gjorda i denna tr�d inte kommer sparas!</strong></p>
         <input type="hidden" name="do" value="fuse">
       <% Case "OWNER" %>
         <p>Du kommer bli �gare av denna tr�d.</p>
         <input type="hidden" name="do" value="owner">
       <% Case "BREAK" %>
         <p>Inl�gget kommer att brytas ut ur tr�den f�r att bli en egen tr�d i forumet.</p>
         <p><strong>Observera att �ndringar nyligen gjorda i detta inl�gg inte kommer sparas!</strong></p>
         <input type="hidden" name="do" value="break">
       <% Case "MOVE" %>
         <p>Inl�gget kommer att flyttas till tr�den du anger nedan.</p>
         <p><strong>Tr�dens ID-nummer:</strong> <input type="text" name="newid" value=0 maxlength=100> <input style="width: 25px;" type="button" value="?" title="Kolla upp om tr�den finns." disabled></p>
         <input type="hidden" name="do" value="move">
       <% End Select %>
       <input type="hidden" name="e" value="<% = lID %>">
     </form>
     
    </div>
  
    <script type="text/javascript">
      if(parent.location != this.location) {
        document.getElementById("doform").style.display = "block";
      } else {
        var str = this.location.toString();

        if(str.substr(str.length-18) != "/trad_settings.asp") {
          this.location = "trad_settings.asp";
        }
      }
    </script>
    
  </body>
</html>

<% End If %>

<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->