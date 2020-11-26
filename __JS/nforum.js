function getPageOffsetLeft (el) {
  var ol=el.offsetLeft;
  while ((el=el.offsetParent) != null) { ol += el.offsetLeft; }
  return ol;
}

function getPageOffsetTop (el) {
  var ot=el.offsetTop;
  while ((el=el.offsetParent) != null) { ot += el.offsetTop; }
  return ot;
}

function LoadImg(obj,img) {
  var e = document.getElementById(obj);
  var preload = new Image();
  preload.src = img;
  
  e.src = preload.src;
}

function SetText(obj,data) {
  var e = document.getElementById(obj);
  e.innerHTML = data;
}

function SetLink(obj,data) {
  var e = document.getElementById(obj);
  e.href = data;
}

function KeepOnline() {
  RetriveData();
  setTimeout("KeepOnline();",60000);
}

function OpenPopBox(url) {
  var e = document.getElementById("popBox");
  var inb = document.getElementById("popBox_Inner");
  var urlLoc;
  urlLoc = "/__AJAX/popbox/" + url
  
  inb.innerHTML = "<p><img style='float: left; margin: 10px;' src='http://grafik.n-forum.se/loader_big.gif'><strong style='float: left; margin-top: 12px;'>Var god vänta, laddar innehållet...</strong><br><br><br><br><br><br></p>";
  
  e.style.display = "block";
  
  loadPopBox(urlLoc);
}

function ClosePopBox() {
  var e = document.getElementById("popBox");
  var inb = document.getElementById("popBox_Inner");
  var bt = document.getElementById("popBox_BT");
  var frame = document.getElementById("popBox_Frame");
  
  //if(confirm("Vill du avbryta?")) {
    e.style.display = "none";
    frame.src = "/__AJAX/popbox/_action/hold.asp";
    inb.innerHTML = " ";
    bt.disabled = true;
  //}
}

function OK_PopBox() {
  var e = document.getElementById("popForm");
  e.submit();
}

function CountItemList(o) {
  var oDiv = document.getElementById(o);
  var itemCnt = 0;
  if(oDiv.firstChild) {
    var oChild = oDiv.firstChild;
    while(oChild) {
      if(oChild.nodeType==1) {itemCnt=itemCnt+1;}
      oChild = oChild.nextSibling;
    }
  }
  
  return itemCnt;
}

function SavedCollection(id) {
  var e = document.getElementById("listicon_" + id);
  if(e != null) {
    e.style.display = "block";
  }
  
  window.alert("Titeln är nu tillagd i din samling!");
  ClosePopBox();
}

function OpenCollection(o, id, pid, doit) {
  OpenPopBox("collection.asp?tp=" + o + "&e=" + id + "&id=" + pid + "&do=" + doit);
}

function OpenReportPost(id) {
  OpenPopBox("report.asp?e=" + id);
}

function DoneReportPost(sStatus, iDone) {
  window.alert(sStatus);
  if(iDone == 1) {ClosePopBox();}
}

function AddToList(id,o,li) {
  var bdata = Right(o.src,7);
  bdata = bdata.toLowerCase();
  
  if(bdata != "_no.png") {
    if(confirm("Vill du lägga till spelet i din samling?")) {
      SendData("/__AJAX/qlist.asp","e=" + id,o,li);
    }
  }
}

function AddToListP2(o,li) {
  o.src       = "http://grafik.n-forum.se/buttons/lista_no.png";
  o.className = "listbutton_no";
  li.style.display = 'block';
  //li.src      = "http://grafik.n-forum.se/icons/listed.gif";
  window.alert("Spelet är nu tillagd i din samling!");
}

function UpdateStats(value) {
  var pmE = document.getElementById("anPM");
  var onE = document.getElementById("anOn");
  
  var values;
  values = value.split(":", 3);

  pmE.innerHTML = "PM (" + values[2] + ")";
  onE.innerHTML = "Online (" + values[1] + ")";
}

function Left(str, n){
  if (n <= 0)
      return "";
  else if (n > String(str).length)
      return str;
  else
      return String(str).substring(0,n);
}
function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}

function addText(o,s) {
  var obj = document.getElementById(o);
  obj.focus();
  
  if (document.selection) {
    var sel = document.selection.createRange();
    if(sel.text.length > 0) {
      sel.text = "[" + s + "]" + sel.text + "[/" + s + "]";
    } else {
      stxt = "[" + s + "][/" + s + "]";
      obj.value = obj.value + stxt;
    }
  } else {
    lStart = obj.selectionStart;
    lEnd = obj.selectionEnd;
    if(lEnd > lStart) {
      tFront = obj.value.substr(0, lStart);
      sSel = obj.value.substr(lStart, (lEnd - lStart));
      tBack = obj.value.substr(lEnd, (obj.value.length - lEnd));
      obj.value = tFront + "[" + s + "]" + sSel + "[/" + s + "]" + tBack;
    } else {
      stxt = "[" + s + "][/" + s + "]";
      obj.value = obj.value + stxt;
    }
  }
}

function addTextEnd(o,s) {
  var obj = document.getElementById(o);
  obj.focus();

  obj.value = obj.value + s;
}

function doActionWithPrompt(sLink, sQ) {
  if (confirm(sQ)) {
    location.href = sLink;
  }
}

function showFrameBox(sLink, sTitle) {
  hide("jsFrameBox");
  document.getElementById("FrameBox_Frame").src = sLink;
  document.getElementById("FrameBox_Title").innerHTML = sTitle;
  show("jsFrameBox");
}

function submitFrameBox() {
  var fBox  = document.getElementById('FrameBox_Frame');
  
  if(fBox.contentWindow.document) {
    var fForm = fBox.contentWindow.document.getElementById("FrameBox_Form");
  } else {
    var fForm = fBox.contentDocument.getElementById("FrameBox_Form");
  }
  
  if(confirm("Vill du utföra årgärden?")) {
    fForm.action = "_action/settings_save.asp";
    fForm.submit();
  }
}

function do_search(url, q, l) {  
  //if(q.length > l) {
    location.href = "/avdelning/" + url + q;
  //} else {
  //  langd = l + 1;
  //  alert("Din sökning måste bestå av FLER än " + langd.toString() + " tecken!");
  //}
  
  return false;
}

var idCount = 1;

function addData(o,dt,cl,w) {
  idCount++;
  var e = document.getElementById(o);
  var c = document.getElementById(dt);
  
  var div = document.createElement('div');
  div.setAttribute('id', "TitelBox_" + idCount.toString());
  div.style.width = w + "px";
  div.className   = cl;

  sHTML = c.innerHTML;
  sHTML = sHTML.replace(/_0000/g,"_" + idCount.toString());
  
  div.innerHTML = sHTML;
  
  e.appendChild(div);
}

function removeData(o,dt) {
  var e = document.getElementById(o);
  var c = document.getElementById(dt);
  
  c.removeChild(e);
}

function fillData(o,data) {
  var e = document.getElementById(o + "_" + idCount.toString());
  
  if(data != "") {
    e.style.color = "#000";
    e.value = data;
  }
}

function showGrade(v) {
  var i = 1;
  var uv = document.getElementById("userbg").value;

  if(v == 0) {
    if(uv == 0) {
      while (i <= 6) {
        document.getElementById("bg" + i).src = "http://grafik.n-forum.se/icons/grade_no.gif";
        i++;
      }
    } else {
      showGrade(uv);
    }
  } else {
    while (i <= 6) {
      if(i <= v) {
        document.getElementById("bg" + i).src = "http://grafik.n-forum.se/icons/grade_on.gif";
      } else {
        document.getElementById("bg" + i).src = "http://grafik.n-forum.se/icons/grade_off.gif";
      }
      i++;
    }
  }
}

function setGrade(id,v) {
  var e = document.getElementById("userbg");
  e.value = v;
  showGrade(0);
  
  saveGrade(id,v);
}

function searchRegCode(id) {
  var e = document.getElementById(id);
  var es = document.getElementById("soktraff");
  
  if(e.value.length > 2) {
    searchCode(e.value);
  }else{
    es.innerHTML = "";
  }
}

function show(id) {
  var e = document.getElementById(id);
  e.style.display = "block";
}

function hide(id) {
  var e = document.getElementById(id);
  e.style.display = "none";
}

function showSwap(id) {
  var e = document.getElementById(id);
  if (e.style.display == "block") {
    hide(id);
  } else {
    show(id);
  }
}

function show_presetImageBox(eBoxName, sImg, smallImg) {
  var eBox      = document.getElementById(eBoxName);
  var smallImg  = document.getElementById(smallImg);
  var eImg      = document.getElementById(eBoxName + "_image");
  
  myTop    = getPageOffsetTop(smallImg);
  myLeft   = getPageOffsetLeft(smallImg);
  myTop    = myTop - 200;
  myLeft   = myLeft + 78;
  
  eBox.style.top  = myTop + "px";
  eBox.style.left = myLeft + "px";
  
  var myImage = new Image();
  myImage.src = sImg;
  
  eImg.src = myImage.src;
  
  eBox.style.display = "block";
}

function hide_presetImageBox(eBoxName) {
  var eBox = document.getElementById(eBoxName);
  var eImg = document.getElementById(eBoxName + "_image");
  
  eBox.style.display = "none";
  
  var myImage = new Image();
  myImage.src = "http://grafik.n-forum.se/img/loading.png";
  
  eImg.src = myImage.src;
}

var xmlhttp=false;
var gErr=true;
/*@cc_on @*/
/*@if (@_jscript_version >= 5)
// JScript gives us Conditional compilation, we can cope with old IE versions.
// and security blocked creation of the objects.
  try {
  xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
  } catch (e) {
   try {
    xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
   } catch (E) {
    xmlhttp = false;
   }
  }
@end @*/
if(!xmlhttp && typeof XMLHttpRequest != 'undefined'){
  xmlhttp = new XMLHttpRequest();
}

var xmlhttp2=false;
var gErr2=true;
/*@cc_on @*/
/*@if (@_jscript_version >= 5)
// JScript gives us Conditional compilation, we can cope with old IE versions.
// and security blocked creation of the objects.
  try {
  xmlhttp2 = new ActiveXObject("Msxml2.XMLHTTP");
  } catch (e) {
   try {
    xmlhttp2 = new ActiveXObject("Microsoft.XMLHTTP");
   } catch (E) {
    xmlhttp2 = false;
   }
  }
@end @*/
if(!xmlhttp2 && typeof XMLHttpRequest != 'undefined'){
  xmlhttp2 = new XMLHttpRequest();
}

var xmlhttp3=false;
var gErr3=true;
/*@cc_on @*/
/*@if (@_jscript_version >= 5)
// JScript gives us Conditional compilation, we can cope with old IE versions.
// and security blocked creation of the objects.
  try {
  xmlhttp3 = new ActiveXObject("Msxml2.XMLHTTP");
  } catch (e) {
   try {
    xmlhttp3 = new ActiveXObject("Microsoft.XMLHTTP");
   } catch (E) {
    xmlhttp3 = false;
   }
  }
@end @*/
if(!xmlhttp3 && typeof XMLHttpRequest != 'undefined'){
  xmlhttp3 = new XMLHttpRequest();
}

function RetriveData() {
  var xmlret;

  xmlhttp.open("GET", "/__AJAX/return.asp");
  xmlhttp.onreadystatechange = function() {
    if (xmlhttp.readyState == 4) {
      switch (xmlhttp.status) {
        case 200:
          xmlret = xmlhttp.responseText;
          break;
        case 404:
          xmlret = "0:0";
          break;
        case 500:
          xmlret = "0:0";
          break;
        default:
          xmlret = "0:0";
      }

      UpdateStats(xmlret);
    }
  }
  xmlhttp.send(null);
}

function SendData(url,params,o,li) {
  var xmlret2;

  xmlhttp2.open("POST", url);
  
  xmlhttp2.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  xmlhttp2.setRequestHeader("Content-length", params.length);
  xmlhttp2.setRequestHeader("Connection", "close");
  
  xmlhttp2.onreadystatechange = function() {
    if (xmlhttp2.readyState == 4) {
      switch (xmlhttp2.status) {
        case 200:
          AddToListP2(o,li);
          break;
        case 404:
          window.alert("Kommer för tillfället inte åt servern, spelet är därför inte listat. Prova igen!");
          break;
        case 500:
          window.alert("Ett scriptfel har uppståt. Prova igen!\nOm felet återkommer kontakta oss på N-Forum.se.");
          break;
        default:
          window.alert("Ett okänt fel uppstod. Prova igen!");
      }
    }
  }
  xmlhttp2.send(params);
}

function saveGrade(id,v) {
  var xmlret;

  xmlhttp.open("GET", "/__AJAX/betyg.asp?e=" + id + "&b=" + v);
  xmlhttp.onreadystatechange = function() {
    if (xmlhttp.readyState == 4) {
      switch (xmlhttp.status) {
        case 200:
          //window.alert(xmlhttp.responseText);
          break;
        case 404:
          window.alert("Kommer för tillfället inte åt servern, betyget är därför inte sparat. Prova igen!");
          break;
        case 500:
          window.alert("Ett scriptfel har uppståt. Prova igen!\nOm felet återkommer kontakta oss på N-Forum.se.");
          break;
        default:
          window.alert("Ett okänt fel uppstod. Prova igen!");
      }
    }
  }
  xmlhttp.send(null);
}

function searchCode(c) {
  var xmlret;
  var e = document.getElementById("soktraff");

  xmlhttp.open("GET", "/__AJAX/findreg.asp?kod=" + c);
  xmlhttp.onreadystatechange = function() {
    if (xmlhttp.readyState == 4) {
      switch (xmlhttp.status) {
        case 200:
          xmlret = xmlhttp.responseText;
          e.innerHTML = xmlret;
          break;
        case 404:
          e.innerHTML = "<p class='nf_center'><em>Kommer för tillfället inte åt servern, sökningen kunde inte utföras. Prova igen!</em></p>";
          break;
        case 500:
          e.innerHTML = "<p class='nf_center'><em>Ett scriptfel har uppståt. Prova igen!\nOm felet återkommer kontakta oss på N-Forum.se.</em></p>";
          break;
        default:
          e.innerHTML = "<p class='nf_center'><em>Ett okänt fel uppstod. Prova igen!</em></p>";
      }
    }
  }
  xmlhttp.send(null);
}

function clearField(o,v) {
  // e = document.getElementById(o);
  if(o.value == v) {
    o.style.color = "#000";
    o.value = "";
  }
}

function retypeField(o,v) {
  //var e = document.getElementById(o);
  if(o.value == "") {
    o.style.color = "#AAA";
    o.value = v;
  }
}

function setChecked(e) {
  var be = document.getElementById(e + "_all");
  if(allBoxesChecked(e)) {
    be.checked = true;
  } else {
    be.checked = false;
  }
}

function setBoxes(val,e) {
  var bAll, bID;
  bAll = true;
  bID = 0;
  
  while(bAll) {
    bID = bID + 1;
    if(document.getElementById(e + "_" + bID.toString())) {
      document.getElementById(e + "_" + bID.toString()).checked = val;
    } else {
      bAll = false;
    }
  }
}

function allBoxesChecked(e) {
  var bAll, bID, bRet;
  bAll = true;
  bRet = true;
  bID = 0;
  
  while(bAll) {
    bID = bID + 1;
    if(document.getElementById(e + "_" + bID.toString())) {
      if(document.getElementById(e + "_" + bID.toString()).checked == false) {bRet = false;}
    } else {
      bAll = false;
    }
  }
  
  return bRet;
}

function loadPopBox(url) {
  var inb = document.getElementById("popBox_Inner");
  var bt = document.getElementById("popBox_BT");
  
  xmlhttp3.open("GET", url);
    
    xmlhttp3.onreadystatechange = function() {
      if (xmlhttp3.readyState == 4) {
        switch (xmlhttp3.status) {
          case 200:
            inb.innerHTML = xmlhttp3.responseText;
            bt.disabled = false;
            break;
          case 404:case 500:default:
            inb.innerHTML = "<p><img style='float: left; margin: 10px;' src='http://grafik.n-forum.se/icons/del.png'><strong style='float: left; margin-top: 14px; color: #A00;'>Ett fel uppstod, kunde inte ladda innehållet!</strong><br><br><br><br><br><br></p>";
        }
      }
    }
    xmlhttp3.send(null);
}

function DeleteCollection(o, id) {
  if(confirm("Vill du ta bort titeln?")) {
   
    var xmlret;

    xmlhttp3.open("GET", "/__AJAX/popbox/_action/deletecollection.asp?e=" + id + "&tp=" + o);
    xmlhttp3.onreadystatechange = function() {
      if (xmlhttp3.readyState == 4) {
        switch (xmlhttp3.status) {
          case 200:
            xmlret = xmlhttp3.responseText;
            if(xmlret == "1") {
              rh_removeRow("titleListed_List","titleListed_Row_", id);
              
              if(CountItemList("titleListed_List") == 0) {
                document.getElementById("titleListed_List").style.display = "none";
                document.getElementById("titleListed_Mess").style.display = "block";
              }
            } else {
              window.alert("Raderingen misslyckades!");
            }
            break;
          case 404:case 500:default:
            window.alert("Raderingen misslyckades!");
        }
      }
    }
    xmlhttp3.send(null);

  }
}

function doPreview(o,a,s) {
  var eP = document.getElementById("edit_preview");
  var wB = document.getElementById("warn_preview");
  var pB = document.getElementById("post_preview");
  var tP = document.getElementById("post_preview_text");
  
  var e = document.getElementById(o);
  var c;
  c = e.value;
  c = c.replace(/\n/g, "[newline]");
  c = c.replace(/&/g, "[ampersand]");
  c = c.replace(/#/g, "[bracket]");
  
  if(e.value == "") {
    window.alert("Det finns inget att förhandsgranska!");
    tP.innerHTML = "";
    eP.style.display = "block";
    wB.style.display = "none";
    pB.style.display = "none";
  }else{
    xmlhttp3.open("GET", "/__AJAX/bbcode.asp?txt=" + c + "&a=" + a + "&s=" + s);
    
    xmlhttp3.onreadystatechange = function() {
      if (xmlhttp3.readyState == 4) {
        switch (xmlhttp3.status) {
          case 200:
            tP.innerHTML = xmlhttp3.responseText;
            eP.style.display = "none";
            wB.style.display = "block";
            pB.style.display = "block";
            break;
          case 404:case 500:default:
            tP.innerHTML = "* Ett fel någonstans, kan inte uppdatera förhandsgranskningen.<br><br>" + tP.innerHTML;
            eP.style.display = "block";
            wB.style.display = "none";
            pB.style.display = "none";
        }
        
        location.href = "#PREVIEW";
      }
    }
    xmlhttp3.send(null);
  } 
}

function closePreview(o,bGo) {
  var eP = document.getElementById("edit_preview");
  var wB = document.getElementById("warn_preview");
  var pB = document.getElementById("post_preview");
  var tP = document.getElementById("post_preview_text");
  
  var e = document.getElementById(o);
  
  tP.innerHTML = "";
  eP.style.display = "block";
  wB.style.display = "none";
  pB.style.display = "none";
  
  if(bGo) {
    location.href = "#EDIT";
    e.focus();
  }
}

function rh_cloneRow(clone, box, cloneName, cloneID, cloneType, cloneData) {
  var eClone  = document.getElementById(clone);
  var eBox    = document.getElementById(box);
  var oNew    = document.createElement(cloneType); /* LI eller DIV */
  var sHTML;
  
  sHTML = eClone.innerHTML;
  
  /* Fixa texterna i blocket som ska klonas */
  
    var chData = cloneData.split(";;");
    for(i = 0; i < chData.length; i++){
      if(chData[i].split("==")[0] != "") {
        sHTML = sHTML.replace(new RegExp("XXXX_" + chData[i].split("==")[0], "g"), chData[i].split("==")[1]);
      }
    }
  
  /* ###################################### */
  
  oNew.innerHTML = sHTML;
  
  oNew.setAttribute("id", cloneName + cloneID);
  eBox.appendChild(oNew);
}

function rh_updateRow(clone, orgrow, cloneName, cloneID, cloneType, cloneData) {
  var eClone  = document.getElementById(clone);
  var oOld    = document.getElementById(orgrow);
  var sHTML;
  
  sHTML = eClone.innerHTML;
  
  /* Fixa texterna i blocket som ska klonas */
  
    var chData = cloneData.split(";;");
    for(i = 0; i < chData.length; i++){
      if(chData[i].split("==")[0] != "") {
        sHTML = sHTML.replace(new RegExp("XXXX_" + chData[i].split("==")[0], "g"), chData[i].split("==")[1]);
      }
    }
  
  /* ###################################### */
  
  oOld.innerHTML = sHTML;
}

function rh_removeRow(box, cloneName, cloneID) {
  var eBox = document.getElementById(box);
  var oOld = document.getElementById(cloneName + cloneID);
  eBox.removeChild(oOld);
}

function toggleBox(o,bt) {
  var e = document.getElementById(o);
  var pm = document.getElementById(bt);
  
  if(e.style.display == "block") {
    e.style.display = "none";
    pm.src = "http://grafik.n-forum.se/icons/plus.gif";
  } else {
    e.style.display = "block";
    pm.src = "http://grafik.n-forum.se/icons/minus.gif";
  }
}