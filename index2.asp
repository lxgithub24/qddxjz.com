<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connbbs.asp" -->
<%
var rsjian = Server.CreateObject("ADODB.Recordset");
rsjian.ActiveConnection = MM_connbbs_STRING;
rsjian.Source = "SELECT * FROM fatie";
rsjian.CursorType = 0;
rsjian.CursorLocation = 2;
rsjian.LockType = 1;
rsjian.Open();
var rsjian_numRows = 0;
%>
<%
var rsfatie = Server.CreateObject("ADODB.Recordset");
rsfatie.ActiveConnection = MM_connbbs_STRING;
rsfatie.Source = "SELECT ArticleId, Topic, Content, PostTime, Username, IP FROM fatie ORDER BY PostTime DESC";
rsfatie.CursorType = 0;
rsfatie.CursorLocation = 2;
rsfatie.LockType = 1;
rsfatie.Open();
var rsfatie_numRows = 0;
%>
<%
var rsjiaqi = Server.CreateObject("ADODB.Recordset");
rsjiaqi.ActiveConnection = MM_connbbs_STRING;
rsjiaqi.Source = "SELECT * FROM shuqitie ORDER BY PostTime DESC";
rsjiaqi.CursorType = 0;
rsjiaqi.CursorLocation = 2;
rsjiaqi.LockType = 1;
rsjiaqi.Open();
var rsjiaqi_numRows = 0;
%>
<%
var reply = Server.CreateObject("ADODB.Recordset");
reply.ActiveConnection = MM_connbbs_STRING;
reply.Source = "SELECT * FROM Replies";
reply.CursorType = 0;
reply.CursorLocation = 2;
reply.LockType = 1;
reply.Open();
var reply_numRows = 0;
%>
<%
var zaqd = Server.CreateObject("ADODB.Recordset");
zaqd.ActiveConnection = MM_connbbs_STRING;
zaqd.Source = "SELECT * FROM zaoanqingda ORDER BY PostTime DESC";
zaqd.CursorType = 0;
zaqd.CursorLocation = 2;
zaqd.LockType = 1;
zaqd.Open();
var zaqd_numRows = 0;
%>
<%
var adpdy = Server.CreateObject("ADODB.Recordset");
adpdy.ActiveConnection = MM_connbbs_STRING;
adpdy.Source = "SELECT * FROM duanqipaidanyuan ORDER BY ArticleId DESC";
adpdy.CursorType = 0;
adpdy.CursorLocation = 2;
adpdy.LockType = 1;
adpdy.Open();
var adpdy_numRows = 0;
%>
<%
var ajpdy = Server.CreateObject("ADODB.Recordset");
ajpdy.ActiveConnection = MM_connbbs_STRING;
ajpdy.Source = "SELECT * FROM jiaqipaidanyuan ORDER BY ArticleId DESC";
ajpdy.CursorType = 0;
ajpdy.CursorLocation = 2;
ajpdy.LockType = 1;
ajpdy.Open();
var ajpdy_numRows = 0;
%>
<%
var ajcxy = Server.CreateObject("ADODB.Recordset");
ajcxy.ActiveConnection = MM_connbbs_STRING;
ajcxy.Source = "SELECT * FROM jiaqicuxiaoyuan ORDER BY ArticleId DESC";
ajcxy.CursorType = 0;
ajcxy.CursorLocation = 2;
ajcxy.LockType = 1;
ajcxy.Open();
var ajcxy_numRows = 0;
%>
<%
var adcxy = Server.CreateObject("ADODB.Recordset");
adcxy.ActiveConnection = MM_connbbs_STRING;
adcxy.Source = "SELECT * FROM duanqicuxiaoyuan ORDER BY ArticleId DESC";
adcxy.CursorType = 0;
adcxy.CursorLocation = 2;
adcxy.LockType = 1;
adcxy.Open();
var adcxy_numRows = 0;
%>
<%
var adfwy = Server.CreateObject("ADODB.Recordset");
adfwy.ActiveConnection = MM_connbbs_STRING;
adfwy.Source = "SELECT * FROM duanqifuwuyuan ORDER BY ArticleId DESC";
adfwy.CursorType = 0;
adfwy.CursorLocation = 2;
adfwy.LockType = 1;
adfwy.Open();
var adfwy_numRows = 0;
%>
<%
var ajfwy = Server.CreateObject("ADODB.Recordset");
ajfwy.ActiveConnection = MM_connbbs_STRING;
ajfwy.Source = "SELECT * FROM jiaqifuwuyuan ORDER BY ArticleId DESC";
ajfwy.CursorType = 0;
ajfwy.CursorLocation = 2;
ajfwy.LockType = 1;
ajfwy.Open();
var ajfwy_numRows = 0;
%>
<%
var adwjdc = Server.CreateObject("ADODB.Recordset");
adwjdc.ActiveConnection = MM_connbbs_STRING;
adwjdc.Source = "SELECT * FROM duanqiwenjuandiaocha ORDER BY ArticleId DESC";
adwjdc.CursorType = 0;
adwjdc.CursorLocation = 2;
adwjdc.LockType = 1;
adwjdc.Open();
var adwjdc_numRows = 0;
%>
<%
var asxs = Server.CreateObject("ADODB.Recordset");
asxs.ActiveConnection = MM_connbbs_STRING;
asxs.Source = "SELECT * FROM jiaqishixisheng ORDER BY ArticleId DESC";
asxs.CursorType = 0;
asxs.CursorLocation = 2;
asxs.LockType = 1;
asxs.Open();
var asxs_numRows = 0;
%>
<%
var ahq = Server.CreateObject("ADODB.Recordset");
ahq.ActiveConnection = MM_connbbs_STRING;
ahq.Source = "SELECT * FROM duanqihunqing ORDER BY ArticleId DESC";
ahq.CursorType = 0;
ahq.CursorLocation = 2;
ahq.LockType = 1;
ahq.Open();
var ahq_numRows = 0;
%>
<%
var aqt = Server.CreateObject("ADODB.Recordset");
aqt.ActiveConnection = MM_connbbs_STRING;
aqt.Source = "SELECT * FROM jiaqiqita ORDER BY ArticleId DESC";
aqt.CursorType = 0;
aqt.CursorLocation = 2;
aqt.LockType = 1;
aqt.Open();
var aqt_numRows = 0;
%>
<%
var adqt = Server.CreateObject("ADODB.Recordset");
adqt.ActiveConnection = MM_connbbs_STRING;
adqt.Source = "SELECT * FROM duanqiqita ORDER BY ArticleId DESC";
adqt.CursorType = 0;
adqt.CursorLocation = 2;
adqt.LockType = 1;
adqt.Open();
var adqt_numRows = 0;
%>
<%
var ajqt2 = Server.CreateObject("ADODB.Recordset");
ajqt2.ActiveConnection = MM_connbbs_STRING;
ajqt2.Source = "SELECT * FROM jiaqiqita2 ORDER BY ArticleId DESC";
ajqt2.CursorType = 0;
ajqt2.CursorLocation = 2;
ajqt2.LockType = 1;
ajqt2.Open();
var ajqt2_numRows = 0;
%>
<%
var xess = Server.CreateObject("ADODB.Recordset");
xess.ActiveConnection = MM_connbbs_STRING;
xess.Source = "SELECT * FROM xuexikeben ORDER BY ArticleId DESC";
xess.CursorType = 0;
xess.CursorLocation = 2;
xess.LockType = 1;
xess.Open();
var xess_numRows = 0;
%>
<%
var yp = Server.CreateObject("ADODB.Recordset");
yp.ActiveConnection = MM_connbbs_STRING;
yp.Source = "SELECT * FROM yuleyongpin ORDER BY ArticleId DESC";
yp.CursorType = 0;
yp.CursorLocation = 2;
yp.LockType = 1;
yp.Open();
var yp_numRows = 0;
%>
<%
var axky = Server.CreateObject("ADODB.Recordset");
axky.ActiveConnection = MM_connbbs_STRING;
axky.Source = "SELECT * FROM xuexikaoyanziliao ORDER BY ArticleId DESC";
axky.CursorType = 0;
axky.CursorLocation = 2;
axky.LockType = 1;
axky.Open();
var axky_numRows = 0;
%>
<%
var aydn = Server.CreateObject("ADODB.Recordset");
aydn.ActiveConnection = MM_connbbs_STRING;
aydn.Source = "SELECT * FROM computer ORDER BY ArticleId DESC";
aydn.CursorType = 0;
aydn.CursorLocation = 2;
aydn.LockType = 1;
aydn.Open();
var aydn_numRows = 0;
%>
<%
var axwy = Server.CreateObject("ADODB.Recordset");
axwy.ActiveConnection = MM_connbbs_STRING;
axwy.Source = "SELECT * FROM xuexiwaiyu ORDER BY ArticleId DESC";
axwy.CursorType = 0;
axwy.CursorLocation = 2;
axwy.LockType = 1;
axwy.Open();
var axwy_numRows = 0;
%>
<%
var aysj = Server.CreateObject("ADODB.Recordset");
aysj.ActiveConnection = MM_connbbs_STRING;
aysj.Source = "SELECT * FROM phone ORDER BY ArticleId DESC";
aysj.CursorType = 0;
aysj.CursorLocation = 2;
aysj.LockType = 1;
aysj.Open();
var aysj_numRows = 0;
%>
<%
var axqt = Server.CreateObject("ADODB.Recordset");
axqt.ActiveConnection = MM_connbbs_STRING;
axqt.Source = "SELECT * FROM xuexiqita ORDER BY ArticleId DESC";
axqt.CursorType = 0;
axqt.CursorLocation = 2;
axqt.LockType = 1;
axqt.Open();
var axqt_numRows = 0;
%>
<%
var ayqt = Server.CreateObject("ADODB.Recordset");
ayqt.ActiveConnection = MM_connbbs_STRING;
ayqt.Source = "SELECT * FROM yuleqita ORDER BY ArticleId DESC";
ayqt.CursorType = 0;
ayqt.CursorLocation = 2;
ayqt.LockType = 1;
ayqt.Open();
var ayqt_numRows = 0;
%>
<%
var Repeat1__numRows = 5;
var Repeat1__index = 0;
rsfatie_numRows += Repeat1__numRows;
%>
<%
var Repeat2__numRows = 5;
var Repeat2__index = 0;
rsjiaqi_numRows += Repeat2__numRows;
%>
<%
var Repeat16__numRows = 5;
var Repeat16__index = 0;
ajqt2_numRows += Repeat16__numRows;
%>
<%
var Repeat3__numRows = 5;
var Repeat3__index = 0;
xess_numRows += Repeat3__numRows;
%>
<%
var Repeat17__numRows = 5;
var Repeat17__index = 0;
yp_numRows += Repeat17__numRows;
%>
<%
var Repeat18__numRows = 5;
var Repeat18__index = 0;
axky_numRows += Repeat18__numRows;
%>
<%
var Repeat19__numRows = 5;
var Repeat19__index = 0;
aydn_numRows += Repeat19__numRows;
%>
<%
var Repeat20__numRows = 5;
var Repeat20__index = 0;
axwy_numRows += Repeat20__numRows;
%>
<%
var Repeat21__numRows = 5;
var Repeat21__index = 0;
aysj_numRows += Repeat21__numRows;
%>
<%
var Repeat22__numRows = 5;
var Repeat22__index = 0;
axqt_numRows += Repeat20__numRows;
%>
<%
var Repeat23__numRows = 5;
var Repeat23__index = 0;
ayqt_numRows += Repeat23__numRows;
%>
<%
var Repeat15__numRows = 5;
var Repeat15__index = 0;
adqt_numRows += Repeat15__numRows;
%>
<%
var Repeat14__numRows = 5;
var Repeat14__index = 0;
aqt_numRows += Repeat14__numRows;
%>
<%
var Repeat13__numRows = 5;
var Repeat13__index = 0;
ahq_numRows += Repeat13__numRows;
%>
<%
var Repeat12__numRows = 5;
var Repeat12__index = 0;
asxs_numRows += Repeat12__numRows;
%>
<%
var Repeat11__numRows = 5;
var Repeat11__index = 0;
adwjdc_numRows += Repeat11__numRows;
%>
<%
var Repeat10__numRows = 5;
var Repeat10__index = 0;
ajfwy_numRows += Repeat10__numRows;
%>
<%
var Repeat9__numRows = 5;
var Repeat9__index = 0;
adfwy_numRows += Repeat9__numRows;
%>
<%
var Repeat8__numRows = 5;
var Repeat8__index = 0;
ajcxy_numRows += Repeat8__numRows;
%>
<%
var Repeat8__numRows = 5;
var Repeat8__index = 0;
adcxy_numRows += Repeat8__numRows;
%>
<%
var Repeat7__numRows = 5;
var Repeat7__index = 0;
ajcxy_numRows += Repeat7__numRows;
%>
<%
var Repeat6__numRows = 5;
var Repeat6__index = 0;
ajpdy_numRows += Repeat6__numRows;
%>
<%
var Repeat5__numRows = 5;
var Repeat5__index = 0;
adpdy_numRows += Repeat5__numRows;
%>
<%
var Repeat4__numRows = 2;
var Repeat4__index = 0;
zaqd_numRows += Repeat4__numRows;
%>
<% var MM_paramName = ""; %>
<%
// *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

// create the list of parameters which should not be maintained
var MM_removeList = "&index=";
if (MM_paramName != "") MM_removeList += "&" + MM_paramName.toLowerCase() + "=";
var MM_keepURL="",MM_keepForm="",MM_keepBoth="",MM_keepNone="";

// add the URL parameters to the MM_keepURL string
for (var items=new Enumerator(Request.QueryString); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepURL += "&" + items.item() + "=" + Server.URLencode(Request.QueryString(items.item()));
  }
}

// add the Form variables to the MM_keepForm string
for (var items=new Enumerator(Request.Form); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepForm += "&" + items.item() + "=" + Server.URLencode(Request.Form(items.item()));
  }
}

// create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL + MM_keepForm;
if (MM_keepBoth.length > 0) MM_keepBoth = MM_keepBoth.substring(1);
if (MM_keepURL.length > 0)  MM_keepURL = MM_keepURL.substring(1);
if (MM_keepForm.length > 0) MM_keepForm = MM_keepForm.substring(1);
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>青大兼职 青岛大学兼职 兼职 大学兼职 青岛大学 纺织服装学院</title>
<style type="text/css">
<!--
.STYLE1 {text-align: center; font-size: 36px; color: #6633CC; }
.zhongjianbiaoti {	font-size: 18px;
}
.STYLE4 {font-size: 24px; }
.STYLE5 {
	color: #000000;
	font-weight: bold;
}
.STYLE6 {font-size: 18px; color: #669999; }
.STYLE8 {color: #FFFFFF}
.STYLE11 {color: #000000}
.STYLE12 {color: #9933FF}
.STYLE15 {font-family: "黑体"}
.STYLE16 {font-size: 36px; color: #0033FF; text-align: center;}
.STYLE17 {color: #0000CC}
.STYLE19 {color: #000000; font-size: 16px; }
.STYLE20 {font-size: 16px}
-->
</style>
</head>
<body bgcolor="#e0e0e0">
	<!--#include file="index-modefy.asp"-->
	<br />
	<table background="" width="100%">
		<tr>
			<td width="12%">
				<table width="122%" height="39%" border="0">
  				<tr>
  					<td>
  						<span class="STYLE17">早安青大</span>
  					</td>
  				</tr>
  				<tr>
  					<td height="29">&nbsp;</td>
  				</tr>
  			</table>
			</td>
			<td width="66%"><% while ((Repeat4__numRows-- != 0) && (!zaqd.EOF)) { %>
    		<table width="100%" border="0">
      		<tr>
        		<td><%=(zaqd.Fields.Item("Content").Value)%></td>
     			</tr><%
					  Repeat4__index++;
					  zaqd.MoveNext();
					}
					%>
				</table>
			</td>
			<td width="22%">
				<table width="100%" height="100%" border="0">
					<tr>
     				<td>
     					<a href="add_articlezaqd.asp">说说</a>您知道的青大新闻
     				</td>
					</tr>
					<tr>
						<td height="35">&nbsp;
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<table width="100%">
		<tr>
			<td width="13%" align="left">
				<img align="left" src="img/青大图标.jpg" width="80" height="80" />
			</td>
			<td width="100%" align="right">
				<div align="center" class="STYLE15"><span class="STYLE16">青岛大学兼职网首页</span></div>
				<div align="center"></div>
			</td>
			<td width="100%" align="right">											   
					<SCRIPT language=JavaScript>
					function Year_Month(){ 
					    var now = new Date(); 
					    var yy = ( now.getYear() < 1900 ) ? ( 1900 + now.getYear() ) : now.getYear();
					    var mm = now.getMonth()+1; 
					    var cl = '<font color="#0000df">'; 
					    if (now.getDay() == 0) cl = '<font color="#c00000">'; 
					    if (now.getDay() == 6) cl = '<font color="#00c000">'; 
					    return(cl +  yy + '年' + mm + '月</font>'); }
					 function Date_of_Today(){ 
					    var now = new Date(); 
					    var cl = '<font color="#ff0000">'; 
					    if (now.getDay() == 0) cl = '<font color="#c00000">'; 
					    if (now.getDay() == 6) cl = '<font color="#00c000">'; 
					    return(cl +  now.getDate() + '</font>'); }
					 function Day_of_Today(){ 
					    var day = new Array(); 
					    day[0] = "星期日"; 
					    day[1] = "星期一"; 
					    day[2] = "星期二"; 
					    day[3] = "星期三"; 
					    day[4] = "星期四"; 
					    day[5] = "星期五"; 
					    day[6] = "星期六"; 
					    var now = new Date(); 
					    var cl = '<font color="#0000df">'; 
					    if (now.getDay() == 0) cl = '<font color="#c00000">'; 
					    if (now.getDay() == 6) cl = '<font color="#00c000">'; 
					    return(cl +  day[now.getDay()] + '</font>'); }
					 function CurentTime(){ 
					    var now = new Date(); 
					    var hh = now.getHours(); 
					    var mm = now.getMinutes(); 
					    var ss = now.getTime() % 60000; 
					    ss = (ss - (ss % 1000)) / 1000; 
					    var clock = hh+':'; 
					    if (mm < 10) clock += '0'; 
					    clock += mm+':'; 
					    if (ss < 10) clock += '0'; 
					    clock += ss; 
					    return(clock); } 
					function refreshCalendarClock(){ 
					document.all.calendarClock1.innerHTML = Year_Month(); 
					document.all.calendarClock2.innerHTML = Date_of_Today(); 
					document.all.calendarClock3.innerHTML = Day_of_Today(); 
					document.all.calendarClock4.innerHTML = CurentTime(); }
					 var webUrl = webUrl; 
					document.write('<table border="0" cellpadding="0" cellspacing="0"><tr><td>'); 
					document.write('<table id="CalendarClockFreeCode" border="0" cellpadding="0" cellspacing="0" width="60" height="70" ');
					document.write('style="position:absolute;visibility:hidden" bgcolor="#eeeeee">');
					document.write('<tr><td align="center"><font ');
					document.write('style="cursor:hand;color:#ff0000;font-family:宋体;font-size:14pt;line-height:120%" ');
					if (webUrl != 'netflower'){ 
					   document.write('</td></tr><tr><td align="center"><font ');
					   document.write('style="cursor:hand;color:#2000ff;font-family:宋体;font-size:9pt;line-height:110%" ');
					} 
					document.write('</td></tr></table>'); 
					document.write('<table border="0" cellpadding="0" cellspacing="0" width="61" bgcolor="#C0C0C0" height="70">');
					document.write('<tr><td valign="top" width="100%" height="100%">');
					document.write('<table border="1" cellpadding="0" cellspacing="0" width="58" bgcolor="#FEFEEF" height="67">');
					document.write('<tr><td align="center" width="100%" height="100%" >');
					document.write('<font id="calendarClock1" style="font-family:宋体;font-size:7pt;line-height:120%"> </font><br>');
					document.write('<font id="calendarClock2" style="color:#ff0000;font-family:Arial;font-size:14pt;line-height:120%"> </font><br>');
					document.write('<font id="calendarClock3" style="font-family:宋体;font-size:9pt;line-height:120%"> </font><br>');
					document.write('<font id="calendarClock4" style="color:#100080;font-family:宋体;font-size:8pt;line-height:120%"><b> </b></font>');
					document.write('</td></tr></table>');
					document.write('</td></tr></table>'); 
					document.write('</td></tr></table>'); 
					setInterval('refreshCalendarClock()',1000);
                                            </SCRIPT>
			</td>
		</tr>
	</table>
	<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" table-layout:fixed>
		<tr>
			<td colspan="2" class="STYLE1" width="90" align="center">&nbsp;</td> 
    </tr>
    <tr width="100%">
    	<td height="35" bgcolor="#FFD495" class="STYLE6 STYLE11">
    		<span class="STYLE4">兼职信息</span>：</td>
      <td bgcolor="#FFD495">
      	<div align="right"><a href="http://www.qdu.edu.cn/" class="STYLE8">转到 &nbsp;我们的校园</a></div>
      </td>
    </tr>
    <tr>
    	<td>短期</td>
    	<td>假期</td>
    </tr>
    <tr>
    	<td height="100%" width="100%">
    		<table width="100%" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
    			<tr>
    				<td class="STYLE12">
    					<table width="100%" border="0">
    						<tr>
                  <td><span class="STYLE19">家教</span></td>
                  <td><div align="right"><a href="alldqjj.asp">查看更多:</a></div></td>
                </tr>
              </table>
            </td>
          </tr><% while ((Repeat1__numRows-- != 0) && (!rsfatie.EOF)) { %>
          <tr>
          	<td><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><a href="viewdqjj.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + rsfatie.Fields.Item("ArticleId").Value %>"><%=(rsfatie.Fields.Item("topic").Value)%></a></div></td>
          	<td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(rsfatie.Fields.Item("content").Value)%></div></td>
          </tr>
          	<%
						  Repeat1__index++;
						  rsfatie.MoveNext();
						}%>
    		</table>
    	</td>
    	<td height="100%" width="100%" colspan="2">
    		<table width="100%" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
    			<tr>
    				<td colspan="2"><table width="100%" border="0">
          <tr>
            <td><span class="STYLE20">家教</span></td>
            <td><div align="right"><a href="alljjj.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
          </tr>
        </table>
    	</td>    	
    </tr>
    <tr>
              <% while ((Repeat2__numRows-- != 0) && (!rsjiaqi.EOF)) { %>
    	<td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><a href="viewjjj.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + rsjiaqi.Fields.Item("ArticleId").Value %>"><%=(rsjiaqi.Fields.Item("topic").Value)%></a></div></td>
    	<td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(rsjiaqi.Fields.Item("content").Value)%></div></td>
    </tr> <%
  Repeat2__index++;
  rsjiaqi.MoveNext();
}
%>
   </table>
  </td>
	</tr>
	<tr>
          <td height="100%"> <table width="460" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2" class="STYLE12"><table width="100%" border="0">
                  <tr>
                    <td><span class="STYLE19">派单员</span></td>
                    <td><div align="right"><a href="adpdy.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
            </tr>
            
              <% while ((Repeat5__numRows-- != 0) && (!adpdy.EOF)) { %>
              <tr>  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vdpdy.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + adpdy.Fields.Item("ArticleId").Value %>"><%=(adpdy.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(adpdy.Fields.Item("content").Value)%></div></td>
            </tr><%
  Repeat5__index++;
  adpdy.MoveNext();
}
%>
          </table></td>
          <td height="100%" colspan="2">
		   <table width="490" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2"><table width="100%" border="0">
                  <tr>
                    <td>派单员</td>
                    <td><div align="right"><a href="ajpdy.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
             </tr>
            <tr>
              <% while ((Repeat6__numRows-- != 0) && (!ajpdy.EOF)) { %>
              <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vjpdy.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + ajpdy.Fields.Item("ArticleId").Value %>"><%=(ajpdy.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(ajpdy.Fields.Item("content").Value)%></div></td>
             </tr> <%
  Repeat6__index++;
  ajpdy.MoveNext();
}
%>
</table></td>
    </tr><tr>
          <td height="100%"> <table width="460" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2" class="STYLE12"><table width="100%" border="0">
                  <tr>
                    <td><span class="STYLE19">促销员</span></td>
                    <td><div align="right"><a href="adcxy.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
            </tr>
                          <% while ((Repeat8__numRows-- != 0) && (!adcxy.EOF)) { %>
              <tr>  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vdcxy.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + adcxy.Fields.Item("ArticleId").Value %>"><%=(adcxy.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(adcxy.Fields.Item("content").Value)%></div></td>
            </tr><%
  Repeat8__index++;
  adcxy.MoveNext();
}
%>
          </table></td>
          <td height="100%" colspan="2">
		   <table width="490" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2"><table width="100%" border="0">
                  <tr>
                    <td>促销员</td>
                    <td><div align="right"><a href="ajcxy.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
             </tr>
            <tr>
              <% while ((Repeat7__numRows-- != 0) && (!ajcxy.EOF)) { %>
<td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vjcxy.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + ajcxy.Fields.Item("ArticleId").Value %>"><%=(ajcxy.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(ajcxy.Fields.Item("content").Value)%></div></td>
             </tr> <%
  Repeat7__index++;
  ajcxy.MoveNext();
}
%>
</table></td>
    </tr><tr>
      <td height="100%"> <table width="460" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2" class="STYLE12"><table width="100%" border="0">
                  <tr>
                    <td><span class="STYLE19">服务员</span></td>
                    <td><div align="right"><a href="adfwy.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
            </tr>
                          <% while ((Repeat9__numRows-- != 0) && (!adfwy.EOF)) { %>
              <tr>  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vdfwy.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + adfwy.Fields.Item("ArticleId").Value %>"><%=(adfwy.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(adfwy.Fields.Item("content").Value)%></div></td>
            </tr><%
  Repeat9__index++;
  adfwy.MoveNext();
}
%>
          </table></td>
          <td height="100%" colspan="2"><table width="490" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
              <tr>
                <td colspan="2"><table width="100%" border="0">
                    <tr>
                      <td>服务员</td>
                      <td><div align="right"><a href="ajfwy.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                    </tr>
                </table></td>
              </tr>
              <% while ((Repeat10__numRows-- != 0) && (!ajfwy.EOF)) { %>
                <tr>
                  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vjfwy.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + ajfwy.Fields.Item("ArticleId").Value %>"><%=(ajfwy.Fields.Item("topic").Value)%></A></div></td>
                  <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(ajfwy.Fields.Item("content").Value)%></div></td>
                </tr>
                <%
  Repeat10__index++;
  ajfwy.MoveNext();
}
%>
</table></td>
    </tr><tr>
      <td height="100%"> <table width="460" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2" class="STYLE12"><table width="100%" border="0">
                  <tr>
                    <td><span class="STYLE19">问卷调查</span></td>
                    <td><div align="right"><a href="adwjdc.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
            </tr>
                          <% while ((Repeat11__numRows-- != 0) && (!adwjdc.EOF)) { %>
              <tr>  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vdwjdc.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + adwjdc.Fields.Item("ArticleId").Value %>"><%=(adwjdc.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(adwjdc.Fields.Item("Topic").Value)%></div></td>
            </tr><%
  Repeat11__index++;
  adwjdc.MoveNext();
}
%>
          </table></td>
          <td height="100%" colspan="2"><table width="490" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
              <tr>
                <td colspan="2"><table width="100%" border="0">
                    <tr>
                      <td>实习生</td>
                      <td><div align="right"><a href="asxs.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                    </tr>
                </table></td>
              </tr>
              <% while ((Repeat12__numRows-- != 0) && (!asxs.EOF)) { %>
                <tr>
                  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vsxs.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + asxs.Fields.Item("ArticleId").Value %>"><%=(asxs.Fields.Item("Topic").Value)%></A></div></td>
                  <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(asxs.Fields.Item("content").Value)%></div></td>
                </tr>
                <%
  Repeat12__index++;
  asxs.MoveNext();
}
%>
</table></td>
    </tr><tr>
      <td height="100%"> <table width="460" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2" class="STYLE12"><table width="100%" border="0">
                  <tr>
                    <td><span class="STYLE11">婚庆</span></td>
                    <td><div align="right"><a href="ahq.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
            </tr>
                          <% while ((Repeat13__numRows-- != 0) && (!ahq.EOF)) { %>
              <tr>  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><a href="vhq.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + ahq.Fields.Item("ArticleId").Value %>"><%=(ahq.Fields.Item("topic").Value)%></a></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(ahq.Fields.Item("Topic").Value)%></div></td>
            </tr><%
  Repeat13__index++;
  ahq.MoveNext();
}
%>
          </table></td>
          <td height="100%" colspan="2"><table width="490" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
              <tr>
                <td colspan="2"><table width="100%" border="0">
                    <tr>
                      <td>其他</td>
                      <td><div align="right"><a href="aqt.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                    </tr>
                </table></td>
              </tr>
              <% while ((Repeat14__numRows-- != 0) && (!aqt.EOF)) { %>
                <tr>
                  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vqt.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + aqt.Fields.Item("ArticleId").Value %>"><%=(aqt.Fields.Item("Topic").Value)%></A></div></td>
                  <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(aqt.Fields.Item("Content").Value)%></div></td>
                </tr>
                <%
  Repeat14__index++;
  aqt.MoveNext();
}
%>
</table></td>
    </tr><tr>
          <td height="100%"> <table width="460" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2" class="STYLE12"><table width="100%" border="0">
                <tr>
                  <td><span class="STYLE11">其他</span></td>
                  <td><div align="right"><a href="adqt.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                </tr>
              </table></td>
            </tr>
            
              <% while ((Repeat15__numRows-- != 0) && (!adqt.EOF)) { %>
              <tr>  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><a href="vdqt.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + adqt.Fields.Item("ArticleId").Value %>"><%=(adqt.Fields.Item("topic").Value)%></a></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(adqt.Fields.Item("content").Value)%></div></td>
            </tr><%
  Repeat15__index++;
  adqt.MoveNext();
}
%>
          </table></td>
          <td height="100%" colspan="2">
		   <table width="490" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2"><table width="100%" border="0">
                  <tr>
                    <td><span class="STYLE20">其他</span></td>
                    <td><div align="right"><a href="ajqt2.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
             </tr>
            <tr>
              <% while ((Repeat16__numRows-- != 0) && (!ajqt2.EOF)) { %>
              <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vjqt2.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + ajqt2.Fields.Item("ArticleId").Value %>"><%=(ajqt2.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(ajqt2.Fields.Item("content").Value)%></div></td>
             </tr> <%
  Repeat16__index++;
  ajqt2.MoveNext();
}
%>
</table></td>
    </tr>
	<tr height="100%" width="100%">
  <td height="100%" width="50%">&nbsp;</td>
      <td height="100%" colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td height="35" colspan="2" bgcolor="#FFD495" class="STYLE6 STYLE11"><span class="STYLE4">校园淘宝</span>：</td>
</tr>
    <tr>
      <td height="180" colspan="2"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" table-layout:fixed>
	  <tr><td>学习</td><td>娱乐</td>
	  </tr>
        <tr>
          <td height="100%"> <table width="460" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2" class="STYLE12"><table width="100%" border="0">
                <tr>
                  <td><span class="STYLE11">二手书</span></td>
                  <td><div align="right"><a href="axess.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                </tr>
              </table></td>
            </tr>
            
              <% while ((Repeat3__numRows-- != 0) && (!xess.EOF)) { %>
              <tr>  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vxess.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + xess.Fields.Item("ArticleId").Value %>"><%=(xess.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(xess.Fields.Item("Content").Value)%></div></td>
            </tr><%
  Repeat3__index++;
  xess.MoveNext();
}
%>
          </table></td>
          <td height="100%" colspan="2">
		   <table width="490" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2"><table width="100%" border="0">
                  <tr>
                    <td><span class="STYLE20">体育用品</span></td>
                    <td><div align="right"><a href="ayyp.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
             </tr>
            <tr>
              <% while ((Repeat17__numRows-- != 0) && (!yp.EOF)) { %>
              <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><a href="vyyp.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + yp.Fields.Item("ArticleId").Value %>"><%=(yp.Fields.Item("Topic").Value)%></a></div></td>
              <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(yp.Fields.Item("Content").Value)%></div></td>
             </tr> <%
  Repeat17__index++;
  yp.MoveNext();
}
%>
</table></td>
    </tr> <tr>
          <td height="100%"> <table width="460" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2" class="STYLE12"><table width="100%" border="0">
                <tr>
                  <td><span class="STYLE11">考研资料</span></td>
                  <td><div align="right"><a href="axky.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                </tr>
              </table></td>
            </tr>
            
              <% while ((Repeat18__numRows-- != 0) && (!axky.EOF)) { %>
              <tr>  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vxky.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + axky.Fields.Item("ArticleId").Value %>"><%=(axky.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(axky.Fields.Item("Content").Value)%></div></td>
            </tr><%
  Repeat18__index++;
  axky.MoveNext();
}
%>
          </table></td>
          <td height="100%" colspan="2">
		   <table width="490" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2"><table width="100%" border="0">
                  <tr>
                    <td><span class="STYLE20">电脑</span></td>
                    <td><div align="right"><a href="aydn.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
             </tr>
            <tr>
              <% while ((Repeat19__numRows-- != 0) && (!aydn.EOF)) { %>
              <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><a href="vydn.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + aydn.Fields.Item("ArticleId").Value %>"><%=(aydn.Fields.Item("Topic").Value)%></a></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(aydn.Fields.Item("Content").Value)%></div></td>
             </tr> <%
  Repeat19__index++;
  aydn.MoveNext();
}
%>
</table></td>
    </tr><tr>
          <td height="100%"> <table width="460" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2" class="STYLE12"><table width="100%" border="0">
                <tr>
                  <td><span class="STYLE11">二专、二外、外语</span></td>
                  <td><div align="right"><a href="axwy.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                </tr>
              </table></td>
            </tr>
            
              <% while ((Repeat20__numRows-- != 0) && (!axwy.EOF)) { %>
              <tr>  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><a href="vxwy.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + axwy.Fields.Item("ArticleId").Value %>"><%=(axwy.Fields.Item("Topic").Value)%></a></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(axwy.Fields.Item("Content").Value)%></div></td>
            </tr><%
  Repeat20__index++;
  axwy.MoveNext();
}
%>
          </table></td>
          <td height="100%" colspan="2">
		   <table width="490" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2"><table width="100%" border="0">
                  <tr>
                    <td><span class="STYLE20">手机</span></td>
                    <td><div align="right"><a href="aysj.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
             </tr>
            <tr>
              <% while ((Repeat21__numRows-- != 0) && (!aysj.EOF)) { %>
              <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vysj.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + aysj.Fields.Item("ArticleId").Value %>"><%=(aysj.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(aysj.Fields.Item("Content").Value)%></div></td>
             </tr> <%
  Repeat21__index++;
  aysj.MoveNext();
}
%>
</table></td>
    </tr><tr>
          <td height="100%"> <table width="460" height="100%" border="0" bordercolor="#666666" cellpadding="01" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2" class="STYLE12"><table width="100%" border="0">
                <tr>
                  <td><span class="STYLE11">其他</span></td>
                  <td><div align="right"><a href="axqt.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                </tr>
              </table></td>
            </tr>
            
              <% while ((Repeat22__numRows-- != 0) && (!axqt.EOF)) { %>
              <tr>  <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vxqt.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + axqt.Fields.Item("ArticleId").Value %>"><%=(axqt.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(axqt.Fields.Item("Content").Value)%></div></td>
            </tr><%
  Repeat22__index++;
  axqt.MoveNext();
}
%>
          </table></td>
          <td height="100%" colspan="2">
		   <table width="490" height="100%" border="0" bordercolor="#666666" cellpadding="1" cellspacing="1" table-layout:fixed>
            <tr>
              <td colspan="2"><table width="100%" border="0">
                  <tr>
                    <td><span class="STYLE11">其他</span></td>
                    <td><div align="right"><a href="ayqt.asp">查看更多:</a>&nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
                </table></td>
             </tr>
            <tr>
              <% while ((Repeat23__numRows-- != 0) && (!ayqt.EOF)) { %>
              <td width="100px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:100px"><A HREF="vyqt.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + ayqt.Fields.Item("ArticleId").Value %>"><%=(ayqt.Fields.Item("Topic").Value)%></A></div></td>
                <td width="350px"><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:350px"><%=(ayqt.Fields.Item("Content").Value)%></div></td>
             </tr> <%
  Repeat23__index++;
  ayqt.MoveNext();
}
%>
</table></td>
    </tr></table></td>
          <td>&nbsp;</td>
        </tr>
        
</table>
  
      <p>&nbsp;</p>
 <p align="center" class="STYLE5"><br /> <br/><br /> <br/><br /> <br/>
 <p class="STYLE5">朋友链接：<a href="http://www.sdjtjz.com/">山东交通学院兼职网</a> &nbsp; &nbsp; <a href="http://www.qdu.edu.cn/">青岛大学欢迎您</a></p>
 <hr />
  <div align="center">地址：青岛市市南区香港东路308号<br />
    咨询电话：15006427725    邮箱：<a href="http://user.qzone.qq.com/1207467967/infocenter">1207467967@qq.com</a><br />
  &copy;青岛大学 All Rights Reserved
  </p>
</div>
  <hr />
</body>
</html>
