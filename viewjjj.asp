<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connbbs.asp" -->
<%
var rsReply__MMColParam = "1";
if (String(Request.QueryString("ArticleId")) != "undefined" && 
    String(Request.QueryString("ArticleId")) != "") { 
  rsReply__MMColParam = String(Request.QueryString("ArticleId"));
}
%>
<%
var rsReply = Server.CreateObject("ADODB.Recordset");
rsReply.ActiveConnection = MM_connbbs_STRING;
rsReply.Source = "SELECT * FROM jjj WHERE ArticleId = "+ rsReply__MMColParam.replace(/'/g, "''") + "";
rsReply.CursorType = 0;
rsReply.CursorLocation = 2;
rsReply.LockType = 1;
rsReply.Open();
var rsReply_numRows = 0;
%>
<%
var rsjiaqi__MMColParam = "1";
if (String(Request.QueryString("ArticleId")) != "undefined" && 
    String(Request.QueryString("ArticleId")) != "") { 
  rsjiaqi__MMColParam = String(Request.QueryString("ArticleId"));
}
%>
<%
var rsjiaqi = Server.CreateObject("ADODB.Recordset");
rsjiaqi.ActiveConnection = MM_connbbs_STRING;
rsjiaqi.Source = "SELECT * FROM shuqitie WHERE ArticleId = "+ rsjiaqi__MMColParam.replace(/'/g, "''") + "";
rsjiaqi.CursorType = 0;
rsjiaqi.CursorLocation = 2;
rsjiaqi.LockType = 1;
rsjiaqi.Open();
var rsjiaqi_numRows = 0;
%>
<%
var Repeat1__numRows = 10;
var Repeat1__index = 0;
rsReply_numRows += Repeat1__numRows;
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>浏览帖子</title>
<script src="SpryAssets/SprySlidingPanels.js" type="text/javascript"></script>
<link href="SpryAssets/SprySlidingPanels.css" rel="stylesheet" type="text/css"/>
<style type="text/css">
<!--
.STYLE1 {font-size: 18px}
.STYLE2 {color: #006633}
.STYLE3 {color: #000000}
-->
</style>
</head>

<body><!--#include file="index-modefy.asp"--><br /><br /><br /><br />
<span class="STYLE1">当前位置：浏览帖子</span>
<table width="100%" height="161" border="1" cellpadding="1" cellspacing="1" bordercolor="#666666">
  <tr bgcolor="#0C5E9E">
    <td height="30" colspan="2">贴子主题：<%=(rsjiaqi.Fields.Item("Topic").Value)%></td>
    <td>发帖作者：<%=(rsjiaqi.Fields.Item("Username").Value)%></td>
  </tr>
  <tr>
    <td width="12%" height="93">帖子内容：</td>
    <td colspan="2"><%=(rsjiaqi.Fields.Item("Content").Value)%></td>
  </tr>
  <tr>
    <td height="32">&nbsp;</td>
    <td width="69%"><a href="replyjjj.asp?ArticleId=<%= rsjiaqi.Fields.Item("ArticleId").Value %>">回复帖子</a></td>
    <td width="19%">发帖时间：<%=(rsjiaqi.Fields.Item("PostTime").Value)%></td>
  </tr>
</table>
<%
var rsjiaqi_total=rsjiaqi.RecordCount;
 if(!rsReply.EOF||rsReply.BOF){ %>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsjiaqi;
var MM_rsCount   = rsjiaqi_total;
var MM_size      = rsjiaqi_numRows;
var MM_uniqueCol = "";
    MM_paramName = "";
var MM_offset = 0;
var MM_atTotal = false;
var MM_paramIsDefined = (MM_paramName != "" && String(Request(MM_paramName)) != "undefined");
%>
<%
// *** Move To Record: handle 'index' or 'offset' parameter

if (!MM_paramIsDefined && MM_rsCount != 0) {

  // use index parameter if defined, otherwise use offset parameter
  r = String(Request("index"));
  if (r == "undefined") r = String(Request("offset"));
  if (r && r != "undefined") MM_offset = parseInt(r);

  // if we have a record count, check if we are past the end of the recordset
  if (MM_rsCount != -1) {
    if (MM_offset >= MM_rsCount || MM_offset == -1) {  // past end or move last
      if ((MM_rsCount % MM_size) != 0) {  // last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount % MM_size);
      } else {
        MM_offset = MM_rsCount - MM_size;
      }
    }
  }

  // move the cursor to the selected record
  for (var i=0; !MM_rs.EOF && (i < MM_offset || MM_offset == -1); i++) {
    MM_rs.MoveNext();
  }
  if (MM_rs.EOF) MM_offset = i;  // set MM_offset to the last possible record
}
%>
<%
// *** Move To Record: if we dont know the record count, check the display range

if (MM_rsCount == -1) {

  // walk to the end of the display range for this page
  for (var i=MM_offset; !MM_rs.EOF && (MM_size < 0 || i < MM_offset + MM_size); i++) {
    MM_rs.MoveNext();
  }

  // if we walked off the end of the recordset, set MM_rsCount and MM_size
  if (MM_rs.EOF) {
    MM_rsCount = i;
    if (MM_size < 0 || MM_size > MM_rsCount) MM_size = MM_rsCount;
  }

  // if we walked off the end, set the offset based on page size
  if (MM_rs.EOF && !MM_paramIsDefined) {
    if ((MM_rsCount % MM_size) != 0) {  // last page not a full repeat region
      MM_offset = MM_rsCount - (MM_rsCount % MM_size);
    } else {
      MM_offset = MM_rsCount - MM_size;
    }
  }

  // reset the cursor to the beginning
  if (MM_rs.CursorType > 0) {
    if (!MM_rs.BOF) MM_rs.MoveFirst();
  } else {
    MM_rs.Requery();
  }

  // move the cursor to the selected record
  for (var i=0; !MM_rs.EOF && i < MM_offset; i++) {
    MM_rs.MoveNext();
  }
}
%>
<%
// *** Move To Record: update recordset stats

// set the first and last displayed record
rsjiaqi_first = MM_offset + 1;
rsjiaqi_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsjiaqi_first = Math.min(rsjiaqi_first, MM_rsCount);
  rsjiaqi_last  = Math.min(rsjiaqi_last, MM_rsCount);
}

// set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount != -1 && MM_offset + MM_size >= MM_rsCount);
%>
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
<%
// *** Move To Record: set the strings for the first, last, next, and previous links

var MM_moveFirst="",MM_moveLast="",MM_moveNext="",MM_movePrev="";
var MM_keepMove = MM_keepBoth;  // keep both Form and URL parameters for moves
var MM_moveParam = "index";

// if the page has a repeated region, remove 'offset' from the maintained parameters
if (MM_size > 1) {
  MM_moveParam = "offset";
  if (MM_keepMove.length > 0) {
    params = MM_keepMove.split("&");
    MM_keepMove = "";
    for (var i=0; i < params.length; i++) {
      var nextItem = params[i].substring(0,params[i].indexOf("="));
      if (nextItem.toLowerCase() != MM_moveParam) {
        MM_keepMove += "&" + params[i];
      }
    }
    if (MM_keepMove.length > 0) MM_keepMove = MM_keepMove.substring(1);
  }
}

// set the strings for the move to links
if (MM_keepMove.length > 0) MM_keepMove = Server.HTMLEncode(MM_keepMove) + "&";
var urlStr = Request.ServerVariables("URL") + "?" + MM_keepMove + MM_moveParam + "=";
MM_moveFirst = urlStr + "0";
MM_moveLast  = urlStr + "-1";
MM_moveNext  = urlStr + (MM_offset + MM_size);
MM_movePrev  = urlStr + Math.max(MM_offset - MM_size,0);
%>
<p>以下是所有楼层的回复。</p>

<hr />
<% while ((Repeat1__numRows++ != 0) && (!rsReply.EOF)) { %>
  <table width="100%" height="173" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr>
      <td height="28" colspan="2"><span class="STYLE2">回帖主题：<span class="STYLE3"><%=(rsReply.Fields.Item("Topic").Value)%></span></span></td>
      <td>回帖作者：<%=(rsReply.Fields.Item("Username").Value)%></td>
    </tr>
    <tr>
      <td width="12%" height="115">帖子内容：</td>
      <td colspan="2"><%=(rsReply.Fields.Item("Content").Value)%></td>
    </tr>
    <tr>
      <td height="22">第 <%=Repeat1__numRows-10%> 楼</td>
      <td width="68%"><a href="replydjj.asp"></a></td>
      <td width="20%">回帖时间：<%=(rsReply.Fields.Item("ReplyTime").Value)%></td>
    </tr>
</table> <hr />
  <hr />
  <br />
  <%
  Repeat1__index--;
  rsReply.MoveNext();
}
%>
<br /><br /><br />
</body>
</html>
<%
rsReply.Close();
%>
<%
rsjiaqi.Close();}
%>
