<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connbbs.asp" -->
<%
var rsshuqitie = Server.CreateObject("ADODB.Recordset");
rsshuqitie.ActiveConnection = MM_connbbs_STRING;
rsshuqitie.Source = "SELECT * FROM shuqitie ORDER BY PostTime DESC";
rsshuqitie.CursorType = 0;
rsshuqitie.CursorLocation = 2;
rsshuqitie.LockType = 1;
rsshuqitie.Open();
var rsshuqitie_numRows = 0;
%>
<%
var Repeat1__numRows = 15;
var Repeat1__index = 0;
rsshuqitie_numRows += Repeat1__numRows;
%>
<%
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsshuqitie_total = rsshuqitie.RecordCount;

// set the number of rows displayed on this page
if (rsshuqitie_numRows < 0) {            // if repeat region set to all records
  rsshuqitie_numRows = rsshuqitie_total;
} else if (rsshuqitie_numRows == 0) {    // if no repeat regions
  rsshuqitie_numRows = 1;
}

// set the first and last displayed record
var rsshuqitie_first = 1;
var rsshuqitie_last  = rsshuqitie_first + rsshuqitie_numRows - 1;

// if we have the correct record count, check the other stats
if (rsshuqitie_total != -1) {
  rsshuqitie_numRows = Math.min(rsshuqitie_numRows, rsshuqitie_total);
  rsshuqitie_first   = Math.min(rsshuqitie_first, rsshuqitie_total);
  rsshuqitie_last    = Math.min(rsshuqitie_last, rsshuqitie_total);
}
%>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsshuqitie;
var MM_rsCount   = rsshuqitie_total;
var MM_size      = rsshuqitie_numRows;
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
rsshuqitie_first = MM_offset + 1;
rsshuqitie_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsshuqitie_first = Math.min(rsshuqitie_first, MM_rsCount);
  rsshuqitie_last  = Math.min(rsshuqitie_last, MM_rsCount);
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>���ڳ��ڼҽ�ȫ����ְ��Ϣ</title>
<style type="text/css">
</style><link href="style.css" rel="stylesheet" type="text/css" />
</head>

<body><!--#include file="index-modefy.asp"--><br /><br /><br />
<table width="950" height="50" border="0" bgcolor="#FFD495" cellspacing="0" cellpadding="0">
  <tr>
    <td><span class="STYLE3">��ǰҳ�棺���ڳ��ڼҽ�ȫ����ְ��Ϣ</span></td>
    <td><div align="right"><a href="add-article2.asp" class="STYLE3">��Ա����</a></div></td>
  </tr>
</table>
<table width="950" height="60" border="0" cellpadding="1" cellspacing="1" bordercolor="#666666" rules="rows">  <tr>
    <td width="188">��������</td>
    <td width="446">��������</td>
    <td width="136">��������</td>
    <td width="157">����ʱ��</td>
  </tr>
  <tr>
    <% while ((Repeat1__numRows-- != 0) && (!rsshuqitie.EOF)) { %>
  <td><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:320px"><A HREF="viewjjj.asp?<%= Server.HTMLEncode(MM_keepNone) + ((MM_keepNone!="")?"&":"") + "ArticleId=" + rsshuqitie.Fields.Item("ArticleId").Value %>"><%=(rsshuqitie.Fields.Item("topic").Value)%></A></div></td>
  <td><div style="text-overflow:ellipsis;overflow:hidden;word-break:keep-all;white-space:nowrap;width:320px"><%=(rsshuqitie.Fields.Item("content").Value)%></div></td>
  <td><%=(rsshuqitie.Fields.Item("Username").Value)%></td>
  <td><%=(rsshuqitie.Fields.Item("PostTime").Value)%></td>
 </tr> <%
  Repeat1__index++;
  rsshuqitie.MoveNext();
}
%>

</table>
<br /><br /><br /><hr />
</body>
</html>
<%
rsshuqitie.Close();
%>
