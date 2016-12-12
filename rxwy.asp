<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connbbs.asp" -->
<%
// *** Restrict Access To Page: Grant or deny access to this page
var MM_authorizedUsers="";
var MM_authFailedURL="load.asp?error=2";
var MM_grantAccess=false;
if (String(Session("MM_Username")) != "undefined") {
  if (true || (String(Session("MM_UserAuthorization"))=="") || (MM_authorizedUsers.indexOf(String(Session("MM_UserAuthorization"))) >=0)) {
    MM_grantAccess = true;
  }
}
if (!MM_grantAccess) {
  var MM_qsChar = "?";
  if (MM_authFailedURL.indexOf("?") >= 0) MM_qsChar = "&";
  var MM_referrer = Request.ServerVariables("URL");
  if (String(Request.QueryString()).length > 0) MM_referrer = MM_referrer + "?" + String(Request.QueryString());
  MM_authFailedURL = MM_authFailedURL + MM_qsChar + "accessdenied=" + Server.URLEncode(MM_referrer);
  Response.Redirect(MM_authFailedURL);
}
%>
<%
// *** Edit Operations: declare variables

// set the form action variable
var MM_editAction = Request.ServerVariables("SCRIPT_NAME");
if (Request.QueryString) {
  MM_editAction += "?" + Server.HTMLEncode(Request.QueryString);
}

// boolean to abort record edit
var MM_abortEdit = false;

// query string to execute
var MM_editQuery = "";
%>
<%
// *** Insert Record: set variables

if (String(Request("MM_insert")) == "form1") {

  var MM_editConnection = MM_connbbs_STRING;
  var MM_editTable  = "xwy";
  var MM_editRedirectUrl = "vxwy.asp";
  var MM_fieldsStr = "txtTopic|value|txtContent|value|hiddenUsername|value|hiddenIP|value|hiddenArticleID|value|Time|value";
  var MM_columnsStr = "Topic|',none,''|Content|',none,''|Username|',none,''|IP|',none,''|ArticleId|none,none,NULL|ReplyTime|',none,NULL";

  // create the MM_fields and MM_columns arrays
  var MM_fields = MM_fieldsStr.split("|");
  var MM_columns = MM_columnsStr.split("|");
  
  // set the form values
  for (var i=0; i+1 < MM_fields.length; i+=2) {
    MM_fields[i+1] = String(Request.Form(MM_fields[i]));
  }

  // append the query string to the redirect URL
  if (MM_editRedirectUrl && Request.QueryString && Request.QueryString.Count > 0) {
    MM_editRedirectUrl += ((MM_editRedirectUrl.indexOf('?') == -1)?"?":"&") + Request.QueryString;
  }
}
%>
<%
// *** Insert Record: construct a sql insert statement and execute it

if (String(Request("MM_insert")) != "undefined") {

  // create the sql insert statement
  var MM_tableValues = "", MM_dbValues = "";
  for (var i=0; i+1 < MM_fields.length; i+=2) {
    var formVal = MM_fields[i+1];
    var MM_typesArray = MM_columns[i+1].split(",");
    var delim =    (MM_typesArray[0] != "none") ? MM_typesArray[0] : "";
    var altVal =   (MM_typesArray[1] != "none") ? MM_typesArray[1] : "";
    var emptyVal = (MM_typesArray[2] != "none") ? MM_typesArray[2] : "";
    if (formVal == "" || formVal == "undefined") {
      formVal = emptyVal;
    } else {
      if (altVal != "") {
        formVal = altVal;
      } else if (delim == "'") { // escape quotes
        formVal = "'" + formVal.replace(/'/g,"''") + "'";
      } else {
        formVal = delim + formVal + delim;
      }
    }
    MM_tableValues += ((i != 0) ? "," : "") + MM_columns[i];
    MM_dbValues += ((i != 0) ? "," : "") + formVal;
  }
  MM_editQuery = "insert into " + MM_editTable + " (" + MM_tableValues + ") values (" + MM_dbValues + ")";

  if (!MM_abortEdit) {
    // execute the insert
    var MM_editCmd = Server.CreateObject('ADODB.Command');
    MM_editCmd.ActiveConnection = MM_editConnection;
    MM_editCmd.CommandText = MM_editQuery;
    MM_editCmd.Execute();
    MM_editCmd.ActiveConnection.Close();

    if (MM_editRedirectUrl) {
      Response.Redirect(MM_editRedirectUrl);
    }
  }

}
%>
<%
var rsReplies__MMColParam = "1";
if (String(Request.QueryString("ArticleId")) != "undefined" && 
    String(Request.QueryString("ArticleId")) != "") { 
  rsReplies__MMColParam = String(Request.QueryString("ArticleId"));
}
%>
<%
var rsReplies = Server.CreateObject("ADODB.Recordset");
rsReplies.ActiveConnection = MM_connbbs_STRING;
rsReplies.Source = "SELECT ArticleId, Topic FROM xuexiwaiyu WHERE ArticleId = "+ rsReplies__MMColParam.replace(/'/g, "''") + "";
rsReplies.CursorType = 0;
rsReplies.CursorLocation = 2;
rsReplies.LockType = 1;
rsReplies.Open();
var rsReplies_numRows = 0;
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>回复帖子</title>
<script src="SpryAssets/SpryValidationTextField.js" type="text/javascript"></script>
<script src="SpryAssets/SpryValidationTextarea.js" type="text/javascript"></script>
<link href="SpryAssets/SpryValidationTextField.css" rel="stylesheet" type="text/css" />
<link href="SpryAssets/SpryValidationTextarea.css" rel="stylesheet" type="text/css" />
</head>

<body><!--#include file="index-modefy.asp"-->
当前位置：回复帖子
<form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
  <table width="100%" height="131" border="0">
    <tr>
      <td>主题：</td>
      <td><span id="sprytextfield1">
        <label for="txtTopic"></label>
        <input name="txtTopic" type="text" id="txtTopic"  />
      <span class="textfieldRequiredMsg">请填写回复主题(必填)。</span></span></td>
    </tr>
    <tr>
      <td>内容：</td>
      <td><span id="sprytextarea1">
        <label for="txtContent"></label>
        <textarea name="txtContent" id="txtContent" cols="45" rows="5"></textarea>
      <span class="textareaRequiredMsg">请填写回复内容(必填)。</span></span></td>
    </tr>
    <tr>
      <td><input name="hiddenUsername" type="hidden" id="hiddenUsername" value="<%=Session("MM_Username")%>" />
      <input name="hiddenIP" type="hidden" id="hiddenIP" value="<%=Request.ServerVariables("REMOTE_ADDR")%>" />
      <input name="hiddenArticleID" type="hidden" id="hiddenArticleID" value="<%=Request.QueryString("ArticleId")%>" /></td>
      <td><label><input type="submit" name="button" id="button" value="回复" />
      <input type="reset" name="button2" id="button2" value="重置" /><input type='hidden' name='Time' id='Time'/>
         </label></td>
    </tr>
  </table>

  <input type="hidden" name="MM_insert" value="form1">
</form>
<script type="text/javascript">
var sprytextfield1 = new Spry.Widget.ValidationTextField("sprytextfield1");
var sprytextarea1 = new Spry.Widget.ValidationTextarea("sprytextarea1");
</script>
<script>
var d=new Date();
document.getElementById("Time").value=d.getFullYear()+"-"+(d.getMonth()+1)+"-"+d.getDate();
</script>

</body>
</html>
<%
rsReplies.Close();
%>
