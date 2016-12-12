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
<!--#include file="Connections/connbbs.asp" -->
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
  var MM_editTable  = "jiaqipaidanyuan";
  var MM_editRedirectUrl = "ajpdy.asp";
  var MM_fieldsStr = "txtTopic|value|txtContent|value|hiddenUsername|value|hiddenIP|value|Time|value";
  var MM_columnsStr = "Topic|',none,''|Content|',none,''|Username|',none,''|IP|',none,''|PostTime|',none,NULL";

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
var jqpdy = Server.CreateObject("ADODB.Recordset");
jqpdy.ActiveConnection = MM_connbbs_STRING;
jqpdy.Source = "SELECT * FROM jiaqipaidanyuan";
jqpdy.CursorType = 0;
jqpdy.CursorLocation = 2;
jqpdy.LockType = 1;
jqpdy.Open();
var jqpdy_numRows = 0;
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>发表帖子</title>
<script src="SpryAssets/SpryValidationTextField.js" type="text/javascript"></script>
<script src="file:///C|/Documents and Settings/Administrator/Application Data/Adobe/Dreamweaver CS5/zh_CN/Configuration/Temp/Assets/eam37.tmp/SpryValidationTextarea.js" type="text/javascript"></script>
<link href="SpryAssets/SpryValidationTextField.css" rel="stylesheet" type="text/css" />
<link href="file:///C|/Documents and Settings/Administrator/Application Data/Adobe/Dreamweaver CS5/zh_CN/Configuration/Temp/Assets/eam37.tmp/SpryValidationTextarea.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE1 {font-size: 24px}
-->
</style>
<link href="style.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE2 {
	font-size: 20px;
	font-weight: bold;
}
.STYLE3 {color: #FFFFFF}
.STYLE4 {font-size: 10}
-->
</style>
</head>

<body>
<!--#include file="index-modefy.asp"-->
当前位置：发表帖子
<form ACTION="<%=MM_editAction%>" METHOD="POST" id="form1" name="form1">
  <table width="70%" height="243" border="1" cellpadding="1" cellspacing="0" bordercolor="#000000">
    <tr>
      <td><table width="101%" height="239" border="0" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF">
        <tr>
          <td colspan="2" bgcolor="#000000"><div align="center" class="STYLE1 STYLE2 STYLE3">发表帖子</div></td>
        </tr>
        <tr>
          <td width="9%" class="right">主题：</td>
          <td width="91%"><span id="sprytextfield1">
            <label for="text1"></label>
            <input type="text" name="txtTopic" id="txtTopic" />
            <span class="textfieldRequiredMsg">请填写贴子主题。</span></span></td>
        </tr>
        <tr>
          <td class="right">内容：</td>
          <td><span id="sprytextarea1">
            <label for="txtContent"></label>
            <span id="sprytextarea1"><span id="sprytextarea1">
            <textarea name="txtContent" id="txtContent" cols="45" rows="9"></textarea>
          </span></span><span class="textareaRequiredMsg">请填写帖子内容。</span></span></td>
        </tr>
        <tr>
          <td><input name="hiddenUsername" type="hidden" id="hiddenUsername" value="<%=(jqpdy.Fields.Item("Username").Value)%>" />
            <input name="hiddenIP" type="hidden" id="hiddenIP" value="<%=(jqpdy.Fields.Item("IP").Value)%>" /><input type='hidden' name='Time' id='Time'/></td>
          <td><label>
            <input type="submit" name="Submit" value="提交" />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="reset" name="Submit2" value="重置" />
          </label></td>
        </tr>
      </table></td>
    </tr>
  </table>
  


<input type="hidden" name="MM_insert" value="form1">
</form>
<script type="text/javascript">
</script>
<script type="text/javascript">var sprytextfield1 = new Spry.Widget.ValidationTextField("sprytextfield1");
var sprytextarea1 = new Spry.Widget.ValidationTextarea("sprytextarea1");
</script>
<script>
var d=new Date();
document.getElementById("Time").value=d.getFullYear()+"-"+(d.getMonth()+1)+"-"+d.getDate();
</script> 
 
</body>
</html>
<%
jqpdy.Close();
%>
