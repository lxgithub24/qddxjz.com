<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connbbs.asp" -->
<%
// *** Validate request to log in to this site.
var MM_LoginAction = Request.ServerVariables("URL");
if (Request.QueryString!="") MM_LoginAction += "?" + Server.HTMLEncode(Request.QueryString);
var MM_valUsername=String(Request.Form("username"));
if (MM_valUsername != "undefined") {
  var MM_fldUserAuthorization="";
  var MM_redirectLoginSuccess="index.asp";
  var MM_redirectLoginFailed="load.asp";
  var MM_flag="ADODB.Recordset";
  var MM_rsUser = Server.CreateObject(MM_flag);
  MM_rsUser.ActiveConnection = MM_connbbs_STRING;
  MM_rsUser.Source = "SELECT Username, Password";
  if (MM_fldUserAuthorization != "") MM_rsUser.Source += "," + MM_fldUserAuthorization;
  MM_rsUser.Source += " FROM Users WHERE Username='" + MM_valUsername.replace(/'/g, "''") + "' AND Password='" + String(Request.Form("password")).replace(/'/g, "''") + "'";
  MM_rsUser.CursorType = 0;
  MM_rsUser.CursorLocation = 2;
  MM_rsUser.LockType = 3;
  MM_rsUser.Open();
  if (!MM_rsUser.EOF || !MM_rsUser.BOF) {
    // username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername;
    if (MM_fldUserAuthorization != "") {
      Session("MM_UserAuthorization") = String(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value);
    } else {
      Session("MM_UserAuthorization") = "";
    }
    if (String(Request.QueryString("accessdenied")) != "undefined" && false) {
      MM_redirectLoginSuccess = Request.QueryString("accessdenied");
    }
    MM_rsUser.Close();
    Response.Redirect(MM_redirectLoginSuccess);
  }
  MM_rsUser.Close();
  Response.Redirect(MM_redirectLoginFailed);
}
%>
<%
if(String(Request.QueryString("action"))=="logout"){
Session.Contents.Remove("MM_Username");
Session.Contents.Remove("MM_UserAuthorization");
Session.Abandon();
}
var msg="";
if(String(Request.QueryString("error"))!="undefined"){
switch(String(Request.QueryString("error"))){
case"1":
msg="�û�����������󣬵�¼ʧ�ܡ�";
break;
case"2":
msg="��Աר���������û���Ȩʹ�ô���ܡ�";
break;
case"3":
msg="�����ǹ���Աר������ͨ��Ա��Ȩʹ�ô���ܡ�";
break;
}
}
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��½ҳ��</title>
<style type="text/css">
<!--
.STYLE1 {font-size: 24px}
-->
</style>
<link href="style.css" rel="stylesheet" type="text/css" />
</head>

<body bgcolor="#E0E0E0" ><!--#include file="index-modefy.asp"-->
<p align="center" class="STYLE1">�ൺ��ѧ��ְ��Ϣ��վ��½</p>
<form ACTION="<%=MM_LoginAction%>" METHOD="POST" id="form1" name="form1">
  <table width="100%" height="152" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="38%" class="right"><div align="right">�û�����</div></td>
      <td width="62%"><label>
        <input name="username" type="text" id="username" size="20" />
      </label></td>
    </tr>
    <tr>
      <td class="right"><div align="right">���룺</div></td>
      <td><label>
        <input name="password" type="password" id="password" size="21" />
      </label></td>
    </tr>
    <tr>
      <td width="38%">&nbsp;</td>
      <td width="62%"><label>
        <input type="submit" name="Submit" value="�ύ" />
        <input type="reset" name="Submit2" value="����" />
      </label></td>
    </tr>
  </table>

    
</form>
<p align="center" class="error" id="msg"><%= msg %></p>

</body>
</html>