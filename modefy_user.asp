<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
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
// *** Update Record: set variables

if (String(Request("MM_update")) == "form1" &&
    String(Request("MM_recordId")) != "undefined") {

  var MM_editConnection = MM_connbbs_STRING;
  var MM_editTable  = "Users";
  var MM_editColumn = "UserId";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "load.asp";
  var MM_fieldsStr = "s1|value|s2|value|s3|value|s4|value|Username|value|txtPassword|value|tNo|value";
  var MM_columnsStr = "institude|',none,''|dean|',none,''|grade|none,none,NULL|class|none,none,NULL|Username|',none,''|Password|',none,''|tNo|',none,''";

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
// *** Update Record: construct a sql update statement and execute it

if (String(Request("MM_update")) != "undefined" &&
    String(Request("MM_recordId")) != "undefined") {

  // create the sql update statement
  MM_editQuery = "update " + MM_editTable + " set ";
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
    MM_editQuery += ((i != 0) ? "," : "") + MM_columns[i] + " = " + formVal;
  }
  MM_editQuery += " where " + MM_editColumn + " = " + MM_recordId;

  if (!MM_abortEdit) {
    // execute the update
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
var rsmodefy = Server.CreateObject("ADODB.Recordset");
rsmodefy.ActiveConnection = MM_connbbs_STRING;
rsmodefy.Source = "SELECT * FROM Users";
rsmodefy.CursorType = 0;
rsmodefy.CursorLocation = 2;
rsmodefy.LockType = 1;
rsmodefy.Open();
var rsmodefy_numRows = 0;
%>
<%
// *** Restrict Access To Page: Grant or deny access to this page
var MM_authorizedUsersmodyfyxinxi="1";
var MM_authFailedURL="load.asp?error=2";
var MM_grantAccess=false;
if (String(Session("MM_Username")) != "undefined") {
  if (true || (String(Session("MM_UserAuthorization"))=="") || (MM_authorizedUsersmodyfyxinxi.indexOf(String(Session("MM_UserAuthorization"))) >=0)) {
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charsmodyfyxinxiet=gb2312" />
<title>������Ϣ�޸�ҳ��</title>
<style type="text/css">
<!--
.hanggeshi {text-align: center;
}
.right {
	text-align: right;
	font-size: 16px;
}
-->
</style>
<script src="SpryAssets/SpryValidationTextField.js" type="text/javascript"></script>
<script src="SpryAssets/SpryValidationPassword.js" type="text/javascript"></script>
<script src="SpryAssets/SpryValidationConfirm.js" type="text/javascript"></script>
<link href="SpryAssets/SpryValidationTextField.css" rel="stylesheet" type="text/css" />
<link href="SpryAssets/SpryValidationPassword.css" rel="stylesheet" type="text/css" />
<link href="SpryAssets/SpryValidationConfirm.css" rel="stylesheet" type="text/css" />
<SCRIPT LANGUAGE="JavaScript">
<!--
function Dsy()
{
 this.Items = {};
}
Dsy.prototype.add = function(id,iArray)
{
 this.Items[id] = iArray;
}
Dsy.prototype.Exists = function(id)
{
 if(typeof(this.Items[id]) == "undefined") return false;
 return true;
}

function change(v){
 var str="0";
 for(i=0;i<v;i++){ str+=("_"+(document.getElementById(s[i]).selectedIndex-1));};
 var ss=document.getElementById(s[v]);
 with(ss){
  length = 0;
  options[0]=new Option(opt0[v],opt0[v]);
  if(v && document.getElementById(s[v-1]).selectedIndex>0 || !v)
  {
   if(dsy.Exists(str)){
    ar = dsy.Items[str];
    for(i=0;i<ar.length;i++)options[length]=new Option(ar[i],ar[i]);
    if(v)options[1].selected = true;
   }
  }
  if(++v<s.length){change(v);}
 }
}

var dsy = new Dsy();

dsy.add("0",["��","����","����","����","��ѧ��ѧ","�����ѧ","���繤��","�Զ�������","��Ϣ����","��ѧ�����뻷��","��֯��װ","ҽ","ʦ��","��","����","������","����","����","�������"]);
dsy.add("0_0",["��������ѧ","����ѧ","���ѧ","�������"]);
dsy.add("0_0_0",["08","09","10","11"]);
dsy.add("0_0_1",["08","09","10","11"]);
dsy.add("0_0_2",["08","09","10","11"]);
dsy.add("0_0_3",["08","09","10","11"]);
dsy.add("0_0_0_0",["01","02"]);
dsy.add("0_0_0_1",["01","02"]);
dsy.add("0_0_0_2",["01","02"]);
dsy.add("0_0_0_3",["01","02"]);
dsy.add("0_0_1_0",["01","02"]);
dsy.add("0_0_1_1",["01","02"]);
dsy.add("0_0_1_2",["01","02"]);
dsy.add("0_0_1_3",["01","02"]);
dsy.add("0_0_2_0",["01","02"]);
dsy.add("0_0_2_1",["01","02"]);
dsy.add("0_0_2_2",["01","02"]);
dsy.add("0_0_2_3",["01","02"]);
dsy.add("0_0_3_0",["01","02"]);
dsy.add("0_0_3_1",["01","02"]);
dsy.add("0_0_3_2",["01","02"]);
dsy.add("0_0_3_3",["01","02"]);
dsy.add("0_1",["Ӣ��","����","����","����","������","��������"]);
dsy.add("0_1_0",["08","09","10","11"]);
dsy.add("0_1_1",["08","09","10","11"]);
dsy.add("0_1_2",["08","09","10","11"]);
dsy.add("0_1_3",["08","09","10","11"]);
dsy.add("0_1_4",["08","09","10","11"]);
dsy.add("0_1_5",["08","09","10","11"]);
dsy.add("0_1_0_0",["01","02"]);
dsy.add("0_1_0_1",["01","02"]);
dsy.add("0_1_0_2",["01","02"]);
dsy.add("0_1_0_3",["01","02"]);
dsy.add("0_1_1_0",["01","02"]);
dsy.add("0_1_1_1",["01","02"]);
dsy.add("0_1_1_2",["01","02"]);
dsy.add("0_1_1_3",["01","02"]);
dsy.add("0_1_2_0",["01","02"]);
dsy.add("0_1_2_1",["01","02"]);
dsy.add("0_1_2_2",["01","02"]);
dsy.add("0_1_2_3",["01","02"]);
dsy.add("0_1_3_0",["01","02"]);
dsy.add("0_1_3_1",["01","02"]);
dsy.add("0_1_3_2",["01","02"]);
dsy.add("0_1_3_3",["01","02"]);
dsy.add("0_1_4_0",["01","02"]);
dsy.add("0_1_4_1",["01","02"]);
dsy.add("0_1_4_2",["01","02"]);
dsy.add("0_1_4_3",["01","02"]);
dsy.add("0_1_5_0",["01","02"]);
dsy.add("0_1_5_1",["01","02"]);
dsy.add("0_1_5_2",["01","02"]);
dsy.add("0_1_5_3",["01","02"]);

dsy.add("0_2",["����ѧ","���ֱ���","������������������"]);
dsy.add("0_2_0",["08","09","10","11"]);
dsy.add("0_2_1",["08","09","10","11"]);
dsy.add("0_2_2",["08","09","10","11"]);

dsy.add("0_2_0_0",["01","02"]);
dsy.add("0_2_0_1",["01","02"]);
dsy.add("0_2_0_2",["01","02"]);

dsy.add("0_2_1_0",["01","02"]);
dsy.add("0_2_1_1",["01","02"]);
dsy.add("0_2_1_2",["01","02"]);

dsy.add("0_2_2_0",["01","02"]);
dsy.add("0_2_2_1",["01","02"]);
dsy.add("0_2_2_2",["01","02"]);

dsy.add("0_3",["�������","�滭"]);
dsy.add("0_3_0",["08","09","10","11"]);
dsy.add("0_3_1",["08","09","10","11"]);

dsy.add("0_3_0_0",["01","02"]);
dsy.add("0_3_0_1",["01","02"]);


dsy.add("0_3_1_0",["01","02"]);
dsy.add("0_3_1_1",["01","02"]);



dsy.add("0_4",["��Ϣ������ѧ","��ѧ��Ӧ����ѧ"]);
dsy.add("0_4_0",["08","09","10","11"]);
dsy.add("0_4_1",["08","09","10","11"]);

dsy.add("0_4_0_0",["01","02"]);
dsy.add("0_4_0_1",["01","02"]);


dsy.add("0_4_1_0",["01","02"]);
dsy.add("0_4_1_1",["01","02"]);

dsy.add("0_5",["Ӧ������ѧ","����ѧ","����Ϣ��ѧ�뼼��","��������"]);
dsy.add("0_5_0",["08","09","10","11"]);
dsy.add("0_5_1",["08","09","10","11"]);
dsy.add("0_5_2",["08","09","10","11"]);
dsy.add("0_5_3",["08","09","10","11"]);
dsy.add("0_5_0_0",["01","02"]);
dsy.add("0_5_0_1",["01","02"]);
dsy.add("0_5_0_2",["01","02"]);
dsy.add("0_5_0_3",["01","02"]);
dsy.add("0_5_1_0",["01","02"]);
dsy.add("0_5_1_1",["01","02"]);
dsy.add("0_5_1_2",["01","02"]);
dsy.add("0_5_1_3",["01","02"]);
dsy.add("0_5_2_0",["01","02"]);
dsy.add("0_5_2_1",["01","02"]);
dsy.add("0_5_2_2",["01","02"]);
dsy.add("0_5_2_3",["01","02"]);
dsy.add("0_5_3_0",["01","02"]);
dsy.add("0_5_3_1",["01","02"]);
dsy.add("0_5_3_2",["01","02"]);
dsy.add("0_5_3_3",["01","02"]);
dsy.add("0_6",["��е���̼��Զ���","��ҵ���","�����붯������"]);
dsy.add("0_6_0",["08","09","10","11"]);
dsy.add("0_6_1",["08","09","10","11"]);
dsy.add("0_6_2",["08","09","10","11"]);

dsy.add("0_6_0_0",["01","02"]);
dsy.add("0_6_0_1",["01","02"]);
dsy.add("0_6_0_2",["01","02"]);

dsy.add("0_6_1_0",["01","02"]);
dsy.add("0_6_1_1",["01","02"]);
dsy.add("0_6_1_2",["01","02"]);

dsy.add("0_6_2_0",["01","02"]);
dsy.add("0_6_2_1",["01","02"]);
dsy.add("0_6_2_2",["01","02"]);

dsy.add("0_7",["�������̼����Զ���","�Զ���","������Ϣ����","ͨ�Ź���","������Ϣ��ѧ�뼼��"]);
dsy.add("0_7_0",["08","09","10","11"]);
dsy.add("0_7_1",["08","09","10","11"]);
dsy.add("0_7_2",["08","09","10","11"]);
dsy.add("0_7_3",["08","09","10","11"]);
dsy.add("0_7_4",["08","09","10","11"]);

dsy.add("0_7_0_0",["01","02"]);
dsy.add("0_7_0_1",["01","02"]);
dsy.add("0_7_0_2",["01","02"]);
dsy.add("0_7_0_3",["01","02"]);
dsy.add("0_7_1_0",["01","02"]);
dsy.add("0_7_1_1",["01","02"]);
dsy.add("0_7_1_2",["01","02"]);
dsy.add("0_7_1_3",["01","02"]);
dsy.add("0_7_2_0",["01","02"]);
dsy.add("0_7_2_1",["01","02"]);
dsy.add("0_7_2_2",["01","02"]);
dsy.add("0_7_2_3",["01","02"]);
dsy.add("0_7_3_0",["01","02"]);
dsy.add("0_7_3_1",["01","02"]);
dsy.add("0_7_3_2",["01","02"]);
dsy.add("0_7_3_3",["01","02"]);
dsy.add("0_7_4_0",["01","02"]);
dsy.add("0_7_4_1",["01","02"]);
dsy.add("0_7_4_2",["01","02"]);
dsy.add("0_7_4_3",["01","02"]);


dsy.add("0_8",["�������ѧ","���繤��","�������"]);
dsy.add("0_8_0",["08","09","10","11"]);
dsy.add("0_8_1",["08","09","10","11"]);
dsy.add("0_8_2",["08","09","10","11"]);

dsy.add("0_8_0_0",["01","02"]);
dsy.add("0_8_0_1",["01","02"]);
dsy.add("0_8_0_2",["01","02"]);

dsy.add("0_8_1_0",["01","02"]);
dsy.add("0_8_1_1",["01","02"]);
dsy.add("0_8_1_2",["01","02"]);

dsy.add("0_8_2_0",["01","02"]);
dsy.add("0_8_2_1",["01","02"]);
dsy.add("0_8_2_2",["01","02"]);
dsy.add("0_9",["Ӧ�û�ѧ","�߷��Ӳ����빤��","��ѧ�����빤��","�ữ����","��ѧ","������ѧ","��������"]);
dsy.add("0_9_0",["08","09","10","11"]);
dsy.add("0_9_1",["08","09","10","11"]);
dsy.add("0_9_2",["08","09","10","11"]);
dsy.add("0_9_3",["08","09","10","11"]);
dsy.add("0_9_4",["08","09","10","11"]);
dsy.add("0_9_5",["08","09","10","11"]);
dsy.add("0_9_6",["08","09","10","11"]);
dsy.add("0_9_0_0",["01","02"]);
dsy.add("0_9_0_1",["01","02"]);
dsy.add("0_9_0_2",["01","02"]);
dsy.add("0_9_0_3",["01","02"]);
dsy.add("0_9_1_0",["01","02"]);
dsy.add("0_9_1_1",["01","02"]);
dsy.add("0_9_1_2",["01","02"]);
dsy.add("0_9_1_3",["01","02"]);
dsy.add("0_9_2_0",["01","02"]);
dsy.add("0_9_2_1",["01","02"]);
dsy.add("0_9_2_2",["01","02"]);
dsy.add("0_9_2_3",["01","02"]);
dsy.add("0_9_3_0",["01","02"]);
dsy.add("0_9_3_1",["01","02"]);
dsy.add("0_9_3_2",["01","02"]);
dsy.add("0_9_3_3",["01","02"]);
dsy.add("0_9_4_0",["01","02"]);
dsy.add("0_9_4_1",["01","02"]);
dsy.add("0_9_4_2",["01","02"]);
dsy.add("0_9_4_3",["01","02"]);
dsy.add("0_9_5_0",["01","02"]);
dsy.add("0_9_5_1",["01","02"]);
dsy.add("0_9_5_2",["01","02"]);
dsy.add("0_9_5_3",["01","02"]);
dsy.add("0_9_6_0",["01","02"]);
dsy.add("0_9_6_1",["01","02"]);
dsy.add("0_9_6_2",["01","02"]);
dsy.add("0_9_6_3",["01","02"]);
dsy.add("0_10",["��֯����","��װ����빤��","��װ���","��װ����"]);
dsy.add("0_10_0",["08","09","10","11"]);
dsy.add("0_10_1",["08","09","10","11"]);
dsy.add("0_10_2",["08","09","10","11"]);
dsy.add("0_10_3",["08","09","10","11"]);
dsy.add("0_10_0_0",["01","02","03","04","05"]);

dsy.add("0_10_0_1",["01","02","03","04","05"]);
dsy.add("0_10_0_2",["01","02","03","04","05"]);
dsy.add("0_10_0_3",["01","02","03","04","05"]);
dsy.add("0_10_1_0",["01","02"]);
dsy.add("0_10_1_1",["01","02"]);
dsy.add("0_10_1_2",["01","02"]);
dsy.add("0_10_1_3",["01","02"]);
dsy.add("0_10_2_0",["01","02"]);
dsy.add("0_10_2_1",["01","02"]);
dsy.add("0_10_2_2",["01","02"]);
dsy.add("0_10_2_3",["01","02"]);
dsy.add("0_10_3_0",["01","02"]);
dsy.add("0_10_3_1",["01","02"]);
dsy.add("0_10_3_2",["01","02"]);
dsy.add("0_10_3_3",["01","02"]);
dsy.add("0_11",["�ٴ�ҽѧ7","�ٴ�ҽѧ5","Ԥ��ҽѧ","ҽѧӰ��ѧ","ҽѧ����","��ǻҽѧ","����ѧ","ҩѧ","���＼��","ʳƷ��ѧ�빤��"]);
dsy.add("0_11_0",["05","06","07","08","09","10","11"]);
dsy.add("0_11_1",["07","08","09","10","11"]);
dsy.add("0_11_2",["08","09","10","11"]);
dsy.add("0_11_3",["08","09","10","11"]);
dsy.add("0_11_4",["08","09","10","11"]);
dsy.add("0_11_5",["08","09","10","11"]);
dsy.add("0_11_6",["08","09","10","11"]);
dsy.add("0_11_7",["08","09","10","11"]);
dsy.add("0_11_8",["08","09","10","11"]);
dsy.add("0_11_9",["08","09","10","11"]);
dsy.add("0_11_0_0",["01","02"]);
dsy.add("0_11_0_1",["01","02"]);
dsy.add("0_11_0_2",["01","02"]);
dsy.add("0_11_0_3",["01","02"]);
dsy.add("0_11_0_4",["01","02"]);
dsy.add("0_11_0_5",["01","02"]);
dsy.add("0_11_0_6",["01","02"]);
dsy.add("0_11_1_0",["01","02"]);
dsy.add("0_11_1_1",["01","02"]);
dsy.add("0_11_1_2",["01","02"]);
dsy.add("0_11_1_3",["01","02"]);
dsy.add("0_11_1_4",["01","02"]);
dsy.add("0_11_2_0",["01","02"]);
dsy.add("0_11_2_1",["01","02"]);
dsy.add("0_11_2_2",["01","02"]);
dsy.add("0_11_2_3",["01","02"]);
dsy.add("0_11_3_0",["01","02"]);
dsy.add("0_11_3_1",["01","02"]);
dsy.add("0_11_3_2",["01","02"]);
dsy.add("0_11_3_3",["01","02"]);
dsy.add("0_11_4_0",["01","02"]);
dsy.add("0_11_4_1",["01","02"]);
dsy.add("0_11_4_2",["01","02"]);
dsy.add("0_11_4_3",["01","02"]);
dsy.add("0_11_5_0",["01","02"]);
dsy.add("0_11_5_1",["01","02"]);
dsy.add("0_11_5_2",["01","02"]);
dsy.add("0_11_5_3",["01","02"]);
dsy.add("0_11_6_0",["01","02"]);
dsy.add("0_11_6_1",["01","02"]);
dsy.add("0_11_6_2",["01","02"]);
dsy.add("0_11_6_3",["01","02"]);
dsy.add("0_11_7_0",["01","02"]);
dsy.add("0_11_7_1",["01","02"]);
dsy.add("0_11_7_2",["01","02"]);
dsy.add("0_11_7_3",["01","02"]);
dsy.add("0_11_8_0",["01","02"]);
dsy.add("0_11_8_1",["01","02"]);
dsy.add("0_11_8_2",["01","02"]);
dsy.add("0_11_8_3",["01","02"]);
dsy.add("0_11_9_0",["01","02"]);
dsy.add("0_11_9_1",["01","02"]);
dsy.add("0_11_9_2",["01","02"]);
dsy.add("0_11_9_3",["01","02"]);
dsy.add("0_12",["��ѧ","��������ѧ","Ӣ��","��ѧ��Ӧ����ѧ","����ѧ","��ѧ","˼�����ν���","��ʷѧ","�����ѧ","��������ѧ","Ӧ������ѧ","Сѧ����","ѧǰ����","��������"]);
dsy.add("0_12_0",["08","09","10","11"]);
dsy.add("0_12_1",["08","09","10","11"]);
dsy.add("0_12_2",["08","09","10","11"]);
dsy.add("0_12_3",["08","09","10","11"]);
dsy.add("0_12_4",["08","09","10","11"]);
dsy.add("0_12_5",["08","09","10","11"]);
dsy.add("0_12_6",["08","09","10","11"]);
dsy.add("0_12_7",["08","09","10","11"]);
dsy.add("0_12_8",["08","09","10","11"]);
dsy.add("0_12_9",["08","09","10","11"]);
dsy.add("0_12_10",["08","09","10","11"]);
dsy.add("0_12_11",["08","09","10","11"]);
dsy.add("0_12_12",["08","09","10","11"]);
dsy.add("0_12_13",["08","09","10","11"]);
dsy.add("0_12_0_0",["01","02"]);
dsy.add("0_12_0_1",["01","02"]);
dsy.add("0_12_0_2",["01","02"]);
dsy.add("0_12_0_3",["01","02"]);

dsy.add("0_12_1_0",["01","02"]);
dsy.add("0_12_1_1",["01","02"]);
dsy.add("0_12_1_2",["01","02"]);
dsy.add("0_12_1_3",["01","02"]);

dsy.add("0_12_2_0",["01","02"]);
dsy.add("0_12_2_1",["01","02"]);
dsy.add("0_12_2_2",["01","02"]);
dsy.add("0_12_2_3",["01","02"]);
dsy.add("0_12_3_0",["01","02"]);
dsy.add("0_12_3_1",["01","02"]);
dsy.add("0_12_3_2",["01","02"]);
dsy.add("0_12_3_3",["01","02"]);
dsy.add("0_12_4_0",["01","02"]);
dsy.add("0_12_4_1",["01","02"]);
dsy.add("0_12_4_2",["01","02"]);
dsy.add("0_12_4_3",["01","02"]);
dsy.add("0_12_5_0",["01","02"]);
dsy.add("0_12_5_1",["01","02"]);
dsy.add("0_12_5_2",["01","02"]);
dsy.add("0_12_5_3",["01","02"]);
dsy.add("0_12_6_0",["01","02"]);
dsy.add("0_12_6_1",["01","02"]);
dsy.add("0_12_6_2",["01","02"]);
dsy.add("0_12_6_3",["01","02"]);
dsy.add("0_12_7_0",["01","02"]);
dsy.add("0_12_7_1",["01","02"]);
dsy.add("0_12_7_2",["01","02"]);
dsy.add("0_12_7_3",["01","02"]);
dsy.add("0_12_8_0",["01","02"]);
dsy.add("0_12_8_1",["01","02"]);
dsy.add("0_12_8_2",["01","02"]);
dsy.add("0_12_8_3",["01","02"]);
dsy.add("0_12_9_0",["01","02"]);
dsy.add("0_12_9_1",["01","02"]);
dsy.add("0_12_9_2",["01","02"]);
dsy.add("0_12_9_3",["01","02"]);
dsy.add("0_12_10_0",["01","02"]);
dsy.add("0_12_10_1",["01","02"]);
dsy.add("0_12_10_2",["01","02"]);
dsy.add("0_12_10_3",["01","02"]);
dsy.add("0_12_11_0",["01","02"]);
dsy.add("0_12_11_1",["01","02"]);
dsy.add("0_12_11_2",["01","02"]);
dsy.add("0_12_11_3",["01","02"]);
dsy.add("0_12_12_0",["01","02"]);
dsy.add("0_12_12_1",["01","02"]);
dsy.add("0_12_12_2",["01","02"]);
dsy.add("0_12_12_3",["01","02"]);
dsy.add("0_12_13_0",["01","02"]);
dsy.add("0_12_13_1",["01","02"]);
dsy.add("0_12_13_2",["01","02"]);
dsy.add("0_12_13_3",["01","02"]);
dsy.add("0_13",["��ѧ","��Ṥ��","��������","����ѧ������ѧ"]);
dsy.add("0_13_0",["08","09","10","11"]);
dsy.add("0_13_1",["08","09","10","11"]);
dsy.add("0_13_2",["08","09","10","11"]);
dsy.add("0_13_3",["08","09","10","11"]);
dsy.add("0_13_0_0",["01","02"]);
dsy.add("0_13_0_1",["01","02"]);
dsy.add("0_13_0_2",["01","02"]);
dsy.add("0_13_0_3",["01","02"]);
dsy.add("0_13_1_0",["01","02"]);
dsy.add("0_13_1_1",["01","02"]);
dsy.add("0_13_1_2",["01","02"]);
dsy.add("0_13_1_3",["01","02"]);
dsy.add("0_13_2_0",["01","02"]);
dsy.add("0_13_2_1",["01","02"]);
dsy.add("0_13_2_2",["01","02"]);
dsy.add("0_13_2_3",["01","02"]);
dsy.add("0_13_3_0",["01","02"]);
dsy.add("0_13_3_1",["01","02"]);
dsy.add("0_13_3_2",["01","02"]);
dsy.add("0_13_3_3",["01","02"]);



dsy.add("0_14",["����ѧ","����ѧ","����ѧ","����","ͳ��ѧ"]);
dsy.add("0_14_0",["08","09","10","11"]);
dsy.add("0_14_1",["08","09","10","11"]);
dsy.add("0_14_2",["08","09","10","11"]);
dsy.add("0_14_3",["08","09","10","11"]);
dsy.add("0_14_4",["08","09","10","11"]);

dsy.add("0_14_0_0",["01","02"]);
dsy.add("0_14_0_1",["01","02"]);
dsy.add("0_14_0_2",["01","02"]);
dsy.add("0_14_0_3",["01","02"]);
dsy.add("0_14_1_0",["01","02"]);
dsy.add("0_14_1_1",["01","02"]);
dsy.add("0_14_1_2",["01","02"]);
dsy.add("0_14_1_3",["01","02"]);
dsy.add("0_14_2_0",["01","02"]);
dsy.add("0_14_2_1",["01","02"]);
dsy.add("0_14_2_2",["01","02"]);
dsy.add("0_14_2_3",["01","02"]);
dsy.add("0_14_3_0",["01","02"]);
dsy.add("0_14_3_1",["01","02"]);
dsy.add("0_14_3_2",["01","02"]);
dsy.add("0_14_3_3",["01","02"]);
dsy.add("0_14_4_0",["01","02"]);
dsy.add("0_14_4_1",["01","02"]);
dsy.add("0_14_4_2",["01","02"]);
dsy.add("0_14_4_3",["01","02"]);

dsy.add("0_15",["���ʾ�����ó��","��������","���̹���","��Ϣ��������Ϣϵͳ","��������"]);
dsy.add("0_15_0",["08","09","10","11"]);
dsy.add("0_15_1",["08","09","10","11"]);
dsy.add("0_15_2",["08","09","10","11"]);
dsy.add("0_15_3",["08","09","10","11"]);
dsy.add("0_15_4",["08","09","10","11"]);

dsy.add("0_15_0_0",["01","02"]);
dsy.add("0_15_0_1",["01","02"]);
dsy.add("0_15_0_2",["01","02"]);
dsy.add("0_15_0_3",["01","02"]);
dsy.add("0_15_1_0",["01","02"]);
dsy.add("0_15_1_1",["01","02"]);
dsy.add("0_15_1_2",["01","02"]);
dsy.add("0_15_1_3",["01","02"]);
dsy.add("0_15_2_0",["01","02"]);
dsy.add("0_15_2_1",["01","02"]);
dsy.add("0_15_2_2",["01","02"]);
dsy.add("0_15_2_3",["01","02"]);
dsy.add("0_15_3_0",["01","02"]);
dsy.add("0_15_3_1",["01","02"]);
dsy.add("0_15_3_2",["01","02"]);
dsy.add("0_15_3_3",["01","02"]);
dsy.add("0_15_4_0",["01","02"]);
dsy.add("0_15_4_1",["01","02"]);
dsy.add("0_15_4_2",["01","02"]);
dsy.add("0_15_4_3",["01","02"]);
dsy.add("0_16",["���ι���"]);
dsy.add("0_16_0",["08","09","10","11"]);

dsy.add("0_16_0_0",["01","02"]);
dsy.add("0_16_0_1",["01","02"]);

dsy.add("0_17",["���ʾ�����ó��","���ι���","���ѧ","������","Ӣ��","��������"]);
dsy.add("0_17_0",["08","09","10","11"]);
dsy.add("0_17_1",["08","09","10","11"]);
dsy.add("0_17_2",["08","09","10","11"]);
dsy.add("0_17_3",["08","09","10","11"]);
dsy.add("0_17_4",["08","09","10","11"]);
dsy.add("0_17_5",["08","09","10","11"]);
dsy.add("0_17_0_0",["01","02"]);
dsy.add("0_17_0_1",["01","02"]);
dsy.add("0_17_0_2",["01","02"]);
dsy.add("0_17_0_3",["01","02"]);
dsy.add("0_17_1_0",["01","02"]);
dsy.add("0_17_1_1",["01","02"]);
dsy.add("0_17_1_2",["01","02"]);
dsy.add("0_17_1_3",["01","02"]);
dsy.add("0_17_2_0",["01","02"]);
dsy.add("0_17_2_1",["01","02"]);
dsy.add("0_17_2_2",["01","02"]);
dsy.add("0_17_2_3",["01","02"]);
dsy.add("0_17_3_0",["01","02"]);
dsy.add("0_17_3_1",["01","02"]);
dsy.add("0_17_3_2",["01","02"]);
dsy.add("0_17_3_3",["01","02"]);
dsy.add("0_17_4_0",["01","02"]);
dsy.add("0_17_4_1",["01","02"]);
dsy.add("0_17_4_2",["01","02"]);
dsy.add("0_17_4_3",["01","02"]);
dsy.add("0_17_5_0",["01","02"]);
dsy.add("0_17_5_1",["01","02"]);
dsy.add("0_17_5_2",["01","02"]);
dsy.add("0_17_5_3",["01","02"]);
dsy.add("0_18",["����ý������"]);
dsy.add("0_18_0",["08","09","10","11"]);

dsy.add("0_18_0_0",["01","02"]);
dsy.add("0_18_0_1",["01","02"]);

//-->
</SCRIPT>
<SCRIPT LANGUAGE = JavaScript>
<!--
//** liuxiang
//

var s=["s1","s2","s3","s4"];
var opt0 = ["Ժ","ϵ","�꼶","�༶"];
function setup()
{
 for(i=0;i<s.length-1;i++)
  document.getElementById(s[i]).onchange=new Function("change("+(i+1)+")");
 change(0);
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</SCRIPT>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE7 {font-size: 14pt}
.STYLE2 {font-size: 12pt}
.STYLE9 {font-size: 16px}
.STYLE10 {font-size: 20px}
-->
</style>
</head>
<script language="javascript">
function check()
{
 var txtPassword=document.getElementById("txtPassword").value;
 
 var password2=document.getElementById("password2").value;
 if(txtPassword!=password2)
 {
  alert("�����������벻һ�£�");
  return false;
 }
 return true;
}
</script>
<body  bgcolor="#E0E0E0" onload="setup()">
<!--#include file="index-modefy.asp"-->
<% if(String(Request.QueryString("update"))!="undefined"){ %>
<p align="center">���������Ѹ��¡�</p><% } %>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1" onSubmit="return check()">
  <table width="100%" height="307" border="0" cellpadding="0" cellspacing="0">
     <tr>
      <td colspan="2"><span class="STYLE10">������Ϣ�޸�ҳ��</span></td>
    </tr>
	<tr>
      <td colspan="2"><div align="center">
        <p class="hanggeshi"> <span class="STYLE9">�ൺ��ѧ</span>--&gt;��   
          <select name="s1" id="s1">
            <option value="ID">Ժ</option>
          </select>
          <select name="s2" size="1" id="s2">
            <option value="ID">ϵ</option>
          </select>
          <select name="s3" size="1" id="s3">
            <option value="ID">�꼶</option>
          </select>
          <select name="s4" size="1" id="s4">
            <option value="ID">�༶</option>
          </select> 
          </p>
        </div></td>
    </tr>
    <tr>
      <td width="40%" class="right">ԭ�û�����</td>	
       <td width="60%"><%=Session("MM_username")%></td>	
    </tr>
    <tr>
      <td height="28" class="right">���û�����</td>

      <td><span id="sprytextfield1">
        <label for="text1"></label>
        <input type="Username" name="Username" id="Username" />
        �������û���.(25��������)</span></td>
    </tr>
    <tr>
      <td height="27" class="right">���룺</td>
      <td><span id="sprypassword1">
        <label for="password1"></label>
        <input type="password" name="txtPassword" id="txtPassword" />
        ���������롣(50�ַ�����)</span></td>
    </tr>
    <tr>
      <td height="27" class="right">ȷ�����룺</td>
      <td><span id="spryconfirm1">
        <label for="password2"></label>
        <span id="spryconfirm1">
        <input type="password" name="txtConfirmPassword" id="password2" />
        </span>���ٴ��������롣 (ͬ��) </span></td>
    </tr>
    <tr>
      <td height="26" class="right">��ϵ��ʽ��</td>
      <td><span id="sprytextfield3">
        <label for="text2"></label>
        <input type="text" name="tNo" id="tNo" />
        ��������ϵ��ʽ��(�̶��绰���ֻ�) </span></td>
    </tr>
    <tr>
      <td height="32">&nbsp;</td>
      <td><label>
        <input type="submit" name="Submit" value="����" />
        <input type="reset" name="Submit2" value="����" />
      </label></td>
    </tr>
    <tr>
      <td><div align="right"><span class="STYLE1"><span class="STYLE7">��ܰ��ʾ</span>��</span></div></td>
      <td><span class="STYLE2">������Ϣ����ȷ���������룬���򽫲���ҳ��򿪴������⡣ֻҪ������䣬������Ϣ����̫��Ҫ�󡣡�������ѧ�Ž����ܷ����仯������^o^</span></td>
    </tr>
  </table>
        <input type="hidden" name="MM_insert" value="form1">    
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rsmodefy.Fields.Item("UserId").Value %>">
</form>
<script type="text/javascript">
var sprytextfield1 = new Spry.Widget.ValidationTextField("sprytextfield1");
var sprytextfield2 = new Spry.Widget.ValidationTextField("sprytextfield2");
var sprypassword1 = new Spry.Widget.ValidationPassword("sprypassword1");
var spryconfirm1 = new Spry.Widget.ValidationConfirm("spryconfirm1","txtPassword",{validateOn:["blur"]});
var sprytextfield3 = new Spry.Widget.ValidationTextField("sprytextfield3");
</script>

</body>
</html>
<%
rsmodefy.Close();
%>
