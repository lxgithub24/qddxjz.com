<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connsqluserssqlusers.asp" -->
<%
var rssqlusers__MMColParam = "1";
if (String(Request.QueryString("学号")) != "undefined" && 
    String(Request.QueryString("学号")) != "") { 
  rssqlusers__MMColParam = String(Request.QueryString("学号"));
}
%>
<%
var rssqlusers = Server.CreateObject("ADODB.Recordset");
rssqlusers.ActiveConnection = MM_connsqluserssqlusers_STRING;
rssqlusers.Source = "SELECT 学号  FROM 青兼信息列  WHERE 学号 = '"+ rssqlusers__MMColParam.replace(/'/g, "''") + "'";
rssqlusers.CurssqlusersorType = 0;
rssqlusers.CurssqlusersorLocation = 2;
rssqlusers.LockType = 1;
rssqlusers.Open();
var rssqlusers_numRows = 0;
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<?xml verssqlusersion="1.0"encoding"gb2312"?>
<%
Response.Expires=0;
Response.ExpiresAbsolute="1980-1-1";
Response.AddHeader("progma","no-cache");
Response.AddHeader("cache-control","private");
Response.CacheControl="no-cache";
Response.ContentType="text/xml";
var str =""+String(Request.QueryString("sNo"));
if(!rssqlusers.EOF){
str+="已被注册";
}else if(String(Request.QueryString("sNo"))!="undefined"){
str+="可用";}
%><result><value><%=str%></value></result>
</html>
<%
rssqlusers.Close();
%>
