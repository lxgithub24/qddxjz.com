<%
var menuitems=[{text:"  青大兼职首页",url:"index.asp"},{text:"注册",url:"register.asp"},{text:"登陆",url:"load.asp"},{text:"个人信息修改",url:"modefy_user.asp"},{text:"退出",url:"load.asp?action=logout"}];
var currentPage=String(Request.ServerVariables("URL"));
var welcome="欢迎您";
if(String(Session("MM_Username")) =="undefined"){
welcome+="，游客！";
}
if(String(Session("MM_Username"))!="undefined"){
welcome+="！"+String(Session("MM_Username"));
}
%>
<link href="style.css" rel="stylesheet" type="text/css">
<div class="menu">
<table width="100%">
	<tr>
		<td align="left">   <%
var count=menuitems.length;
Response.Write(welcome+"&nbsp;");
for (var i=0;i<count;i++){
if(currentPage.indexOf(menuitems[i].url)!=-1){
Response.Write("<span class='currentPage'>"+menuitems[i].text+"</sapn>")}
else{
Response.Write("<a href='"+menuitems[i].url+"'>"+menuitems[i].text+"</a>");}
if(menuitems[i].text.indexOf("退出")==-1){
Response.Write("|");
}
}
%> 
		</td>
		<td align="right">
			<a href="happy.asp">happy一下</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
</table>

</div>
