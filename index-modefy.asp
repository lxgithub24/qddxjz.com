<%
var menuitems=[{text:"  ����ְ��ҳ",url:"index.asp"},{text:"ע��",url:"register.asp"},{text:"��½",url:"load.asp"},{text:"������Ϣ�޸�",url:"modefy_user.asp"},{text:"�˳�",url:"load.asp?action=logout"}];
var currentPage=String(Request.ServerVariables("URL"));
var welcome="��ӭ��";
if(String(Session("MM_Username")) =="undefined"){
welcome+="���οͣ�";
}
if(String(Session("MM_Username"))!="undefined"){
welcome+="��"+String(Session("MM_Username"));
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
if(menuitems[i].text.indexOf("�˳�")==-1){
Response.Write("|");
}
}
%> 
		</td>
		<td align="right">
			<a href="happy.asp">happyһ��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
</table>

</div>