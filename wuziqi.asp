<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>五子棋</title>
</head>

<body><!--#include file="index-modefy.asp"-->
<style>
*{
 margin:0;
 padding:0;
}
#main{
 width:500px;
 height:500px;
 border-top:#000 solid 1px;
 border-right:#000 solid 1px;
 background-color:#e0e0e0;
 margin:0px auto;
 position:absolute;
 z-index:1;
}
#main .div{
 width:19px;
 height:19px;
 border-bottom:#000 solid 1px;
 border-left:#000 solid 1px;
 float:left;
 display:inline;
 line-height:19px;
}
#main2{
 position:absolute;
 z-index:2;
 padding:10px;
 width:480px;
 height:480px;
}
#main2 span{
 width:20px;
 height:20px;
 display:block;
 float:left;
 cursor:pointer;
       line-height:20px;
 font-size:1px;
}
#main2 span.wb{
 background:url(http://home.blueidea.com/attachment/200811/7/172754_1226047522jMpx.gif) no-repeat;
}
#main2 span.bb{
 background:url(http://home.blueidea.com/attachment/200811/7/172754_122604752285lq.gif) no-repeat;
}
#aa{
 text-align:center;
 height:35px;
 line-height:35px;
 font-weight:700;
 font-size:18px;
 font-family:"微软雅黑"
}
</style>
<script type="text/javascript" language="javascript">
function $(o){
 return document.getElementById(o);
 }
var flag=0;
var black="黑"
var witer="白"
function played(){
 var piece=$("main2").getElementsByTagName("span");
 for(var i=0;i<piece.length;i++){
  piece[i].onclick = function(){
   showPieces(this);
   return false
  }
 }
}
function showPieces(t){
 if(t.className!="wb" && t.className!="bb"){
        if(flag==0){
 t.className="wb"
 flag=1;
 panduan(witer)
 }else{
 t.className="bb"
 flag=0;
 panduan(black)
 }
}
}
function panduan(te){
 $("aa").innerHTML="当前执棋:"+te
}
</script>
</head>
<body bgcolor="#CCCC99">
<div id="aa"></div>
<div id="main"></div>
<div id="main2"></div>
<script type="text/javascript" language="javascript">
 var _div=[]
 var _span=[]
 for(i=0;i<625;i++){
  _div[i]=document.createElement("div");
  _div[i].className="div"
  $("main").appendChild(_div[i])
 }
 for(j=0;j<576;j++){
  _span[j]=document.createElement("span")
  $("main2").appendChild(_span[j])
 }
 panduan(black)
 played()
</script>


</body>
</html>
