<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>网页特效---反应速度测试!</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8">

<SCRIPT language=JavaScript type=text/javascript>

isNS4 = (document.layers) ? true : false;
isIE4 = (document.all  &&  !document.getElementById) ? true : false;
isIE5 = (document.all  &&  document.getElementById) ? true : false;
isNS6 = (!document.all  &&  document.getElementById) ? true : false;

var curX, curY, curX2, curY2, boxX, boxY, moving=0, touch=0;
var gametime=0, started=0, speed;
var starttime, endtime, finaltime=0; //pass finaltime to popup window to ask for initials
var enemyxdir = new Array(1,1,1,1);
var enemyydir = new Array(1,1,1,1);

if (isNS4 || isNS6){
document.captureEvents(Event.MOUSEUP|Event.MOUSEDOWN|Event.MOUSEMOVE);
}
document.onmousedown = start;
document.onmousemove = checkLocation;
document.onmouseup = stop;

function startclock() {var today = new Date(); starttime = today.getTime();}
function endclock() {var today = new Date(); endtime = today.getTime();}
function calctime() {var time = (endtime - starttime - 0)/1000; return time;}

function giveposX(divname) {
 if (isNS4) var posLeft = document.layers[divname].left;
 else if (isIE4 || isIE5) var posLeft = document.all(divname).style.pixelLeft;
 else if (isNS6) var posLeft = parseInt(document.getElementById(divname).style.left + "");
 return posLeft;
}

function giveposY(divname) {
 if (isNS4) var posTop = document.layers[divname].top;
 else if (isIE4 || isIE5) var posTop = document.all(divname).style.pixelTop;
 else if (isNS6) var posTop = parseInt(document.getElementById(divname).style.top + "");
 return posTop;
}

function setposX(divname, xpos) {
 if (isNS4) document.layers[divname].left = xpos;
 else if (isIE4 || isIE5) document.all(divname).style.pixelLeft = xpos;
 else if (isNS6) document.getElementById(divname).style.left = xpos;
}

function setposY(divname, ypos) {
 if (isNS4) document.layers[divname].top = ypos;
 else if (isIE4 || isIE5) document.all(divname).style.pixelTop = ypos;
 else if (isNS6) document.getElementById(divname).style.top = ypos;
}

function givesize(divname, dimension) {
 var divsize = 0;
  if (dimension == 'y') {
   if (isNS4) divsize = document.layers[divname].clip.height;
   else if (isIE4 || isIE5) divsize = document.all(divname).style.pixelHeight;
   else if (isNS6) divsize = parseInt(document.getElementById(divname).style.height + "");
  }
  else if (dimension == 'x') {
   if (isNS4) divsize = document.layers[divname].clip.width;
   else if (isIE4 || isIE5) divsize = document.all(divname).style.pixelWidth;
   else if (isNS6) divsize = parseInt(document.getElementById(divname).style.width + "");
  }

 return divsize;
}


// check to see if 'box' is touching 'enemy1'
function checktouching(num) {

 var enemy = "enemy" + num + ""
 var difX = giveposX('box') - giveposX(enemy) - 0; // -0 converts to integer
 var difY = giveposY('box') - giveposY(enemy) - 0;

 // set touch = 1 if it is touching an enemy
 if (difX > (-1 * givesize('box', 'x'))  &&  difX < givesize(enemy, 'x')  &&  difY > (-1 * givesize('box', 'y'))  &&  difY < givesize(enemy, 'y')) {
  touch = 1;
 }
 else touch = 0;

}

function movenemy(num,step_x,step_y){

 var enemy = "enemy" + num + ""
 var enemyx = givesize(enemy, 'x');
 var enemyy = givesize(enemy, 'y');

 if (giveposX(enemy) >= (450 - enemyx) || giveposX(enemy) <= 0) {
  enemyxdir[num] = -1 * enemyxdir[num];
  }
 if (giveposY(enemy) >= (450 - enemyy) || giveposY(enemy) <= 0) {
  enemyydir[num] = -1 * enemyydir[num];
  }

 var newposx = giveposX(enemy) + (step_x*enemyxdir[num]) + 0;
 var newposy = giveposY(enemy) + (step_y*enemyydir[num]) + 0;

 setposX(enemy, newposx);
 setposY(enemy, newposy);

 checktouching(num + "");
 if (touch == 1) {
  stop(); reset();
  }
}

function movenemies() {

 gametime = gametime + 1

 if (gametime >= 0  &&  gametime < 100) speed = 80;
 else if (gametime >= 100  &&   gametime < 200) speed = 60;
 else if (gametime >= 200  &&   gametime < 300) speed = 40;
 else if (gametime >= 300  &&   gametime < 400) speed = 30;
 else if (gametime >= 400  &&   gametime < 500) speed = 20;
 else speed = 10;
 // window.status = "speed:  " + speed + "   gametime: " + gametime;

 movenemy(0,-10,12);
 movenemy(1,-12,-20);
 movenemy(2,15,-13);
 movenemy(3,17,11);

 setTimeout(movenemies,speed);
}

function start(e) {

 if (started == 0) { movenemies();  startclock();  started = 1; }

 curX = (isNS4 || isNS6) ? e.pageX : window.event.x ;
    curY = (isNS4 || isNS6) ? e.pageY : window.event.y ;

 curX2 = eval(curX - 40);
 curY2 = eval(curY - 40);

 boxX = eval(curX - 20);
 boxY = eval(curY - 20);

 var boxleft = giveposX('box');
 var boxtop = giveposY('box');

 if (curX > boxleft  &&  curX2 < boxleft  &&  curY > boxtop  &&  curY2 < boxtop) {

  moving = 1;
  setposX('box', boxX);
  setposY('box', boxY);

  if (isNS4 || isNS6){
  document.captureEvents(Event.MOUSEMOVE);
  }
 }
}

function stop(e){
    moving=0;
 if (isNS4 || isNS6){
 document.releaseEvents(Event.MOUSEMOVE);
 }
}

function reset(e){
    endclock();
 moving=0;
 if (isNS4 || isNS6){
  document.releaseEvents(Event.MOUSEMOVE);
  }
 if (finaltime == 0) {
  finaltime = calctime();
  window.alert('您坚持了 ' + finaltime + ' 秒…');

//  var entername = window.confirm('Enter your name?');
//   if (entername) {
//   window.open("?" + finaltime,'winwin','width=300,height=500,left=40,top=40,status=1,resizable');
//   document.location.reload();
//   }
//   else document.location.reload();
   document.location.reload();
  }
}

function checkLocation(e){

  curX = (isNS4 || isNS6) ? e.pageX : window.event.x ;
        curY = (isNS4 || isNS6) ? e.pageY : window.event.y ;

  boxX = eval(curX - 20);
  boxY = eval(curY - 20);

 checktouching('1');

 if (moving == 1  &&  touch == 0){

   setposX('box',boxX);
   setposY('box',boxY);

   if (curY > 69  &&  curX > 69  &&  curY < 381  &&  curX < 381) return false;
   else stop(); reset();
 }

 else if (touch == 1){
 stop(); reset();
 }

}

</SCRIPT>

<META content="MSHTML 6.00.2900.5512" name=GENERATOR>
<style type="text/css">
<!--
.style1 {color: #ECE9D8}
.style10 {color: #FF6600}
.style4 {font-size: 12;
 color: #333333;
}
.style9 {color: #669966; font-size: 12px; }
.style2 {font-weight: bold}
.style3 {font-size: 14px}
.style5 {font-size: 16px}
-->
</style>
</HEAD>
<BODY text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<DIV id=box 
style="LEFT: 205px; WIDTH: 40px; POSITION: absolute; TOP: 205px; HEIGHT: 40px; BACKGROUND-COLOR: #990000; layer-background-color: #990000">
<TABLE height=40 width=40>
  <TBODY>
  <TR>
    <TD> </TD></TD></TR></TBODY></TABLE></DIV>
<DIV id=enemy0 
style="LEFT: 270px; WIDTH: 60px; POSITION: absolute; TOP: 60px; HEIGHT: 50px; BACKGROUND-COLOR: #000099; layer-background-color: #000099">
<TABLE height=50 width=60>
  <TBODY>
  <TR>
    <TD> </TD></TR></TBODY></TABLE></DIV>
<DIV id=enemy1 
style="LEFT: 300px; WIDTH: 100px; POSITION: absolute; TOP: 330px; HEIGHT: 20px; BACKGROUND-COLOR: #000099; layer-background-color: #000099">
<TABLE height=20 width=100>
  <TBODY>
  <TR>
    <TD> </TD></TR></TBODY></TABLE></DIV>
<DIV id=enemy2 
style="LEFT: 70px; WIDTH: 30px; POSITION: absolute; TOP: 320px; HEIGHT: 60px; BACKGROUND-COLOR: #000099; layer-background-color: #000099">
<TABLE height=60 width=30>
  <TBODY>
  <TR>
    <TD> </TD></TR></TBODY></TABLE></DIV>
<DIV id=enemy3 
style="LEFT: 67px; WIDTH: 60px; POSITION: absolute; TOP: 67px; HEIGHT: 60px; BACKGROUND-COLOR: #000099; layer-background-color: #000099">
<TABLE height=60 width=60>
  <TBODY>
  <TR>
<TD> </TD></TR></TBODY></TABLE></DIV>
<TABLE cellSpacing=0 cellPadding=0 border=0><!-- row 1 -->
  <TBODY>
  <TR bgcolor="#333333">
    <TD width=50 height=50 bgColor=#333333>
      <TABLE>
        <TBODY>
        <TR>
      <TD class="style1"></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
    <TD></TD></TR></TBODY></TABLE></TD></TR><!-- row 2 -->
  <TR>
    <TD width=50 bgColor=#333333 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD class="style1"></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50 bgColor=#333333>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD></TR><!-- row 3 -->
  <TR>
    <TD width=50 bgColor=#333333 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD class="style1"></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50 bgColor=#333333>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD></TR><!-- row 4 -->
  <TR>
    <TD width=50 bgColor=#333333 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD class="style1"></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50 bgColor=#333333>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD></TR><!-- row 5 -->
  <TR>
    <TD width=50 bgColor=#333333 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD class="style1"></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50 bgColor=#333333>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD></TR><!-- row 6 -->
  <TR>
    <TD width=50 bgColor=#333333 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD class="style1"></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50 bgColor=#333333>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD></TR><!-- row 7 -->
  <TR>
    <TD width=50 bgColor=#333333 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD class="style1"></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50 bgColor=#333333>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD></TR><!-- row 8 -->
  <TR>
    <TD width=50 bgColor=#333333 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD class="style1"></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
          <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50 bgColor=#333333>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD></TR><!-- row 9 -->
  <TR bgcolor="#333333">
    <TD width=50 height=50 bgColor=#333333>
      <TABLE>
        <TBODY>
        <TR>
      <TD class="style1"></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD width=50 height=50>
      <TABLE>
        <TBODY>
        <TR>
      <TD></TD></TR></TBODY></TABLE></TD>
    <TD height=50>      <table>
      <tbody>
        <tr>
          <td></td>
        </tr>
      </tbody>
    </table></TD>
  </TR></TBODY></TABLE>
<DIV id=layer1 
style="Z-INDEX: 1; LEFT: 489px; WIDTH: 285px; POSITION: absolute; TOP: 6px; HEIGHT: 279px">
  <table width="284" border="0">
    <tr>
      <td><span class="style9 style5"><strong>反应速度测试之-红色方块</strong></span></td>
    </tr>
    <tr>
      <td><hr size="1"></td>
    </tr>
  </table>
  <p class="style9 style5"><strong>游戏说明：</strong></p>
  <p>用鼠标点住红色的方块来移动它:<BR>
      <BR>
    移动红色的方块来躲避蓝色方块<BR>
    <BR>
    不能碰到黑色的墙壁<BR>
    <BR>
    如果你能坚持25秒以上，你就是天才！<BR>
    <BR>
    据说美国空军要求玩到两分钟。 <BR>
  </p>
</DIV>

<br>
<FONT 
face=verdana,arial,helvetica size=2><BR>
 </FONT>
<DIV id=layer1 
style="Z-INDEX: 1; LEFT: 489px; WIDTH: 288px; POSITION: absolute; TOP: 351px; HEIGHT: 60px">
  <p class="style9 style5"><BR>
  </p>
<p><BR>
  </p>
</DIV>
<DIV style="DISPLAY: none">
</DIV></BODY></HTML>
