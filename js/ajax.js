// by  : cm.ivan
// blog: http://blog.163.com/cm.vian
// time: 2010-01-01

//定义鼠标移过样式
var mcolor;
//用于记忆原来的颜色
function cursor_(ctype,objs){
    if (ctype=="on"){
	   mcolor=objs.style.backgroundColor;
	   objs.style.backgroundColor="#eeeeee";
	}else{
       objs.style.backgroundColor=mcolor;
	}
}


//获取鼠标当前位置
function mousePos(){
return {x:document.body.scrollLeft+event.clientX,y:event.clientY};}

var show_on;  //当前div id值
     show_on="";
var Timename;
function showNote(show_id,id){
var pos=mousePos();
if (show_on!=""){
document.getElementById(show_on).style.display='none';}
collist=document.getElementById(show_id);
collist.style.left=pos.x-5+'px';
collist.style.top =pos.y-5+'px';
collist.style.display='block';
show_on=show_id;
gAjax("get_note.asp?id=",show_id,id);
mouseOutDiv(show_id);
}

function mouseOutDiv(id){
Timename=setTimeout("document.getElementById('"+id+"').style.display='none';",700);
}
function mouseOverDiv(id){
clearTimeout(Timename);
document.getElementById(id).style.display='block';
}





// JavaScript Document
// 初始化一个xmlhttp对象
function InitAjax()
{
　var ajax=false; 
　try { 
　　ajax = new ActiveXObject("Msxml2.XMLHTTP"); 
　} catch (e) { 
　　try { 
　　　ajax = new ActiveXObject("Microsoft.XMLHTTP"); 
　　} catch (E) { 
　　　ajax = false; 
　　} 
　}
　if (!ajax && typeof XMLHttpRequest!='undefined') { 
　　ajax = new XMLHttpRequest(); 
　} 
　return ajax;
}
function gAjax(toUrl,toStr,toID)
{
  var thisDate=new Date();
  var rndDate =thisDate.getHours() +"_"+ thisDate.getMinutes() +"_"+ thisDate.getSeconds();
  //var rndDate =thisDate.getHours() + thisDate.getMinutes();
　//如果没有把参数toID传进来
　if (typeof(toID) == 'undefined')
　{return false;}
　//需要进行Ajax的Url地址 如:ajax_show.asp?id=1
　var Url = "ajax_" + toUrl + toID + "&rndDate=" + rndDate;
      Url = encodeURI(Url);
　//获取新闻显示层的位置
　var show = document.getElementById("ajax_" + toStr);
  show.innerHTML="<div class='loadingBox'><img src='images/loading.gif'></div>";
　//实例化Ajax对象
　var ajax = InitAjax();
　//使用Get方式进行请求
　ajax.open("GET", Url, true); 
　//获取执行状态
　ajax.onreadystatechange = function() { 
  if(ajax.readyState == 4 && ajax.status == 200){
show.innerHTML = ajax.responseText;}
  }
　ajax.send(null);  //发送空
}