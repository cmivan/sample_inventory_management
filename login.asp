<!--#include file="data/conn.asp"-->
<!--#include file="data/md5.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Inventory Systems v1.0 - all by:Cm.ivan contact:cm.ivan@163.com</title>
<script language="javascript" src="js/ajax.js"></script>
<link href="css/style.css" rel="stylesheet" type="text/css" />
</head>
<body style="padding:10px;">
<!--
//////////////////////////////
By   : cm.ivan
Time : 0:51 2010-10-2
For  : Inventory Management
Email: cm.ivan@163.com
//////////////////////////////
-->

<%
dim errStr,password
password=request.Form("password")
if password<>"" then
   password_1=replace(password,"'","")
if password_1<>password then errStr="请勿输入非法字符~"
if errStr="" then
   password_1=md5(password)
   set rs=conn.execute("select * from admin where password='"&password_1&"'")
       if rs.eof then
          errStr="密码有误,登录失败~"
       else
          session("cm_ivan_sys_user")=password_1
          call backPage("登陆成功 Time:"&now()&"!","index.asp",0)
       end if
   set rs=nothing
end if

end if
%>
<div class="mainWidth" id="ajax_main"><br />
  <table border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
    <tr>
      <td><table border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#CCCCCC">
<form action="" method="post">
<tr>
<td align="right">&nbsp;&nbsp;密码：</td>
<td width="116"><input name="password" type="password" id="password" /></td>
<td><input type="submit" name="Submit" value="登录(Go)" /></td>
</tr>
<%if errStr<>"" then%>
<tr><td align="center" style="color:#999999;">&nbsp;</td>
  <td colspan="2" align="left" style="color:#999999;"><%=errStr%></td>
  </tr>
<%end if%>
</form>
      </table></td>
    </tr>
  </table>
</div>
</body>
</html>

