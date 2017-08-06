<!--#include file="data/conn.asp"-->
<!--#include file="data/md5.asp"-->
<!--#include file="data/session.asp"-->
<%
'//////////////////////////////
'By   : cm.ivan
'Time : 0:51 2010-10-2
'For  : Inventory Management
'Email: cm.ivan@163.com
'//////////////////////////////


dim errStr,password
password=request("password")
password_new=request("password_new")
if password<>"" then
   password_1     =replace(password,"'","")
   password_new_1 =replace(password_new,"'","")
if password_1<>password         then errStr="请勿输入非法字符~"
if password_new_1<>password_new then errStr="请勿输入非法字符~"

password_1    =md5(password)
password_new_1=md5(password_new)
if password_new="" and errStr="" then errStr="请填写新密码~"
if password_new_1=session("cm_ivan_sys_user") and errStr="" then errStr="你填写的新密码和旧密码一样!"

if errStr="" then
   set rs=conn.execute("select * from admin where password='"&password_1&"'")
       if not rs.eof then
          conn.execute("update admin set password='"&password_new_1&"' where id="&rs("id"))
          if err=0 then
             session("cm_ivan_sys_user")=password_new_1
             errStr="密码更改成功!"
          else
             errStr="密码更改失败!"
          end if
       else
          errStr="原密码不正确,密码更改失败!"
       end if
   set rs=nothing
end if

end if
%>
<div style="padding-top:80px; padding-bottom:80px;">
<table border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
<tr><td><table border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#CCCCCC">
<form name="epass" action="" method="post">
<tr>
<td align="right">&nbsp;&nbsp;原密码：</td>
<td width="116"><input name="password" type="password" id="password" /></td>
<td>&nbsp;</td>
</tr>
<tr>
<td align="right">&nbsp;&nbsp;新密码：</td>
<td width="116"><input name="password_new" type="password" id="password_new" /></td>
<td><input type="button" name="Submit" value="保存(Save)" onclick='gAjax("edit_password.asp?password="+document.epass.password.value+"&password_new="+document.epass.password_new.value,"list","");' /></td>
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