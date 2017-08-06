<%
'//////////////////////////////
'By   : cm.ivan
'Time : 0:51 2010-10-2
'For  : Inventory Management
'Email: cm.ivan@163.com
'//////////////////////////////
dim login
    login=request.QueryString("login")
 if login="out" then
    session.Abandon()
    call backPage("成功退出系统!","login.asp",0)
 end if
 
if session("cm_ivan_sys_user")="" then
   call backPage("登陆超时,请重新登录!","login.asp",0)
end if
%>