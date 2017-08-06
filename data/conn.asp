<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
'//////////////////////////////
'By   : cm.ivan
'Time : 0:51 2010-10-2
'For  : Inventory Management
'Email: cm.ivan@163.com
'//////////////////////////////

session.Timeout=300
'=++++++++++++++++++++++++++++++++++++++++=
'on error resume next
dim conns,connstr,mdbs
'-----------  Access数据库连接字符串 -----------
    mdbs="data/KM_09_12_20.mdb"           '数据库文件目录
    Connstr="DRIVER=Microsoft Access Driver (*.mdb);DBQ="+server.mappath(mdbs)
'-----------  SQL数据库连接字符串    -----------
'	Connstr="driver={SQL Server};server=192.168.0.1,7788;database=KM_09_12_20;UID=sa;PWD=sa"
'---------------------------------------------
Set conn=Server.CreateObject("ADODB.Connection") 
    conn.Open connstr
'------------------------------------------==
 If Err Then
    Err.Clear
    Set Conns = Nothing
    Response.Write "数据库连接有误!请检查相应文件!"
    Response.End
 End If
'=++++++++++++++++++++++++++++++++++++++++=
%>


<%
'### 鼠标经过表格样式
    onTable=" onmouseover=""cursor_('on',this);"" onmouseout=""cursor_('',this);"""
function key(str,keyword)
    dim keywords
    keywords="<span style='color:#ff0000'>"&keyword&"</span>"
    key=replace(str,keyword,keywords)
end function
'***********************************
'    通用调用单条信息
'***********************************
function get_one(dbTable,gid,gfield)
     'gid未指定值的情况下，自动接收参数
	  if gid=0 then gid=request.QueryString("id")
  set get_one_rs=server.CreateObject("adodb.recordset")
      if gid="" or isnumeric(gid)=false then
	     get_one_sql="select top 1 * from "&dbTable&" order by order_id asc,id desc"
	  else
	  	 get_one_sql="select * from "&dbTable&" where id="&int(gid)
	  end if
	  get_one_rs.open get_one_sql,conn,1,3
	  if not get_one_rs.eof then
	     if gfield="hit" then
		    if get_one_rs("hits")="" or get_one_rs("hits")=null or isnumeric(get_one_rs("hits"))=false then get_one_rs("hits")=0
	        get_one_rs("hits")=get_one_rs("hits")+1        '累计点击
			get_one_rs.update
			else
			get_one=get_one_rs(gfield)  '读取指定字段
		end if
	  end if
	  get_one_rs.close  
  set get_one_rs=nothing
end function
'***********************************
'    过滤HTML字符
'***********************************
Function RemoveHTML(strHTML) 
   Dim objRegExp, Match, Matches 
   Set objRegExp = New Regexp 
       objRegExp.IgnoreCase = True 
       objRegExp.Global = True 
       '取闭合的<> 
       objRegExp.Pattern = "<.+?>" 
       '进行匹配 
   Set Matches = objRegExp.Execute(strHTML) 
       '遍历匹配集合，并替换掉匹配的项目 
   For Each Match in Matches 
       strHtml=Replace(strHTML,Match.Value,"") 
   Next
       RemoveHTML=strHTML 
   Set objRegExp = Nothing 
End Function
'***********************************
'    截取指定长度字符
'***********************************
function cutStr(str,lens)
dim i
i=1
y=0
txt=RemoveHTML(trim(str))
for i=1 to len(txt)
j=mid(txt,i,1)
if ascw(j)>=0 and ascw(j)<=127 then '汉字外的其他符号,如:!@#,数字,大小写英文字母
y=y+0.5
else '汉字
y=y+1
end if
if -int(-y) >= lens then '截取长度
txt = left(txt,i)&"..."
exit for
end if
next
cutStr=txt
end function



'× --------------------------------------
'× ----------  返回提示信息 ---------------
'× --------------------------------------
  function backPage(backStr,backUrl,backType)
    back =""
	back =back&"<meta http-equiv=Content-Type content=text/html; charset=utf-8 />"
	back =back&"<link href='css/style.css' rel='stylesheet' type='text/css' />"
	
	if backType=0 then
	    'meta自动跳转到指定页面
        back =back&"<meta http-equiv=refresh content=1;url="&backUrl&">"
		back =back&"<body style=""overflow:hidden;"">"
	    back =back&"<br><TABLE border=0 align=center cellpadding=0 cellspacing=10 bgcolor=#FFFFFF><tr><td width=90% class=forumRow>"
		back =back&"<table width=100% border=0 align=center cellpadding=3 cellspacing=1 class=forMy><tr><td class=forumRow align=center>"
		back =back&backStr
		back =back&"</tr></table></td></tr></table>"
	elseif backType=1 then
	    'js弹出提示，返回指定页面
	    back =back&"<script language='javascript'>alert('"&backStr&"');window.location.href='"&backUrl&"';</script>"
	elseif backType=2 then
	    'js弹出提示，返回上一级
	    back =back&"<script language='javascript'>alert('"&backStr&"');history.back(1);</script>"
	elseif backType=3 then
	    'js弹出提示，返回指定页
	    back =back&"<script language='javascript'>window.location.href='"&backUrl&"';</script>"
	end if
	response.Write(back)
	response.End()
 end function
%>





