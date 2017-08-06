<!--#include file="data/conn.asp"-->
<!--#include file="data/session.asp"-->
<%
'//////////////////////////////
'By   : cm.ivan
'Time : 0:51 2010-10-2
'For  : Inventory Management
'Email: cm.ivan@163.com
'//////////////////////////////


dim db_table
    db_table="article"

dim Type_id,keyword,rsSql,ok_id,del_id
    Type_id =request.QueryString("Type_id")
    keyword =request("keyword")
    keyword =replace(keyword,"'","")
    'ok_id   =request.QueryString("ok_id")
    del_id  =request.QueryString("del_id")

'///出货的、、、、、
'if ok_id<>"" and isnumeric(ok_id) then conn.execute("update "&db_table&" set ok=1,Shipped_Time='"&now()&"' where id="&ok_id)
if del_id<>"" and isnumeric(del_id) then conn.execute("delete * from "&db_table&" where id="&del_id)

    if keyword<>"" then
       keywordStr1=" and title like '%"&keyword&"%'"
       keywordStr2=" where title like '%"&keyword&"%'"
    end if

    if Type_id<>"" and isnumeric(Type_id) then
       rsSql="select * from "&db_table&" where Type_id="&Type_id&keywordStr1&" and ok=0 order by id desc"
    else
       if keywordStr2<>"" then
       rsSql="select * from "&db_table&keywordStr2&" and ok=0 order by id desc"
       else
       rsSql="select * from "&db_table&" where ok=0 order by id desc"
       end if
    end if

set rs=server.createobject("adodb.recordset")
	rs.open rsSql,conn,1,1 
    if rs.eof then
%>
<div align="center" class="page_none">暂无相关信息!</div>
<%
    else


	rs.PageSize =10 '每页记录条数
	iCount   =rs.RecordCount '记录总数
	iPageSize=rs.PageSize
	maxpage  =rs.PageCount 
	page     =request("page")
	if Not IsNumeric(page) or page="" then
	   page=1
	else
	   page=cint(page)
	end if
	if page<1 then
	   page=1
	elseif  page>maxpage then
	   page=maxpage
	end if
	rs.AbsolutePage=Page
	if page=maxpage then
	   x=iCount-(maxpage-1)*iPageSize
	else
	   x=iPageSize
	end if
%>
<div id="div_box" onmouseover="mouseOverDiv('div_box');" onmouseout="mouseOutDiv('div_box');"><div id="ajax_div_box">loading</div></div>
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="list_box">
  <tr>
    <td bgcolor="#BBBBBB" class="list_title_bg">产品编码</td>
    <td width="50" align="center" bgcolor="#BBBBBB">点数</td>
    <td width="150" align="center" bgcolor="#BBBBBB">入库日期</td>
    <td width="85" align="center" bgcolor="#BBBBBB">操作</td>
  </tr>
<%For i=1 To x%>
<tr <%=onTable%> class="list_title" ondblclick="showNote('div_box',<%=rs("id")%>);" onmouseover="mouseOutDiv('div_box');" title="双击，可以查看详情!">
    <td height="21" class="list_title_bg"><%=key(rs("title"),keyword)%></td>
    <td align="center"><%=rs("num")%></td>
    <td align="center"><span class="info"><%=rs("add_data")%></span></td>
<td align="center">
<input type="button" name="Submit2" value=" × " onClick='if(confirm("确定要删除？")){gAjax("<%=db_table%>_list.asp?Type_id=<%=rs("Type_id")%>&page=<%=page%>&del_id=","list",<%=rs("id")%>);}else{return false;}'/><input type="button" name="Submit" value=" 查看 " onClick='gAjax("admin_article_edit.asp?id=","list","<%=rs("id")%>");' /></td>
</tr>
<%
   rs.movenext
next
%>
</table>
<%
'分页子程序
'生成上一页下一页链接
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
   
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next
%>
<!--#include file="data/article_paging.asp"-->
<%
    end if
set rs=nothing
%>