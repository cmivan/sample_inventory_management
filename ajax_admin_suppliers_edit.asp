<!--#include file="data/conn.asp"-->
<!--#include file="data/session.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="js/ajax.js"></script>
<script>
//进入编辑状态
function getEdit(id){window.location.href='?id='+id;}
</script>
<body style="background-image:none; background-color:#cccccc;">
<!--
//////////////////////////////
By   : cm.ivan
Time : 0:51 2010-10-2
For  : Inventory Management
Email: cm.ivan@163.com
//////////////////////////////
-->

<%
'################| 处理删除排序问题 |################
   add_id_b=request.QueryString("add_id_b")
if add_id_b<>"" and isnumeric(add_id_b) then session("type_id")=add_id_b


'################| 处理删除排序问题 |################
   class_b_del=request.QueryString("class_b_del")
if class_b_del<>"" and isnumeric(class_b_del) then 
'-===================| 删除大类 |=====================-
  sql="delete from suppliers where id="&class_b_del   
  conn.execute(sql)
  '删除相应文章
  del_Info="成功删除"
  call backPage(del_info,"?",0)
end if



'################| 处理分类排序问题 |################
  order     =request.QueryString("order")
  id_b=request.QueryString("id_b")
  if order<>"" and id_b<>"" and isnumeric(id_b) then
	 order_sql="select * from suppliers where id="&id_b
	 set row_now=conn.execute(order_sql)
     if not row_now.eof then
        row_now_order_id=row_now("order_id")
		row_now_id      =row_now("id")
	 else
		call backPage("参数有误!","?",0)
	 end if
  end if


  if order="up" then
    '---------------------------------------------
set row_up=server.CreateObject("adodb.recordset")
    order_sql_up="select * from suppliers where order_id<"&row_now_order_id&orderStr&" order by order_id desc"
    row_up.open order_sql_up,conn,1,1
	 if not row_up.eof then
        row_up_order_id=row_up("order_id")
		conn.execute("update suppliers set order_id="&row_now_order_id&" where order_id="&row_up_order_id&orderStr)
		conn.execute("update suppliers set order_id="&row_up_order_id&" where id="&row_now_id&orderStr)
     else
		call backPage("排序已到上限!","?",0)
	 end if
	row_up.close
set row_up=nothing

  elseif order="down" then
set row_down=server.CreateObject("adodb.recordset")
    order_sql_down="select * from suppliers where order_id>"&row_now_order_id&orderStr&" order by order_id asc"
    row_down.open order_sql_down,conn,1,1
	 if not row_down.eof then
        row_down_order_id=row_down("order_id")
		row_down_query ="update suppliers set order_id="&row_now_order_id&" where order_id="&row_down_order_id&orderStr
		row_now_query  ="update suppliers set order_id="&row_down_order_id&" where id="&row_now_id&orderStr
		conn.execute(row_down_query)
		conn.execute(row_now_query)
	 else
		call backPage("排序已到下限!","?",0)
	 end if
	row_down.close
set row_down=nothing
	 
  end if

  
'################| 处理添加分类问题 |################
   id=request("id")
   edit=request("edit")
   
   title   = request.form("title")
   order_id= request.form("order_id")
   pic     = request.form("pic")
   type_id = request.form("type_id")
   
IF order_id="" or isNumeric(order_id)=false then order_id=0
IF type_id="" or isNumeric(type_id)=false then type_id=0

if edit<>"" then
   IF title=""  then 
      response.Write("<script language=javascript>alert('分类名称不能为空!');history.go(-1)</script>") 
      response.end()
   End IF

set rs=server.createobject("adodb.recordset")
 if edit="update" and id<>"" and isnumeric(id) then
    editStr="修改"
    sql="select * from suppliers where id="&id 
    rs.open sql,conn,1,3
 else
    editStr="添加"
    sql="select * from suppliers"
    rs.open sql,conn,1,3
    rs.addnew
 end if

'######### 写入数据 #############
    rs("title")   = title
    rs("order_id")= order_id
	rs("pic")     = pic
	
 if edit="add" then rs("type_id") =type_id
    session("type_id")=type_id

    rs.update
    rs.close
set rs=nothing
Response.Write "<script>window.location.href='?';</script>" 
end if 
%>




<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <TR><td>

<table width="100%" border="0" cellpadding="0" cellspacing="1" class="list_box">
<form action="?edit=add" method="post" name="mclass_type_add" id="mclass_type_add">
          <tr>
            <td align="center" bgcolor="#BBBBBB" class="forumRaw" width="49">编号</td>
            <td width="380" align="left" bgcolor="#BBBBBB" class="forumRaw">&nbsp;
            <input name="title" type="text" id="title" style="width:120px" />
            <input name="order_id" type="text" id="order_id" style="width:40px; text-align:center;" value="0" size="5" />
            <input name="submit" type="submit" class="input_but" id="submit" value="添加供应商" />
            </td>
            <td colspan="2" align="center" bgcolor="#BBBBBB" class="forumRaw" width="85">排序</td>
            <td align="center" bgcolor="#BBBBBB" class="forumRaw">操作</td>
          </tr>
</form>
<%	
	set row_b=server.createobject("adodb.recordset") 
	    row_b_sql="select * from suppliers where type_id=0 order by order_id asc" 
	    row_b.open row_b_sql,conn,1,3
    do while not row_b.eof
	'#### 重写排序 ####
	r_b=r_b+1
	row_b("order_id")=r_b
	row_b.update()
%>
        <form name="article_type<%=row_b("id")%>" method="post" action="?edit=update&id=<%=row_b("id")%>">
         
<%if cstr(request("id"))=cstr(row_b("id")) then%>
<tr <%=onTable%> class="list_title" onDblClick="submit();" title="双击可完成编辑!">
<td align="center"><%=row_b("id")%><a href="article_manage.asp?Type_id=<%=row_b("id")%>">
              <input name="id" type="hidden" id="id" value="<%=row_b("id")%>" />
            </a></td>
           
            <td width="380" align="left" class="list_title_bg"><input name="title" type="text" class="input2" id="title" value="<%=row_b("title")%>" /></td>
		    <td width="60" align="center"><input style="background-color:#F7F7EE; text-align:center" name="order_id" type="text" value="<%=row_b("order_id")%>" size="4"></td>
            <td width="50" align="center">
<a href="?order=up&id_b=<%=row_b("id")%>"><img src="images/up_ico.gif" width="12" height="12" border="0"></a>

<a href="?order=down&id_b=<%=row_b("id")%>"><img src="images/down_ico.gif" width="12" height="12" border="0"></a></td>
            <td align="center"><input name="update2" type="submit" class="input_but" value="保存" id="input_but" /></td>
</tr>
<%else%>
<tr <%=onTable%> class="list_title" onDblClick="getEdit(<%=row_b("id")%>);" title="双击可编辑该分类!">
<td width="49" align="center"><%=row_b("id")%><input name="id" type="hidden" id="id" value="<%=row_b("id")%>" /></td>
           <td width="380" align="left" class="list_title_bg">&nbsp;<%=row_b("title")%></td>
		    <td width="60" align="center" style="font-weight:bold"><%=row_b("order_id")%></td>
            <td width="50" align="center">
<a href="?order=up&id_b=<%=row_b("id")%>"><img src="images/up_ico.gif" width="12" height="12" border="0"></a>

<a href="?order=down&id_b=<%=row_b("id")%>"><img src="images/down_ico.gif" width="12" height="12" border="0"></a>			</td>
<td align="center">
<input name="delete" type="button" class="input_but" onClick="if(confirm('确定要删除此导航栏吗？删除后不可恢复!')){window.location.href='?class_b_del=<%=row_b("id")%>';}else{return false;}" value="删除"/></td>
</tr>
<%end if%> 
        </form>
<%
	row_b.movenext
	loop
	row_b.close
set row_b=nothing
%>
  </table>

</td></TR></TABLE>
</body>