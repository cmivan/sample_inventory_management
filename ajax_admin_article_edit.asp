<!--#include file="data/conn.asp"-->
<!--#include file="data/session.asp"-->
<%
'//////////////////////////////
'By   : cm.ivan
'Time : 0:51 2010-10-2
'For  : Inventory Management
'Email: cm.ivan@163.com
'//////////////////////////////


'### 配置信息 ###
 db_table    ="article"   '表段

'///////////  处理提交数据部分 ////////// 
   edit_id   =request("id")
if edit_id="" or isnumeric(edit_id)=false then
   editStr="添加"
   else
   editStr="修改"
end if

   
if request.Form("edit")="ok" then
'——————接收数据
    title       =request.Form("title")
    note        =request.Form("note")
    Type_id     =request.Form("Type_id")
    num         =request.Form("num")
    Price       =request.Form("Price")
    cm_number   =request.Form("cm_number")
    suppliers_id=request.Form("suppliers_id")
    R_unit      =request.Form("R_unit")
    Add_data    =request.Form("Add_data")
    Shipped_Time=request.Form("Shipped_Time")

if Type_id="" or isnumeric(Type_id)=false then
   response.Write("<script>alert('参数有误,请正确填写!');</script>")
   response.End()
end if 
if num="" or isnumeric(num)=false then
   response.Write("<script>alert('请正确输入点数!');</script>")
   response.End()
end if 
if cm_number="" then
   response.Write("<script>alert('请输入编号!');</script>")
   response.End()
end if
if suppliers_id="" or isnumeric(suppliers_id)=false then
   response.Write("<script>alert('请输入供应商 !');</script>")
   response.End()
end if
if Add_data="" or isdate(Add_data)=false then
   response.Write("<script>alert('请正确输入入库日期!');</script>")
   response.End()
end if
if Shipped_Time<>"" and isdate(Shipped_Time)=false then
   response.Write("<script>alert('请正确输入出库日期!');</script>")
   response.End()
end if




'///////////  写入数据部分 //////////
   set rs=server.createobject("adodb.recordset") 
    if edit_id="" or isnumeric(edit_id)=false then
       exec="select * from "&db_table                       '判断，添加数据
	   rs.open exec,conn,1,3
       rs.addnew
	else
	   exec="select * from "&db_table&" where id="&edit_id  '判断，修改数据
       rs.open exec,conn,1,3
	end if

	if edit_id<>"" and isnumeric(edit_id) and rs.eof then
	   response.Write("写入数据失败!")
	else

    dim cm_title
        cm_title=get_one(db_table&"_type",Type_id,"title")
	rs("title")    =cm_title&cm_number
	rs("note")     =note
	rs("Type_id")  =Type_id
	rs("num")      =Num
	rs("Price")    =Price
	rs("cm_number")=cm_number

    rs("suppliers_id")=suppliers_id
	rs("R_unit")      =R_unit
	rs("Add_data")    =Add_data

    if Shipped_Time<>"" then
       rs("Shipped_Time")=Shipped_Time
       rs("ok")=1
    else
       'rs("Shipped_Time")=""
       rs("ok")=0
    end if

    response.Write("<script>alert('"&editStr&"操作成功!');</script>")
	end if
	rs.update
	rs.close
set rs=nothing

end if



'#################  读取数据部分 #################

   id=request.QueryString("id") 
if id<>"" and isnumeric(id) then
  '/// 当前状态:编辑
   edit_stat="edit"
   set rs=server.createobject("adodb.recordset") 
       exec="select * from "&db_table&" where id="&id
       rs.open exec,conn,1,1 
	   if not rs.eof then
	      title    =rs("title")
	      note     =rs("note")
		  Type_id  =rs("Type_id")
		  '记录分类,用于分类下拉
		  session("type_id")=Type_id
	      num   =rs("num")
	      Price =rs("Price")
	      cm_number =rs("cm_number")
          suppliers_id=rs("suppliers_id")
	      R_unit      =rs("R_unit")
	      Add_data    =rs("Add_data")
          Shipped_Time=rs("Shipped_Time")
          ok  =rs("ok")
	   end if
	   rs.close
   set rs=nothing
   
else
  '/// 当前状态:添加
   edit_stat="add"
end if

if Add_data="" or isdate(Add_data)=false then Add_data=now()
%>

<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
  <TR><td class="forumRow">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forMy" style="background-color:#FFFFFF">
<form name="article_update" method="post" action="ajax_admin_article_edit.asp" target="admin_edit_box">		  
<tr>
<td width="22%" align="right" class="forumRow">型号：</td>
<td width="78%" class="forumRow">
<select name="type_id">
<%
	dim rsc
	'///######### 大类
Set rsc=Conn.Execute("select * from "&db_table&"_type where type_id=0 order by order_id asc")
	while not rsc.eof
	selected=""
	if cstr(session("type_id"))=cstr(rsc("id")) then selected="selected"
	   response.Write("<option value='" & rsc("id") & "' "&selected&">" & rsc("title") & "</option>")		   
    rsc.movenext
    wend
	rsc.close
set rsc=nothing
%>
</select>


&nbsp;&nbsp;
<input name="num" type="text" id="num" value="<%=num%>" size="10" style="width:50px;" />
(点数)</td>
</tr>

<tr>
<td align="right" valign="top" class="forumRow">编号：</td>
<td class="forumRow"><input name="cm_number" type="text" id="cm_number" value="<%=cm_number%>" /></td>
</tr>
<tr>
<td align="right" valign="top" class="forumRow">单价：</td>
<td class="forumRow"><input name="Price" type="text" id="Price" value="<%=Price%>" />
(元)</td>
</tr>
<tr>
<td align="right" valign="top" class="forumRow">供应商：</td>
<td class="forumRow">
<select name="suppliers_id">
<%
Set rsc=Conn.Execute("select * from suppliers where type_id=0 order by order_id asc")
	while not rsc.eof
	selected=""
	if cstr(session("suppliers_id"))=cstr(rsc("id")) then selected="selected"
	   response.Write("<option value='" & rsc("id") & "' "&selected&">" & rsc("title") & "</option>")		   
    rsc.movenext
    wend
	rsc.close
set rsc=nothing
%>
</select>
</td>
</tr>

<tr>
<td align="right" valign="top" class="forumRow">接收单位：</td>
<td class="forumRow"><input name="R_unit" type="text" id="R_unit" value="<%=R_unit%>" /></td>
</tr>


<tr>
<td align="right" valign="top" class="forumRow">入库日期：</td>
<td class="forumRow"><input name="Add_data" type="text" id="Add_data" value="<%=Add_data%>" /></td>
</tr>
<tr>
<td align="right" valign="top" class="forumRow">出库日期：</td>
<td class="forumRow"><input name="Shipped_Time" type="text" id="Shipped_Time" value="<%if ok=1 then response.Write(Shipped_Time) end if%>" onchange="if(this.value!=''){ok.checked='checked'}else{ok.checked=''}" />
  &nbsp;&nbsp;出库
  <input name="ok" type="checkbox" id="ok" value="1" disabled="disabled"<%if ok=1 then%> checked="checked"<%end if%> /></td>
</tr>
<tr>
<td align="right" valign="top" class="forumRow">备注：</td>
<td class="forumRow"><textarea name="note" cols="45" rows="3" id="note"><%=note%></textarea></td>
</tr>
		  
          <tr>
            <td align="right" valign="top" class="forumRow"><p class="submit">
                <input name="id" type="hidden" value="<%=id%>" />
				<input name="edit" type="hidden" value="ok" />
            </p></td>
            <td class="forumRow"><span class="submit">
              <input name="submit" type="submit" value="保存<%=editStr%>" />
            </span></td>
          </tr>
        </form>
      </table>
</td></TR>
</TABLE>
<iframe src="" name="admin_edit_box" scrolling="0" frameborder="0" height="1" width="1"></iframe>