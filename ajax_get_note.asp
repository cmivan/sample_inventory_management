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
dim id,rsSql
    id=request.QueryString("id")
    if id<>"" and isnumeric(id) then
       rsSql="select top 1 * from "&db_table&" where id="&id&" order by id desc"
    else
       rsSql="select top 1 * from "&db_table&" order by id desc"
    end if
set rs=conn.execute(rsSql)
    if not rs.eof then
%>
<table width="350" border="0" cellpadding="0" cellspacing="3">
  <tr>
    <td width="77" align="right" valign="top">产品编码：</td>
    <td width="267" valign="top"><%=rs("title")%></td>
  </tr>
  <tr>
    <td width="77" align="right" valign="top">点数：</td>
    <td width="267" valign="top"><%=rs("num")%></td>
  </tr>
  <tr>
    <td align="right" valign="top">单价：</td>
    <td valign="top"><%=rs("Price")%></td>
  </tr>
  <tr>
    <td align="right" valign="top">供应商：</td>
    <td valign="top"><%=rs("suppliers_id")%></td>
  </tr>
  <tr>
    <td align="right" valign="top">接收单位：</td>
    <td valign="top"><%=rs("R_unit")%></td>
  </tr>
  <tr>
    <td align="right" valign="top">入库日期：</td>
    <td valign="top"><%=rs("Add_data")%></td>
  </tr>
  <tr>
    <td align="right" valign="top">出库日期：</td>
    <td valign="top"><%=rs("Shipped_Time")%></td>
  </tr>
  <tr>
    <td align="right" valign="top">备注：</td>
    <td valign="top"><%=rs("note")%></td>
  </tr>
</table>
<%
    else
%>
暂无相关信息!
<%
    end if
set rs=nothing
%>