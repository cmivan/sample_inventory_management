<!--#include file="data/conn.asp"-->
<!--#include file="data/session.asp"-->
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
dim db_table
    db_table="article"
%>
<div class="mainWidth" id="ajax_main">
<div class="main_left">
<div class="main_left_box">
<div>
<input name="keyword" onkeyup='gAjax("<%=db_table%>_list.asp?keyword=","list",this.value);' type="text" class="search_bg" id="keyword" value="<%=keyword%>" size="10" maxlength="10"/>
</div>
<div class="leftNav">
<%
set rs=conn.execute("select * from "&db_table&"_type where type_id=0 order by order_id asc,id desc")
    do while not rs.eof
%>
<a href='javascript:gAjax("<%=db_table%>_list.asp?Type_id=","list",<%=rs("id")%>);'><%=rs("title")%></a>
<%
    rs.movenext
    loop
set rs=nothing
%>
<a href='javascript:gAjax("to_excel.asp?","list","");'>导(Excel)&nbsp;<img src="images/down_ico.gif" width="12" height="12" border="0" align="absmiddle" /></a></div>
<div class="clear"></div>
</div>
<div class="clear_padding"></div>

<div><img src="images/not7_logo.gif" width="110" border="0" /></div>

</div>

<div class="main_right">
<div class="main_right_border">
<div><table border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><input type="button" value="库存(Inventory)" onclick='gAjax("article_list.asp?Type_id=","list","");' /></td>
    <td><input type="button" value="出库(Shipped)" onclick='gAjax("shipped_list.asp?Type_id=","list","");' /></td>
    <td><input type="button" value="入库(Entry)" onclick='gAjax("admin_article_edit.asp?","list","");' /></td>
    <td><input type="button" value="型号(No.)" onclick='gAjax("admin_article_type.asp?","list","");' /></td>
    <td><input type="button" value="供应商" onclick='gAjax("admin_suppliers.asp?","list","");' /></td>
    <td width="40" align="center" style="line-height:15px;"><a style="text-decoration:underline" href='javascript:gAjax("edit_password.asp?","list","");'>密码?</a></td>
    <td width="60" align="center" style="line-height:15px;"><a style="text-decoration:underline" href="?login=out">注销(Out)</a></td>
  </tr>
  
</table>
</div>
<div id="ajax_list"></div>
</div></div>
<div class="clear"></div>
</div>
<div class="mainWidth">&nbsp;</div>

<script>
gAjax("<%=db_table%>_list.asp?Type_id=","list","");  //初始化页面内容
</script>
</body>
</html>

