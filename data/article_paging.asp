<div class="list_paging">
<form method=get onsubmit=""document.location = '" & action & "?" & temp & "page='+ this.page.value;return false;"">
<%if page<=1 then%>
首页 上一页
<%else%>   
<a class="liti" href='javascript:gAjax("<%=db_table%>_list.asp?<%=temp%>page=","list",1);'>首页</A>
<a class="liti" href='javascript:gAjax("<%=db_table%>_list.asp?<%=temp%>page=","list",<%=Page-1%>);'>上一页</A>
<%end if%>
	
<%if page>=maxpage then%>
下一页 尾页
<%else%>   
<a class="liti" href='javascript:gAjax("<%=db_table%>_list.asp?<%=temp%>page=","list",<%=Page+1%>);'>下一页</A>
<a class="liti" href='javascript:gAjax("<%=db_table%>_list.asp?<%=temp%>page=","list",<%=maxpage%>);'>尾页</A>    
<%end if%>
页次：<%=page%>/<%=maxpage%>页
共 <%=iCount%> 条记录
</form>   
</div>