<!--#include file="data/conn.asp"-->
<!--#include file="data/session.asp"-->
<%
'//////////////////////////////
'By   : cm.ivan
'Time : 0:51 2010-10-2
'For  : Inventory Management
'Email: cm.ivan@163.com
'//////////////////////////////

sub to_excel(types)
dim rs,sql,typesName,filename,fs,myfile,x,filedata,toFilename,toUrl
    if types=1 then '//存
       typesName="库存(Inventory)"
       sql= "select * from article where ok=0 order by id desc"
    elseif types=2 then '//销
       typesName="出库(Shipped)"
       sql= "select * from article where ok=1 order by id desc"
    else '//所有
       typesName="入库(All)"
       sql= "select * from article order by id desc"
    end if

    filedata  =year(now)&"_"&month(now)&"_"&day(now)&".xls"
    toFilename=typesName&"_"&filedata
    toUrl     ="download/"&toFilename
Set fs = server.CreateObject("scripting.filesystemobject")
'--假设你想让生成的EXCEL文件做如下的存放  
    filename = Server.MapPath("./"&toUrl)
'--如果原来的EXCEL文件存在的话就删除
 if fs.FileExists(filename) then fs.DeleteFile(filename)
'--创建EXCEL文件 
set myfile = fs.CreateTextFile(filename,true)
set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql,conn,1,1
if rs.EOF and rs.BOF then
else
dim strLine,responsestr
    strLine=""

 if types=1 then '//存
    strLine = strLine & "入库日期" & chr(9)
 elseif types=2 then '//销
    strLine = strLine & "出库日期" & chr(9)
 else '//所有
    strLine = strLine & "入库日期" & chr(9)
    strLine = strLine & "出库日期" & chr(9)
 end if

    strLine = strLine & "产品型号" & chr(9)
    strLine = strLine & "产品编码" & chr(9)
    strLine = strLine & "单价（元）" & chr(9)
    strLine = strLine & "数量（点）" & chr(9)
 if types<>2 then '//存
    strLine = strLine & "供应商" & chr(9)
 end if
 if types<>1 then '//销
    strLine = strLine & "接收单位" & chr(9)
 end if
    strLine = strLine & "备注" & chr(9)
    myfile.writeline strLine '--将表的列名先写入EXCEL

 Do while Not rs.EOF
    strLine=""
 if types=1 then '//存
    strLine = strLine & rs("Add_data") & chr(9)
 elseif types=2 then '//销
    strLine = strLine & rs("Shipped_Time") & chr(9)
 else '//所有
    strLine = strLine & rs("Add_data") & chr(9)
    strLine = strLine & rs("Shipped_Time") & chr(9)
 end if

    strLine = strLine & get_one("article_type",rs("Type_id"),"title") & chr(9)
    strLine = strLine & rs("title") & chr(9)
    strLine = strLine & rs("Price") & chr(9)
    strLine = strLine & rs("Num") & chr(9)
 if types<>2 then '//存
    strLine = strLine & get_one("suppliers",rs("suppliers_id"),"title") & chr(9)
 end if
 if types<>1 then '//存
    strLine = strLine & rs("R_unit") & chr(9)
 end if
    strLine = strLine & rs("note") & chr(9)
    myfile.writeline strLine '--将表的数据写入EXCEL
    rs.MoveNext
 loop

if err=0 then response.Write("<a href="""&toUrl&""" target=""download_box"">保存["&toFilename&"]</a><br>")

end if
rs.Close
set rs = nothing
end sub
%>  
<div style="padding:20px;">
<%
call to_excel(0)
call to_excel(1)
call to_excel(2)
%>
</div>
<iframe src="" name="download_box" scrolling="0" frameborder="0" height="1" width="1"></iframe>