<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/HS_Function.asp" -->
<!--#include file="../../FS_InterFace/AP_Function.asp" -->
<!--#include file="Func_page.asp" -->
<!--#include file="../../FS_Inc/Cls_SysConfig.asp"-->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn,label_Conn,Labelclass_SQL,obj_Labelclass_rs,label_rs,ssql,label_srs
	MF_Default_Conn
	MF_Session_TF 
	MF_Label_Conn
    dim actions,sql
	actions = request.QueryString("action")
	if actions="out" then
			dim osql,news_rss,csql,c_rs
			osql="delete from FS_Lable"
			set news_rss = label_Conn.execute(osql)
			
			csql="delete from FS_LableClass"
			set c_rs = label_Conn.execute(csql)
			
			Labelclass_SQL = "Select * from FS_MF_Lable where isDel=0 order by id desc"
			Set obj_Labelclass_rs = server.CreateObject(G_FS_RS)
			obj_Labelclass_rs.Open Labelclass_SQL,Conn,1,3
			do while not obj_Labelclass_rs.eof 
				sql ="select * from FS_Lable where 1=1"
				Set label_rs = server.CreateObject(G_FS_RS)
				label_rs.Open sql,label_Conn,1,3
				label_rs.addnew
				label_rs("LableName")=obj_Labelclass_rs("LableName")
				label_rs("LableDesc")=obj_Labelclass_rs("LableDesc")
				label_rs("LableContent")=obj_Labelclass_rs("LableContent")
				label_rs("AddDate")=obj_Labelclass_rs("AddDate")
				label_rs("LablelType")=obj_Labelclass_rs("LablelType")
				label_rs("LableClassID")=obj_Labelclass_rs("LableClassID")
				label_rs.update
				label_rs.close:Set label_rs = nothing
				obj_Labelclass_rs.movenext	
			loop

			dim c_SQL,obj_cclass_rs
			dim lsql,l_rs
			c_SQL = "Select * from FS_MF_Lableclass order by id desc"
			Set obj_cclass_rs = server.CreateObject(G_FS_RS)
			obj_cclass_rs.Open c_SQL,Conn,1,3
			do while not obj_cclass_rs.eof 
				lsql ="select * from FS_LableClass where 1=1"
				Set l_rs = server.CreateObject(G_FS_RS)
				l_rs.Open lsql,label_Conn,1,3
				l_rs.addnew
				l_rs("id")=obj_cclass_rs("id")
				l_rs("ClassName")=obj_cclass_rs("ClassName")
				l_rs("ClassContent")=obj_cclass_rs("ClassContent")
				l_rs("ParentID")=obj_cclass_rs("ParentID")
				l_rs("ClassType")=obj_cclass_rs("ClassType")
				l_rs.update
				l_rs.close:Set l_rs = nothing
				obj_cclass_rs.movenext	
			loop
			
			obj_cclass_rs.close:Set obj_cclass_rs = nothing
			response.Write("<script>alert(""导出标签成功。标签存放于"&G_LABEL_DATA_STR&"表中。,点击确定下载数据库到本地"");location.href="""&G_VIRTUAL_ROOT_DIR&"/"&G_LABEL_DATA_STR&"""</script>")
			response.End()
	else
			Labelclass_SQL = "Select * from FS_Lable order by id desc"
			
			Set obj_Labelclass_rs = server.CreateObject(G_FS_RS)
			obj_Labelclass_rs.Open Labelclass_SQL,label_Conn,1,3
			do while not obj_Labelclass_rs.eof 
				sql ="select * from FS_MF_Lable where 1=1"
				Set label_rs = server.CreateObject(G_FS_RS)
				label_rs.Open sql,Conn,1,3
				ssql="select count(*) from FS_MF_Lable where LableName='"&obj_Labelclass_rs("LableName")&"'"
				set label_srs = conn.execute(ssql)
				label_rs.addnew
				if label_srs(0)>0 then
					randomize 
					dim rnb
					rnb = Int(10 * Rnd + 10)
					label_rs("LableName")=replace(obj_Labelclass_rs("LableName"),"}","_"&rnb&"}")
				else
					label_rs("LableName")=obj_Labelclass_rs("LableName")
				end if
				label_rs("LableDesc")=obj_Labelclass_rs("LableDesc")
				label_rs("LableContent")=obj_Labelclass_rs("LableContent")
				label_rs("AddDate")=obj_Labelclass_rs("AddDate")
				label_rs("LablelType")=obj_Labelclass_rs("LablelType")
				label_rs("LableClassID")=obj_Labelclass_rs("LableClassID")
				label_rs.update
				label_rs.close:Set label_rs = nothing
				obj_Labelclass_rs.movenext	
			loop
	
			dim insql,in_rs,iisql,iirs
			insql = "Select * from FS_Lableclass order by id desc"
			
			Set in_rs = server.CreateObject(G_FS_RS)
			in_rs.Open insql,label_Conn,1,3
			do while not in_rs.eof 
				iisql ="select * from FS_MF_Lableclass where ClassName='"&in_rs("ClassName")&"'"
				Set iirs = server.CreateObject(G_FS_RS)
				iirs.Open iisql,Conn,1,3
				if iirs.eof then
					iirs.addnew
					iirs("ClassName")=in_rs("ClassName")
					iirs("ClassContent")=in_rs("ClassContent")
					iirs("ParentID")=in_rs("ParentID")
					iirs("ClassType")=in_rs("ClassType")
					iirs.update
					iirs.close:Set iirs = nothing
				end if
				in_rs.movenext	
			loop
			
			
			in_rs.close:Set in_rs = nothing
			response.Write("<script>alert(""导入标签成功。"");location.href=""all_label_stock.asp""</script>")
			response.End()
	
	end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>导入导出</title>
</head>

<body>
</body>
</html>
