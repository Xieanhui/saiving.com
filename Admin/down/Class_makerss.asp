<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
Dim Conn,User_Conn,strShowErr,Fs_news,obj_mf_sys_obj,MF_Domain,MF_Site_Name,tmp_c_path
MF_Default_Conn
MF_User_Conn
'session判断
MF_Session_TF 
'权限判断
set Fs_news = new Cls_News
Fs_News.GetSysParam()
set obj_mf_sys_obj = Conn.execute("select top 1 MF_Domain,MF_Site_Name from FS_MF_Config")
if obj_mf_sys_obj.eof then
	strShowErr = "<li>找不到主系统配置信息！</li>"
	Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	MF_Domain = obj_mf_sys_obj("MF_Domain")
	MF_Site_Name = obj_mf_sys_obj("MF_Site_Name")
end if
obj_mf_sys_obj.close:set obj_mf_sys_obj = nothing
tmp_c_path =MF_Domain &"/"&G_VIRTUAL_ROOT_DIR
'If Not Fs_news.IsSelfRefer Then response.write "非法提交数据":Response.end
dim Classid,signxml
Classid = NoSqlHack(Request.QueryString("cid"))
signxml = NoSqlHack(Request.QueryString("signxml"))
%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>栏目管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style>
.RefreshLen{
	height: 20px;
	width: 400px;
	border: 1px solid #104a7b;
	text-align: left;
	MARGIN-top:50px;
	margin-bottom: 5px;
}
</style>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">栏目管理<a href="../../help?Lable=NS_Class_Action" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><a href="Class_Manage.asp">管理首页</a>┆<a href="Class_add.asp?ClassID=&Action=add">添加根栏目</a>┆<a href="Class_Action.asp?Action=one">一级栏目排序</a>┆<a href="Class_Action.asp?Action=n">N级栏目排序</a>┆<a href="Class_Action.asp?Action=reset"   onClick="{if(confirm('确认复位所有栏目？\n\n如果选择确定，所有的栏目将设置为一级分类!!')){return true;}return false;}">复位所有栏目</a>┆<a href="Class_Action.asp?Action=unite">栏目合并</a>┆<a href="Class_Action.asp?Action=allmove">栏目转移</a> 
        ┆ <a href="Class_Action.asp?Action=clearClass"  onClick="{if(confirm('确认清空所有栏目里的数据吗？\n\n如果选择确定,所有的栏目的下载将被放到回收站中!!')){return true;}return false;}">删除所有栏目</a>┆<a href="../../help?Lable=NS_Class_Action_1" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div></td>
  </tr>
</table>
<%
'待修正加入RSS参数
Server.ScriptTimeOut=999999999
Dim ArrClassid,tmp_k,i,savepath,obj_classxml_rs,obj_news_rs,p1,xml_c_list,obj_class_rs,i_k,class_tmp
dim obj_all_rs,xml_c_all_list
savepath = Replace("\"&G_VIRTUAL_ROOT_DIR&"\xml\","\\","\")
if signxml="one" then
		set obj_news_rs = Server.CreateObject(G_FS_RS)
		obj_news_rs.open "select  top "& CintStr(Fs_news.rssNumber)&" id,newsid,newstitle,content,addtime,author From FS_NS_News where classid='"&NoSqlHack(classid)&"' and isdraft=0 and isRecyle=0 and isLock=0 order by PopId desc,id desc",conn,1,3
		set class_tmp = conn.execute("select ClassName,ClassEName,SavePath,IsURL,[Domain],FileExtName from FS_DS_Class where Classid='"& NoSqlHack(Classid) &"'")
		if class_tmp("IsURL")=1 then
			strShowErr = "<li>外部拦目不能生成！</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		call headxml(xml_c_list,class_tmp("ClassName"),class_tmp("ClassEName"),class_tmp("SavePath"),class_tmp("Domain"),class_tmp("FileExtName"))
		class_tmp.close:set class_tmp =nothing
		if not obj_news_rs.eof then
			do while not obj_news_rs.eof
				xml_c_list =  xml_c_list & "<item>"& chr(13) & chr(10)
				xml_c_list =  xml_c_list & "		<title>"& obj_news_rs("NewsTitle")&"</title>"& chr(13) & chr(10)
				xml_c_list =  xml_c_list & "		<link>"& obj_news_rs("NewsID")&"</link>"& chr(13) & chr(10)
				if trim(obj_news_rs("Content"))<>empty  or not isnull(trim(obj_news_rs("Content"))) then
					xml_c_list =  xml_c_list & "		<description><![CDATA["& GotTopic(obj_news_rs("Content"),Fs_news.rssContentNumber)&"]]></description>"& chr(13) & chr(10)
				else
					xml_c_list =  xml_c_list & "		<description><![CDATA[无内容]]></description>"& chr(13) & chr(10)
				end if
				xml_c_list =  xml_c_list & "		<pubDate>"& obj_news_rs("addtime")&"</pubDate>"& chr(13) & chr(10)
				xml_c_list =  xml_c_list & "		<author>"& obj_news_rs("Author")&"</author>"& chr(13) & chr(10)
				xml_c_list =  xml_c_list & "</item>"& chr(13) & chr(10)
				obj_news_rs.movenext
			loop
				xml_c_list =  xml_c_list & "</channel>"& chr(13) & chr(10)
				xml_c_list =  xml_c_list & "</rss>"& chr(13) & chr(10)
			call SaveFile(xml_c_list,Classid,"xml",savepath,"DS")
			strShowErr = "<li>Xml生成成功！</li>"
		else
			strShowErr = "<li>Xml失败，没有符合条件的下载！</li>"
		end if
		obj_news_rs.close:set obj_news_rs=nothing
		call newslist()
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
else
	if instr(Classid,",")=0 then 
		strShowErr = "<li>批量生成xml至少选择2项</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		call Allxml()
		call newslist()
	end if
end if
sub Allxml()
	ArrClassid =split(Classid,",")
	p1=UBound(ArrClassid)
	response.Write("<div style=""text-align: center;"">")
	response.Write("<div class=""RefreshLen""><div class=""xingmu"" id=""RefreshLen""></div></div><span id=""result_str""></span><br><br>")
	i=0
	i_k=0
	for  tmp_k = 0 to UBound(ArrClassid)
		set obj_class_rs = conn.execute("select classename,classname from FS_DS_Class where classid='"& NoSqlHack(ArrClassid(tmp_k)) &"'")
		if obj_class_rs.eof then
				i=i+1
		else
			set obj_news_rs = Server.CreateObject(G_FS_RS)
			obj_news_rs.open "select  top "& CintStr(Fs_news.rssNumber)&" id,newsid,newstitle,content,addtime,author From FS_NS_News where classid='"&NoSqlHack(ArrClassid(tmp_k)) &"' and isdraft=0 and isRecyle=0 and isLock=0 order by PopId desc,id desc",conn,1,3
			set class_tmp = conn.execute("select ClassName,ClassEName,SavePath,[Domain],FileExtName from FS_DS_Class where Classid='"& NoSqlHack(ArrClassid(tmp_k)) &"'")
			call headxml(xml_c_list,class_tmp("ClassName"),class_tmp("ClassEName"),class_tmp("SavePath"),class_tmp("Domain"),class_tmp("FileExtName"))
			class_tmp.close:set class_tmp =nothing
				if i = p1 then
					Response.Write "<script>RefreshLen.style.width ="""&int(i/p1*100)&"%"";RefreshLen.innerHTML=""&nbsp;<font color=black>"&int(i/p1*100)&"%</font>"";result_str.innerHTML=""当前栏目:"&obj_class_rs("classname")&"<br><span class=tx>生成完毕....ok!!共:"& i+1 &"个,生成了"& i_k &"个&nbsp;&nbsp;&nbsp;未生成原因：栏目下找不到符合条件的RSS聚合下载</span>"";</script>" & VbCrLf
				else
					Response.Write "<script>RefreshLen.style.width ="""&int(i/p1*100)&"%"";RefreshLen.innerHTML=""&nbsp;<font color=black>"&int(i/p1*100)&"%</font>"";result_str.innerHTML=""当前栏目:"&obj_class_rs("classname")&""";</script>" & VbCrLf
				end if
				Response.Flush
			if not obj_news_rs.eof then
				do while not obj_news_rs.eof
					xml_c_list =  xml_c_list & "<item>"& chr(13) & chr(10)
					xml_c_list =  xml_c_list & "		<title>"& obj_news_rs("NewsTitle")&"</title>"& chr(13) & chr(10)
					xml_c_list =  xml_c_list & "		<link>"& obj_news_rs("NewsID")&"</link>"& chr(13) & chr(10)
					if trim(obj_news_rs("Content"))<>empty  or not isnull(trim(obj_news_rs("Content"))) then
						xml_c_list =  xml_c_list & "		<description><![CDATA["& GotTopic(obj_news_rs("Content"),Fs_news.rssContentNumber)&"]]></description>"& chr(13) & chr(10)
					else
						xml_c_list =  xml_c_list & "		<description><![CDATA[无内容]]></description>"& chr(13) & chr(10)
					end if
					xml_c_list =  xml_c_list & "		<pubDate>"& obj_news_rs("addtime")&"</pubDate>"& chr(13) & chr(10)
					xml_c_list =  xml_c_list & "		<author>"& obj_news_rs("Author")&"</author>"& chr(13) & chr(10)
					xml_c_list =  xml_c_list & "</item>"& chr(13) & chr(10)
					obj_news_rs.movenext
				loop
					xml_c_list =  xml_c_list & "</channel>"& chr(13) & chr(10)
					xml_c_list =  xml_c_list & "</rss>"& chr(13) & chr(10)
				call SaveFile(xml_c_list,ArrClassid(tmp_k),"xml",savepath,"DS")
				set obj_news_rs=nothing
				i_k = i_k + 1
			end  if
			i=i+1
		end if
	next
		response.Write("</div>")
End sub
sub newslist()
		set obj_all_rs = Server.CreateObject(G_FS_RS)
		obj_all_rs.open "select  top "& Fs_news.rssNumber&" id,newsid,newstitle,content,addtime,author From FS_NS_News where isdraft=0 and isRecyle=0 and isLock=0 order by PopId desc,id desc",conn,1,3
		if not obj_all_rs.eof then
			set class_tmp = conn.execute("select ClassName,ClassEName,SavePath,[Domain],FileExtName from FS_DS_Class where Classid='"& NoSqlHack(Classid) &"'")
				xml_c_all_list = "<?xml version=""1.0"" encoding=""gb2312""?>" & chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "<rss version=""2.0"">"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "<channel>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "<title>"& Fs_news.siteName &"</title>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "<image>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "		<title>"& Fs_news.siteName &"</title>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "		<link>http://" & Replace(tmp_c_path &Fs_news.RSSPIC,"//","/") &"</link>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "		<url>http://" & Replace(tmp_c_path &Fs_news.RSSPIC,"//","/") &"</url>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "</image>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "<description>"& Fs_news.rssdescript&"</description>"& chr(13) & chr(10)
			class_tmp.close:set class_tmp =nothing
			do while not obj_all_rs.eof
				xml_c_all_list =  xml_c_all_list & "<item>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "		<title>"& obj_all_rs("NewsTitle")&"</title>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "		<link>"& obj_all_rs("NewsID")&"</link>"& chr(13) & chr(10)
				if trim(obj_all_rs("Content"))<>empty  or not isnull(trim(obj_all_rs("Content"))) then
					xml_c_all_list =  xml_c_all_list & "		<description><![CDATA["& GotTopic(obj_all_rs("Content"),Fs_news.rssContentNumber)&"]]></description>"& chr(13) & chr(10)
				else
					xml_c_all_list =  xml_c_all_list & "		<description><![CDATA[无内容]]></description>"& chr(13) & chr(10)
				end if
				xml_c_all_list =  xml_c_all_list & "		<pubDate>"& obj_all_rs("addtime")&"</pubDate>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "		<author>"& obj_all_rs("Author")&"</author>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "</item>"& chr(13) & chr(10)
				obj_all_rs.movenext
			loop
				xml_c_all_list =  xml_c_all_list & "</channel>"& chr(13) & chr(10)
				xml_c_all_list =  xml_c_all_list & "</rss>"& chr(13) & chr(10)
				DIM savepath1
			call SaveFile(xml_c_all_list,"now","xml",savepath,"DS")
	end if
	obj_all_rs.close:set obj_all_rs= nothing
end sub
function headxml(f_char,f_className,f_classeName,f_SavePath,f_Domain,f_FileExtName)
		f_char = "<?xml version=""1.0"" encoding=""gb2312""?>" & chr(13) & chr(10)
		f_char =  f_char & "<rss version=""2.0"">"& chr(13) & chr(10)
		f_char =  f_char & "<channel>"& chr(13) & chr(10)
		f_char =  f_char & "<title>"& Fs_news.siteName &"</title>"& chr(13) & chr(10)
		f_char =  f_char & "<image>"& chr(13) & chr(10)
		f_char =  f_char & "		<title>"& Fs_news.siteName &"</title>"& chr(13) & chr(10)
		f_char =  f_char & "		<link>"& f_classeName &"</link>"& chr(13) & chr(10)
		f_char =  f_char & "		<url>http://"&Replace(tmp_c_path &Fs_news.RSSPIC,"//","/")&"</url>"& chr(13) & chr(10)
		f_char =  f_char & "</image>"& chr(13) & chr(10)
		f_char =  f_char & "<description>"& Fs_news.rssdescript&"</description>"& chr(13) & chr(10)
		headxml = f_char
end function
%>
</body>
</html>
<%
set Fs_news = nothing
%>






