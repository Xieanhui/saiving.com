<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_InterFace/CLS_Foosun.asp" -->
<!--#include file="../../FS_Inc/Cls_SysConfig.asp" -->
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
if not Get_SubPop_TF(Classid,"NS022","NS","class") then
	Response.Redirect("lib/error.asp?ErrCodes=缺少权限&ErrorUrl=")
	Response.end
End if
Dim sysObj
Set sysObj=New Cls_SysConfig
sysObj.getSysParam()

Function GetOneNewsLink(f_RS)
	Dim f_NewsLinkRecordSet,f_NewsLink
	Set f_NewsLinkRecordSet = New CLS_FoosunRecordSet
	Set f_NewsLinkRecordSet.Values("ClassEName,Domain,SavePath,IsURL,URLAddress,SaveNewsPath,FileName,FileExtName") = f_RS
	Set f_NewsLink = New CLS_FoosunLink
	GetOneNewsLink = f_NewsLink.NewsLink(f_NewsLinkRecordSet)
	Set f_NewsLink = Nothing
	Set f_NewsLinkRecordSet = Nothing
End Function

Dim AD_EmailStr
AD_EmailStr = Conn.ExeCute("Select MF_Mail_Name From FS_MF_Config")(0)
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
        ┆ <a href="Class_Action.asp?Action=clearClass"  onClick="{if(confirm('确认清空所有栏目里的数据吗？\n\n如果选择确定,所有的栏目的新闻将被放到回收站中!!')){return true;}return false;}">删除所有栏目</a>┆<a href="../../help?Lable=NS_Class_Action_1" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div></td>
  </tr>
</table>
<%
Server.ScriptTimeOut=999999999
Dim ArrClassid,tmp_k,i,savepath,obj_classxml_rs,obj_news_rs,p1,xml_c_list,obj_class_rs,i_k,class_tmp
dim obj_all_rs,xml_c_all_list
savepath = Replace("\"&G_VIRTUAL_ROOT_DIR&"\xml\","\\","\")
If signxml="one" Then
	set obj_news_rs = Server.CreateObject(G_FS_RS)
	obj_news_rs.open "select top "& Fs_news.rssNumber&" News.id,newsid,newstitle,content,News.addtime,author,NewsPicFile,Class.IsURL as ClassIsURL,Class.FileExtName as ClassFileExtName,ClassName,ClassEName,[Domain],SavePath,News.IsURL,News.URLAddress,SaveNewsPath,FileName,News.FileExtName From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.classid='"&classid&"' and isdraft=0 and isRecyle=0 and isLock=0 order by PopId desc,News.id desc",conn,0,1
	if not obj_news_rs.eof then
		if obj_news_rs("ClassIsURL")=1 then
			strShowErr = "<li>外部拦目不能生成！</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
			Set sysObj = Nothing
		end if
		call headxml(xml_c_list,obj_news_rs("ClassName"),obj_news_rs("ClassEName"),obj_news_rs("SavePath"),obj_news_rs("Domain"),obj_news_rs("ClassFileExtName"))
		do while not obj_news_rs.eof
			xml_c_list =  xml_c_list & "  <item>"& chr(13) & chr(10)
			xml_c_list =  xml_c_list & "	<title>"& obj_news_rs("NewsTitle")&"</title>"& chr(13) & chr(10)
			xml_c_list =  xml_c_list & "	<link>"&GetOneNewsLink(obj_news_rs)&"</link>"& chr(13) & chr(10)
			if trim(obj_news_rs("Content"))<>empty  or not isnull(trim(obj_news_rs("Content"))) then
				xml_c_list =  xml_c_list & "	<text>"& LoseEnter(GotTopic(Lose_Html(ReplaceHtml(obj_news_rs("Content"))),Fs_news.rssContentNumber))&"</text>"& chr(13) & chr(10)
			else
				xml_c_list =  xml_c_list & "	<text>无内容</text>"& chr(13) & chr(10)
			end if
			IF trim(obj_news_rs("NewsPicFile")) <> "" And Not IsNull(trim(obj_news_rs("NewsPicFile"))) Then
				xml_c_list =  xml_c_list & "	<image>" & obj_news_rs("NewsPicFile") & "</image>"& chr(13) & chr(10)
			Else
				xml_c_list =  xml_c_list & "	<image></image>"& chr(13) & chr(10)
			End IF		
			xml_c_list =  xml_c_list & "	<author>"& obj_news_rs("Author")&"</author>"& chr(13) & chr(10)
			xml_c_list =  xml_c_list & "	<pubDate>"& obj_news_rs("addtime")&"</pubDate>"& chr(13) & chr(10)
			xml_c_list =  xml_c_list & "  </item>"& chr(13) & chr(10)
			obj_news_rs.movenext
		loop
		xml_c_list =  xml_c_list & "</document>"& chr(13) & chr(10)
		call SaveFile(xml_c_list,Classid,"xml",savepath,"NS")
		strShowErr = "<li>Xml生成成功！</li>"
	else
		strShowErr = "<li>Xml失败，没有符合条件的新闻！</li>"
	end if
	obj_news_rs.close:set obj_news_rs=nothing
	call newslist()
	Set sysObj = Nothing
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
Else
	If signxml="All" Then
		Classid=""
		Set class_tmp = conn.execute("select ClassID from FS_NS_NewsClass where IsURL=0")
		While Not class_tmp.Eof
			If Classid="" Then
				Classid=class_tmp(0)
			Else
				Classid=Classid&","&class_tmp(0)
			End If
			class_tmp.MoveNext
		Wend
		class_tmp.close
		Set class_tmp=Nothing
	End If
	If instr(Classid,",")=0 Then 
		strShowErr = "<li>批量生成xml至少选择2项</li>"
		Set sysObj = Nothing
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.End 
	Else 
		Call Allxml()
		Call newslist()
	End If 
End if
sub Allxml()
	ArrClassid =split(Classid,",")
	p1=UBound(ArrClassid)
	response.Write("<div style=""text-align: center;"">")
	response.Write("<div class=""RefreshLen""><div class=""xingmu"" id=""RefreshLen""></div></div><span id=""result_str""></span><br><br>")
	i=0
	i_k=0
	for tmp_k = 0 to UBound(ArrClassid)
		set obj_class_rs = conn.execute("select classename,classname,IsURL from FS_NS_NewsClass where classid='"& ArrClassid(tmp_k) &"'")
		if obj_class_rs.eof then
			i=i+1
		else
			if obj_class_rs("IsURL") = 0 then
				set obj_news_rs = Server.CreateObject(G_FS_RS)
				obj_news_rs.open "select  top "& Fs_news.rssNumber&" News.id,newsid,newstitle,content,News.addtime,author,NewsPicFile,ClassName,ClassEName,[Domain],SavePath,News.IsURL,News.URLAddress,SaveNewsPath,FileName,News.FileExtName,Class.FileExtName as ClassFileExtName From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And News.classid='"&ArrClassid(tmp_k) &"' and isdraft=0 and isRecyle=0 and isLock=0 order by PopId desc,News.id desc",conn,1,3
				if not obj_news_rs.eof then
					call headxml(xml_c_list,obj_news_rs("ClassName"),obj_news_rs("ClassEName"),obj_news_rs("SavePath"),obj_news_rs("Domain"),obj_news_rs("ClassFileExtName"))
					if i = p1 then
						Response.Write "<script>RefreshLen.style.width ="""&int(i/p1*100)&"%"";RefreshLen.innerHTML=""&nbsp;<span class=xingmu>"&int(i/p1*100)&"%</span>"";result_str.innerHTML=""当前栏目:"&obj_news_rs("classname")&"<br><span class=tx>生成完毕....ok!!共:"& i+1 &"个,生成了"& i_k &"个&nbsp;&nbsp;&nbsp;未生成原因：栏目下找不到符合条件的RSS聚合新闻</span>"";</script>" & VbCrLf
					else
						Response.Write "<script>RefreshLen.style.width ="""&int(i/p1*100)&"%"";RefreshLen.innerHTML=""&nbsp;<span class=xingmu>"&int(i/p1*100)&"%</span>"";result_str.innerHTML=""当前栏目:"&obj_news_rs("classname")&""";</script>" & VbCrLf
					end if
					Response.Flush
					do while not obj_news_rs.eof
						xml_c_list =  xml_c_list & "  <item>"& chr(13) & chr(10)
						xml_c_list =  xml_c_list & "	<title>"& obj_news_rs("NewsTitle")&"</title>"& chr(13) & chr(10)
						xml_c_list =  xml_c_list & "	<link>"&GetOneNewsLink(obj_news_rs) &"</link>"& chr(13) & chr(10)
						if trim(obj_news_rs("Content"))<>empty  or not isnull(trim(obj_news_rs("Content"))) then
							xml_c_list =  xml_c_list & "	<text>"& LoseEnter(GotTopic(Lose_Html(ReplaceHtml(obj_news_rs("Content"))),Fs_news.rssContentNumber))&"</text>"& chr(13) & chr(10)
						else
							xml_c_list =  xml_c_list & "	<text>无内容</text>"& chr(13) & chr(10)
						end if
						IF trim(obj_news_rs("NewsPicFile")) <> "" And Not IsNull(trim(obj_news_rs("NewsPicFile"))) Then
							xml_c_list =  xml_c_list & "	<image>" & obj_news_rs("NewsPicFile") & "</image>"& chr(13) & chr(10)
						Else
							xml_c_list =  xml_c_list & "	<image></image>"& chr(13) & chr(10)
						End IF		
						xml_c_list =  xml_c_list & "	<author>"& obj_news_rs("Author")&"</author>"& chr(13) & chr(10)
						xml_c_list =  xml_c_list & "	<pubDate>"& obj_news_rs("addtime")&"</pubDate>"& chr(13) & chr(10)
						xml_c_list =  xml_c_list & "  </item>"& chr(13) & chr(10)
						obj_news_rs.movenext
					loop
					xml_c_list =  xml_c_list & "</document>"& chr(13) & chr(10)
					call SaveFile(xml_c_list,ArrClassid(tmp_k),"xml",savepath,"NS")
					set obj_news_rs=nothing
					i_k = i_k + 1
				end  if
			end if
			i=i+1
		end if
		obj_class_rs.Close
		Set obj_class_rs = Nothing
	next
	response.Write("</div>")
End sub
sub newslist()
	set obj_all_rs = Server.CreateObject(G_FS_RS)
	obj_all_rs.open "select  top "& Fs_news.rssNumber&" News.id,newsid,newstitle,content,News.addtime,author,NewsPicFile,ClassName,ClassEName,[Domain],SavePath,News.IsURL,News.URLAddress,SaveNewsPath,FileName,News.FileExtName,Class.FileExtName as ClassFileExtName From FS_NS_News as News,FS_NS_NewsClass as Class where News.ClassID=Class.ClassID And isdraft=0 and isRecyle=0 and isLock=0 order by PopId desc,Class.id desc",conn,1,3
	if not obj_all_rs.eof then
		xml_c_all_list = "<?xml version=""1.0"" encoding=""gb2312""?>" & chr(13) & chr(10)
		xml_c_all_list = xml_c_all_list &"<?xml-stylesheet type=""text/xsl"" href=""/sys_images/Rss.xsl""?>"& chr(13) & chr(10)
		xml_c_all_list =  xml_c_all_list & "<document>"& chr(13) & chr(10)
		xml_c_all_list =  xml_c_all_list & "  <webSite>"& Fs_news.siteName &"</webSite>"& chr(13) & chr(10)
		xml_c_all_list =  xml_c_all_list & "  <webMaster>" & AD_EmailStr & "</webMaster>"& chr(13) & chr(10)
		xml_c_all_list =  xml_c_all_list & "  <updatePeri>60</updatePeri>"& chr(13) & chr(10)
		do while not obj_all_rs.eof
			xml_c_all_list =  xml_c_all_list & "  <item>"& chr(13) & chr(10)
			xml_c_all_list =  xml_c_all_list & "	<title>"& obj_all_rs("NewsTitle")&"</title>"& chr(13) & chr(10)
			xml_c_all_list =  xml_c_all_list & "	<link>"&GetOneNewsLink(obj_all_rs) &"</link>"& chr(13) & chr(10)
			if trim(obj_all_rs("Content"))<>empty  or not isnull(trim(obj_all_rs("Content"))) then
				xml_c_all_list =  xml_c_all_list & "	<text>"& LoseEnter(GotTopic(ClearHtml(Lose_Html(ReplaceHtml(obj_all_rs("Content")))),Fs_news.rssContentNumber))&"</text>"& chr(13) & chr(10)
			else
				xml_c_all_list =  xml_c_all_list & "	<text>无内容</text>"& chr(13) & chr(10)
			end if
			IF trim(obj_all_rs("NewsPicFile")) <> "" And Not IsNull(trim(obj_all_rs("NewsPicFile"))) Then
				xml_c_all_list =  xml_c_all_list & "	<image>" & obj_all_rs("NewsPicFile") & "</image>"& chr(13) & chr(10)
			Else
				xml_c_all_list =  xml_c_all_list & "	<image></image>"& chr(13) & chr(10)
			End IF
			xml_c_all_list =  xml_c_all_list & "	<author>"& obj_all_rs("Author")&"</author>"& chr(13) & chr(10)
			xml_c_all_list =  xml_c_all_list & "	<pubDate>"& obj_all_rs("addtime")&"</pubDate>"& chr(13) & chr(10)
			xml_c_all_list =  xml_c_all_list & "  </item>"& chr(13) & chr(10)
			obj_all_rs.movenext
		loop
			xml_c_all_list =  xml_c_all_list & "</document>"& chr(13) & chr(10)
			DIM savepath1
		call SaveFile(xml_c_all_list,"now","xml",savepath,"NS")
	end if
	obj_all_rs.close:set obj_all_rs= nothing
end sub
function headxml(f_char,f_className,f_classeName,f_SavePath,f_Domain,f_FileExtName)
	f_char = "<?xml version=""1.0"" encoding=""gb2312""?>" & chr(13) & chr(10)
	f_char = f_char &"<?xml-stylesheet type=""text/xsl"" href=""/sys_images/Rss.xsl""?>" & chr(13) & chr(10)
	f_char =  f_char & "<document>"& chr(13) & chr(10)
	f_char =  f_char & "  <webSite>"& Fs_news.siteName &"</webSite>"& chr(13) & chr(10)
	f_char =  f_char & "  <webMaster>" & AD_EmailStr & "</webMaster>"& chr(13) & chr(10)
	f_char =  f_char & "  <updatePeri>60</updatePeri>"& chr(13) & chr(10)
	headxml = f_char
end function

Function ReplaceHtml(Str)
	Dim Str_Con
	Str_Con = Str & ""
	Str_Con = Replace(Str_Con,"&lt;","<")
	Str_Con = Replace(Str_Con,"&gt;",">")
	Str_Con = Replace(Str_Con,"&nbsp;"," ")
	ReplaceHtml = Str_Con
End Function

Function ClearHtml(Str)
	Dim objRegExp, strOutput
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "<.+?>"
	strOutput = objRegExp.Replace(Str, "")
	strOutput = Replace(strOutput, "<", "<")
	strOutput = Replace(strOutput, ">", ">")
	ClearHtml = strOutput 'Return the value of strOutput
	
	Set objRegExp = Nothing
End Function
Function LoseEnter(Str)
	LoseEnter = Replace(Str,Chr(13) & Chr(10),"")
End Function
%>
</body>
</html>
<%
set Fs_news = nothing
Set sysObj = Nothing
%>