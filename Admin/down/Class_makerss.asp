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
'session�ж�
MF_Session_TF 
'Ȩ���ж�
set Fs_news = new Cls_News
Fs_News.GetSysParam()
set obj_mf_sys_obj = Conn.execute("select top 1 MF_Domain,MF_Site_Name from FS_MF_Config")
if obj_mf_sys_obj.eof then
	strShowErr = "<li>�Ҳ�����ϵͳ������Ϣ��</li>"
	Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	MF_Domain = obj_mf_sys_obj("MF_Domain")
	MF_Site_Name = obj_mf_sys_obj("MF_Site_Name")
end if
obj_mf_sys_obj.close:set obj_mf_sys_obj = nothing
tmp_c_path =MF_Domain &"/"&G_VIRTUAL_ROOT_DIR
'If Not Fs_news.IsSelfRefer Then response.write "�Ƿ��ύ����":Response.end
dim Classid,signxml
Classid = NoSqlHack(Request.QueryString("cid"))
signxml = NoSqlHack(Request.QueryString("signxml"))
%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ŀ����___Powered by foosun Inc.</title>
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
    <td class="xingmu">��Ŀ����<a href="../../help?Lable=NS_Class_Action" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><a href="Class_Manage.asp">������ҳ</a>��<a href="Class_add.asp?ClassID=&Action=add">��Ӹ���Ŀ</a>��<a href="Class_Action.asp?Action=one">һ����Ŀ����</a>��<a href="Class_Action.asp?Action=n">N����Ŀ����</a>��<a href="Class_Action.asp?Action=reset"   onClick="{if(confirm('ȷ�ϸ�λ������Ŀ��\n\n���ѡ��ȷ�������е���Ŀ������Ϊһ������!!')){return true;}return false;}">��λ������Ŀ</a>��<a href="Class_Action.asp?Action=unite">��Ŀ�ϲ�</a>��<a href="Class_Action.asp?Action=allmove">��Ŀת��</a> 
        �� <a href="Class_Action.asp?Action=clearClass"  onClick="{if(confirm('ȷ�����������Ŀ���������\n\n���ѡ��ȷ��,���е���Ŀ�����ؽ����ŵ�����վ��!!')){return true;}return false;}">ɾ��������Ŀ</a>��<a href="../../help?Lable=NS_Class_Action_1" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div></td>
  </tr>
</table>
<%
'����������RSS����
Server.ScriptTimeOut=999999999
Dim ArrClassid,tmp_k,i,savepath,obj_classxml_rs,obj_news_rs,p1,xml_c_list,obj_class_rs,i_k,class_tmp
dim obj_all_rs,xml_c_all_list
savepath = Replace("\"&G_VIRTUAL_ROOT_DIR&"\xml\","\\","\")
if signxml="one" then
		set obj_news_rs = Server.CreateObject(G_FS_RS)
		obj_news_rs.open "select  top "& CintStr(Fs_news.rssNumber)&" id,newsid,newstitle,content,addtime,author From FS_NS_News where classid='"&NoSqlHack(classid)&"' and isdraft=0 and isRecyle=0 and isLock=0 order by PopId desc,id desc",conn,1,3
		set class_tmp = conn.execute("select ClassName,ClassEName,SavePath,IsURL,[Domain],FileExtName from FS_DS_Class where Classid='"& NoSqlHack(Classid) &"'")
		if class_tmp("IsURL")=1 then
			strShowErr = "<li>�ⲿ��Ŀ�������ɣ�</li>"
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
					xml_c_list =  xml_c_list & "		<description><![CDATA[������]]></description>"& chr(13) & chr(10)
				end if
				xml_c_list =  xml_c_list & "		<pubDate>"& obj_news_rs("addtime")&"</pubDate>"& chr(13) & chr(10)
				xml_c_list =  xml_c_list & "		<author>"& obj_news_rs("Author")&"</author>"& chr(13) & chr(10)
				xml_c_list =  xml_c_list & "</item>"& chr(13) & chr(10)
				obj_news_rs.movenext
			loop
				xml_c_list =  xml_c_list & "</channel>"& chr(13) & chr(10)
				xml_c_list =  xml_c_list & "</rss>"& chr(13) & chr(10)
			call SaveFile(xml_c_list,Classid,"xml",savepath,"DS")
			strShowErr = "<li>Xml���ɳɹ���</li>"
		else
			strShowErr = "<li>Xmlʧ�ܣ�û�з������������أ�</li>"
		end if
		obj_news_rs.close:set obj_news_rs=nothing
		call newslist()
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
else
	if instr(Classid,",")=0 then 
		strShowErr = "<li>��������xml����ѡ��2��</li>"
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
					Response.Write "<script>RefreshLen.style.width ="""&int(i/p1*100)&"%"";RefreshLen.innerHTML=""&nbsp;<font color=black>"&int(i/p1*100)&"%</font>"";result_str.innerHTML=""��ǰ��Ŀ:"&obj_class_rs("classname")&"<br><span class=tx>�������....ok!!��:"& i+1 &"��,������"& i_k &"��&nbsp;&nbsp;&nbsp;δ����ԭ����Ŀ���Ҳ�������������RSS�ۺ�����</span>"";</script>" & VbCrLf
				else
					Response.Write "<script>RefreshLen.style.width ="""&int(i/p1*100)&"%"";RefreshLen.innerHTML=""&nbsp;<font color=black>"&int(i/p1*100)&"%</font>"";result_str.innerHTML=""��ǰ��Ŀ:"&obj_class_rs("classname")&""";</script>" & VbCrLf
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
						xml_c_list =  xml_c_list & "		<description><![CDATA[������]]></description>"& chr(13) & chr(10)
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
					xml_c_all_list =  xml_c_all_list & "		<description><![CDATA[������]]></description>"& chr(13) & chr(10)
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






