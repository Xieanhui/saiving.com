<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<!--#include file="../FS_Inc/Cls_SysConfig.asp"-->
<%
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
MF_Session_TF

'---2007-02-12 By Ken
Dim ConStrRs,ConStrNum
Set ConStrRs = Server.CreateObject(G_FS_RS)
ConStrRs.Open "Select ContID from FS_ME_InfoContribution where IsPublic=1 and IsLock=0 and AuditTF=0",User_Conn,1,1
If ConStrRs.Eof then
	ConStrNum = 0
Else
	ConStrNum = Clng(ConStrRs.RecordCount)
End If	
ConStrRs.Close : Set ConStrRs = NOthing

Dim MaxDefineNum,GetSysConfigObj
Set GetSysConfigObj = New Cls_SysConfig
GetSysConfigObj.getSysParam()
MaxDefineNum = Clng(GetSysConfigObj.Define_MaxNum)
%>
<HTML>
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../FS_Inc/Prototype.js"></script>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" height="62" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<tr class="back">
		<td class="xingmu">��ӭʹ�÷�Ѷ��վ����ϵͳ(FoosunCMS)V<%=Request.Cookies("FoosunMFCookies")("FoosunMFVersion")%> For ASP Version����������Ȩ�ţ�2004SR11453</td>
	</tr>
	<tr class="back">
		<td height="22" class="hback">
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
				<tr>
					<td width="30%">�汾��: <%=Request.Cookies("FoosunMFCookies")("FoosunMFVersion")%> ����</td>
					<td width="70%" align="right">
						<div id="Foosun_server_version"></div>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="back">
		<td height="1" class="hback">
			<div id="Foosun_server_announce"></div>
		</td>
	</tr>
</table>
<table width="98%" height="166" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<%
	dim tmp_news_rs,tmp_sub_rs,tmp_Log_s_rs,tmp_Log_login_rs,tmp_admin_rs,tmp_define_rs
	Set tmp_news_rs	= CreateObject(G_FS_RS)
	tmp_news_rs.Open "select ID from FS_NS_News  where islock=1 and isRecyle=0 and isdraft=0 order by id asc",Conn,1,3
	Set tmp_sub_rs	= CreateObject(G_FS_RS)
	tmp_sub_rs.Open "select ID from FS_MF_Sub_Sys order by id asc",Conn,1,3
	Set tmp_Log_login_rs	= CreateObject(G_FS_RS)
	tmp_Log_login_rs.Open "select ID from FS_MF_Login_Log order by id asc",Conn,1,3
	Set tmp_Log_s_rs	= CreateObject(G_FS_RS)
	tmp_Log_s_rs.Open "select ID from FS_MF_Oper_Log order by id asc",Conn,1,3
	Set tmp_admin_rs	= CreateObject(G_FS_RS)
	tmp_admin_rs.Open "select ID from FS_MF_Admin order by id asc",Conn,1,3
	Set tmp_define_rs	= CreateObject(G_FS_RS)
	tmp_define_rs.Open "select DefineID from FS_MF_DefineTable order by DefineID asc",Conn,1,3
  %>
	<tr class="back">
		<td class="xingmu">��Ϣ��¼</td>
	</tr>
	<tr class="back">
		<td height="28" class="hback"><strong><%=session("Admin_Name")%>&nbsp;</strong>���ã���
			<%if Session("Admin_Is_Super") =1 then:response.Write("��ݣ���������Ա"):else:response.Write("��ݣ�һ�����Ա"):end if%>
			�����ǣ�
			<script language="JavaScript" type="text/JavaScript" >var day="";
		var month="";
		var ampm="";
		var ampmhour="";
		var myweekday="";
		var year="";
		mydate=new Date();
		myweekday=mydate.getDay();
		mymonth=mydate.getMonth()+1;
		myday= mydate.getDate();
		myyear= mydate.getYear();
		year=(myyear > 200) ? myyear : 1900 + myyear;
		if(myweekday == 0)
		weekday=" ������ ";
		else if(myweekday == 1)
		weekday=" ����һ ";
		else if(myweekday == 2)
		weekday=" ���ڶ� ";
		else if(myweekday == 3)
		weekday=" ������ ";
		else if(myweekday == 4)
		weekday=" ������ ";
		else if(myweekday == 5)
		weekday=" ������ ";
		else if(myweekday == 6)
		weekday=" ������ ";
		document.write(year+"��"+mymonth+"��"+myday+"�� "+weekday);
	</script>
		</td>
	</tr>
	<tr class="back">
		<td height="27" class="hback">��ݲ˵��� <a href="News/News_add.asp">�������</a> <a href="News/News_manage.asp">����</a>&nbsp;��&nbsp;<a href="News/Class_add.asp?ClassID=&Action=add">���������Ŀ</a> <a href="News/Class_Manage.asp">����</a>&nbsp;��&nbsp;<a href="SysAdmin_list.asp">����Ա����</a>&nbsp;��&nbsp; <a href="Templets_List.asp">ģ�����</a>&nbsp;��&nbsp;<a href="Sys_Oper_Log.asp">��־����</a> </td>
	</tr>
	<tr class="back">
		<td height="35" class="hback">
			<p><a href="News/News_Manage.asp?ClassID=&isCheck=0&Keyword=&ktype=">��������</a>��<span class="tx">
				<% = tmp_news_rs.Recordcount%>
				</span> ƪ ��<a href="News/Constr_Manage.asp">����Ͷ��</a>��<span class="tx"><% = ConStrNum %></span>&nbsp;ƪ <a href="SubSysSet_List.asp">��ϵͳ</a>��&nbsp;<span class="tx">
				<% = tmp_sub_rs.Recordcount%>
				</span>&nbsp;��<a href="SysAdmin_list.asp">����Ա</a>&nbsp;<span class="tx">
				<% = tmp_admin_rs.Recordcount%>
				</span>&nbsp;���� <a href="DefineTable_Manage.asp">�Զ����ֶ�</a>��<span class="tx">
				<% = tmp_define_rs.Recordcount%>
				</span>&nbsp;��,�������<span class="tx"><% = MaxDefineNum %></span>���Զ����ֶ�<br>
				<a href="Sys_Login_Log.asp">��ȫ��־</a>��<span class="tx">
				<% = tmp_Log_login_rs.Recordcount%>
				</span>&nbsp;����<a href="Sys_Oper_Log.asp">������־</a>��<span class="tx">
				<% = tmp_Log_s_rs.Recordcount%>
				</span>&nbsp;��
				<%if tmp_Log_s_rs.Recordcount+tmp_Log_login_rs.Recordcount>1000 then response.Write("��־�Ѿ�����1000�����뼰ʱɾ����")%>
			</p>
		</td>
	</tr>
	<tr class="back">
		<td height="24" class="hback">����ACCESS2000���ݿ��û����뾭��������<a href="DataManage.asp?Type=fix">�ݿ��޸�ѹ��</a>������߳����ִ̨��Ч�ʡ�Ϊ�˰�ȫ��������鶨��<a href="DataManage.asp?Type=bak">�������ݿ�</a>��</td>
	</tr>
	<tr class="back">
		<td height="20" class="hback">����SQL Server 2000 �û����뿪��SQL Server 2000 ������񣬶��ڶ����ݿ���������ݡ�</td>
	</tr>
</table>
<div align="center">
	<%
	tmp_news_rs.close:set tmp_news_rs=nothing
	tmp_sub_rs.close:set tmp_sub_rs=nothing
	tmp_admin_rs.close:set tmp_admin_rs=nothing
	tmp_define_rs.close:set tmp_define_rs=nothing
	tmp_Log_login_rs.close:set tmp_Log_login_rs=nothing
	tmp_Log_s_rs.close:set tmp_Log_s_rs=nothing
	Dim theInstalledObjects(23)
	theInstalledObjects(0) = "MSWC.AdRotator"
	theInstalledObjects(1) = "MSWC.BrowserType"
	theInstalledObjects(2) = "MSWC.NextLink"
	theInstalledObjects(3) = "MSWC.Tools"
	theInstalledObjects(4) = "MSWC.Status"
	theInstalledObjects(5) = "MSWC.Counters"
	theInstalledObjects(6) = "IISSample.ContentRotator"
	theInstalledObjects(7) = "IISSample.PageCounter"
	theInstalledObjects(8) = "MSWC.PermissionChecker"
	theInstalledObjects(9) = G_FS_FSO
	theInstalledObjects(10) = G_FS_CONN
	
	theInstalledObjects(11) = "SoftArtisans.FileUp"
	theInstalledObjects(12) = "SoftArtisans.FileManager"
	theInstalledObjects(13) = "JMail.SMTPMail"
	theInstalledObjects(14) = "CDONTS.NewMail"
	theInstalledObjects(15) = "Persits.MailSender"
	theInstalledObjects(16) = "LyfUpload.UploadFile"
	theInstalledObjects(17) = "Persits.Upload.1"
	theInstalledObjects(18) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
	theInstalledObjects(19)	= "Persits.Jpeg"				'AspJpeg
	theInstalledObjects(20) = "SoftArtisans.ImageGen"		'SoftArtisans ImgWriter V1.21
	theInstalledObjects(21) = "sjCatSoft.Thumbnail"
	theInstalledObjects(22) = "Microsoft.XMLHTTP"
	theInstalledObjects(23) = "Adodb.Stream"
%>
</div>
<table width="98%" height="197" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<tr class="hback">
		<td  colspan="4" class="xingmu">��������Ϣ��������Ϣ���ñ�ϵͳ�Դ���<a href="SystemCheckplus.asp" class="sd"><strong><font color="#FF0000">asp̽��</font></strong></a></td>
	</tr>
	<tr class="hback">
		<td height="32">�����ط���������������IP��ַ<%=Request.ServerVariables("SERVER_NAME")%></td>
		<td height="32">��վ������·����<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
	</tr>
	<tr class="hback">
		<td width="48%" height="32">�����ط�������������Ķ˿ڣ�<%=Request.ServerVariables("SERVER_PORT")%></td>
		<td width="52%" height="32">������������ϵͳ��<%=Request.ServerVariables("OS")%></td>
	</tr>
	<tr class="hback">
		<td width="48%" height="32">���ű���������<span class="small2">��</span><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>��</td>
		<td width="52%" height="37">��WEB�����������ƺͰ汾��<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	</tr>
	<tr class="hback">
		<td width="48%" height="32">���ű���ʱʱ�䣺<%=Server.ScriptTimeout%> ��</td>
		<td width="52%" height="32">��CDONTS���֧��<span class="small2">��</span>
			<%
		On Error Resume Next
		Server.CreateObject(G_CDONTS_NEWMAIL)
		if err=0 then 
			response.write("��")
		else
			response.write("��")
		end if	 
		err=0
	%>
		</td>
	</tr>
	<tr class="hback">
		<td width="48%" height="32">������·����<%=Request.ServerVariables("SCRIPT_NAME")%></td>
		<td width="52%" height="32">��Jmail�������֧�֣�
			<%If Not IsObjInstalled(theInstalledObjects(13)) Then%>
			��
			<%else%>
			��
			<%end if%>
		</td>
	</tr>
</table>
<script language="JavaScript">
<!--
function Get_Foosun_Server(){
	var Userid="";
	GetInfo(String.fromCharCode(83,121,115,67,104,101,99,107,86,101,114,46,97,115,112),String.fromCharCode(65,99,116,61,86,101,114));
	GetInfo(String.fromCharCode(83,121,115,67,104,101,99,107,86,101,114,46,97,115,112),String.fromCharCode(65,99,116,61,78,101,119,115));
}

function GetInfo(url,Action){

	var myAjax = new Ajax.Request(
		url,
		{method:'post',
		parameters:Action,
		onComplete:GetInfo_Receive
		}
		);
}
function GetInfo_Receive(OriginalRequest){
	var Info="";
	var Str_Info="";
	
	Info=OriginalRequest.responseText.split("||");
	if (Info.length>2)
	{
		Str_Info=Info[2];
	}else{
		Str_Info="δ֪����";
	}

	if (Info[1]=="Ver"){
		$('Foosun_server_version').innerHTML=Str_Info;
	}else if (Info[1]=="News"){
		$('Foosun_server_announce').innerHTML=Str_Info;
	}
}
window.onload=Get_Foosun_Server;
//-->
</script>
<table width="98%" height="105" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<tr class="hback">
		<td height="25" colspan="2" class="xingmu">ʹ�ñ��������ȷ�ϵķ�����������������������Ҫ��</td>
	</tr>
	<tr class="hback">
		<td width="48%" height="25">��JRO.JetEngine(ACCESS&nbsp; ���ݿ�<span class="small2">)��</span>
			<%
		On Error Resume Next
		Server.CreateObject(G_JRO_JETENGINE)
		if err=0 then 
		  response.write("��")
		else
		  response.write("��")
		end if	 
		err=0
	 %>
		</td>
		<td width="52%" height="25">�����ݿ�ʹ��:
			<%
		On Error Resume Next
		Server.CreateObject(G_FS_CONN)
		if err=0 then 
		  response.write("��,����ʹ�ñ�ϵͳ")
		else
		  response.write("��,����ʹ�ñ�ϵͳ")
		end if	 
		err=0
	%>
		</td>
	</tr>
	<tr class="hback">
		<td height="25">��<span class="small2">FSO</span>�ı��ļ���д<span class="small2">��</span>
			<%
					On Error Resume Next
					Server.CreateObject(G_FS_FSO)
					if err=0 then 
					  response.write("��,����ʹ�ñ�ϵͳ")
					else
					  response.write("��������ʹ�ô�ϵͳ")
					end if	 
					err=0
				   %>
		</td>
		<td height="25">��Microsoft.XMLHTTP
			<%If Not IsObjInstalled(theInstalledObjects(22)) Then%>
			��
			<%else%>
			��
		  <%end if%>
			(�Ǳ���) ��Adodb.Stream
			<%If Not IsObjInstalled(theInstalledObjects(23)) Then%>
			��
			<%else%>
			��
			<%end if%>
		</td>
	</tr>
	<tr class="hback">
		<td height="25" colspan="2">���ͻ���������汾��
			<%
		    Dim Agent,Browser,version,tmpstr
		    Agent=Request.ServerVariables("HTTP_USER_AGENT")
		    Agent=Split(Agent,";")
		    If InStr(Agent(1),"MSIE")>0 Then
				Browser="MS Internet Explorer "
				version=Trim(Left(Replace(Agent(1),"MSIE",""),6))
			ElseIf InStr(Agent(4),"Netscape")>0 Then 
				Browser="Netscape "
				tmpstr=Split(Agent(4),"/")
				version=tmpstr(UBound(tmpstr))
			ElseIf InStr(Agent(4),"rv:")>0 Then
				Browser="Mozilla "
				tmpstr=Split(Agent(4),":")
				version=tmpstr(UBound(tmpstr))
				If InStr(version,")") > 0 Then 
					tmpstr=Split(version,")")
					version=tmpstr(0)
				End If
			End If
			response.Write(""&Browser&"  "&version&"")
		  %>
			[��ҪIE5.5������,�������������Windows 2000��Windows 2003 Server]</td>
	</tr>
</table>
<table width="98%" height="132" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<tr class="hback">
		<td  colspan="4" class="xingmu">��ϵ����</td>
	</tr>
	<tr class="hback">
		<td height="20">
			<div align="center"> ��Ʒ����</div>		</td>
		<td height="20">�Ĵ���Ѷ�Ƽ���չ���޹�˾</td>
		<td>&nbsp;</td>
	  <td>&nbsp;</td>
  </tr>
	<tr class="hback">
		<td width="13%" height="20">
			<div align="center">�ܻ��绰</div>		</td>
		<td width="30%" height="20">028-85336900 85336900 </td>
<td width="12%">
			<div align="center">��Ʒ��ѯ</div>		</td>
		<td width="45%">028-85336900-601\605\606\607</td>
  </tr>
	<tr class="hback">
		<td width="13%" height="20">
			<div align="center">��������</div>		</td>
		<td width="30%" height="20">028-85336900-603</td>
  <td width="12%">
			<div align="center">�ͷ��绰</div>		</td>
		<td width="45%">028-85336900-608</td>
  </tr>
	<tr class="hback">
		<td width="13%" height="20">
			<div align="center">�ٷ���վ</div>		</td>
		<td width="30%" height="20"><a href="http://www.Foosun.cn">Foosun.cn</a></td>
		<td width="12%">
			<div align="center">��������</div>		</td>
		<td width="45%"><a href="http://Help.foosun.net" target="_blank">Help.foosun.net</a> ����̳��<a href="http://bbs.foosun.net">bbs.foosun.net</a></td>
	</tr>
	<tr class="hback">
		<td height="20" colspan="4">
			<div align="center">&copy;2002-2008 CopyRight 
				Foosun Inc.����All Rights Reserved</div>		</td>
	</tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>