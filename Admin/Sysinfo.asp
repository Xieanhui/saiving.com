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
		<td class="xingmu">欢迎使用风讯网站管理系统(FoosunCMS)V<%=Request.Cookies("FoosunMFCookies")("FoosunMFVersion")%> For ASP Version　　　　版权号：2004SR11453</td>
	</tr>
	<tr class="back">
		<td height="22" class="hback">
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
				<tr>
					<td width="30%">版本号: <%=Request.Cookies("FoosunMFCookies")("FoosunMFVersion")%> 　　</td>
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
		<td class="xingmu">信息记录</td>
	</tr>
	<tr class="back">
		<td height="28" class="hback"><strong><%=session("Admin_Name")%>&nbsp;</strong>您好！　
			<%if Session("Admin_Is_Super") =1 then:response.Write("身份：超级管理员"):else:response.Write("身份：一般管理员"):end if%>
			今天是：
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
		weekday=" 星期日 ";
		else if(myweekday == 1)
		weekday=" 星期一 ";
		else if(myweekday == 2)
		weekday=" 星期二 ";
		else if(myweekday == 3)
		weekday=" 星期三 ";
		else if(myweekday == 4)
		weekday=" 星期四 ";
		else if(myweekday == 5)
		weekday=" 星期五 ";
		else if(myweekday == 6)
		weekday=" 星期六 ";
		document.write(year+"年"+mymonth+"月"+myday+"日 "+weekday);
	</script>
		</td>
	</tr>
	<tr class="back">
		<td height="27" class="hback">快捷菜单： <a href="News/News_add.asp">添加新闻</a> <a href="News/News_manage.asp">管理</a>&nbsp;｜&nbsp;<a href="News/Class_add.asp?ClassID=&Action=add">添加新闻栏目</a> <a href="News/Class_Manage.asp">管理</a>&nbsp;｜&nbsp;<a href="SysAdmin_list.asp">管理员管理</a>&nbsp;｜&nbsp; <a href="Templets_List.asp">模板管理</a>&nbsp;｜&nbsp;<a href="Sys_Oper_Log.asp">日志管理</a> </td>
	</tr>
	<tr class="back">
		<td height="35" class="hback">
			<p><a href="News/News_Manage.asp?ClassID=&isCheck=0&Keyword=&ktype=">待审新闻</a>：<span class="tx">
				<% = tmp_news_rs.Recordcount%>
				</span> 篇 　<a href="News/Constr_Manage.asp">待审投稿</a>：<span class="tx"><% = ConStrNum %></span>&nbsp;篇 <a href="SubSysSet_List.asp">子系统</a>：&nbsp;<span class="tx">
				<% = tmp_sub_rs.Recordcount%>
				</span>&nbsp;，<a href="SysAdmin_list.asp">管理员</a>&nbsp;<span class="tx">
				<% = tmp_admin_rs.Recordcount%>
				</span>&nbsp;个。 <a href="DefineTable_Manage.asp">自定义字段</a>：<span class="tx">
				<% = tmp_define_rs.Recordcount%>
				</span>&nbsp;个,最多允许<span class="tx"><% = MaxDefineNum %></span>个自定义字段<br>
				<a href="Sys_Login_Log.asp">安全日志</a>：<span class="tx">
				<% = tmp_Log_login_rs.Recordcount%>
				</span>&nbsp;个，<a href="Sys_Oper_Log.asp">操作日志</a>：<span class="tx">
				<% = tmp_Log_s_rs.Recordcount%>
				</span>&nbsp;个
				<%if tmp_Log_s_rs.Recordcount+tmp_Log_login_rs.Recordcount>1000 then response.Write("日志已经超过1000个，请及时删除。")%>
			</p>
		</td>
	</tr>
	<tr class="back">
		<td height="24" class="hback">对于ACCESS2000数据库用户：请经常进行数<a href="DataManage.asp?Type=fix">据库修复压缩</a>，以提高程序后台执行效率。为了安全起见，建议定期<a href="DataManage.asp?Type=bak">备份数据库</a>。</td>
	</tr>
	<tr class="back">
		<td height="20" class="hback">对于SQL Server 2000 用户：请开启SQL Server 2000 代理服务，定期对数据库服务器备份。</td>
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
		<td  colspan="4" class="xingmu">服务器信息　更多信息请用本系统自带的<a href="SystemCheckplus.asp" class="sd"><strong><font color="#FF0000">asp探针</font></strong></a></td>
	</tr>
	<tr class="hback">
		<td height="32">　返回服务器的主机名，IP地址<%=Request.ServerVariables("SERVER_NAME")%></td>
		<td height="32">　站点物理路径：<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
	</tr>
	<tr class="hback">
		<td width="48%" height="32">　返回服务器处理请求的端口：<%=Request.ServerVariables("SERVER_PORT")%></td>
		<td width="52%" height="32">　服务器操作系统：<%=Request.ServerVariables("OS")%></td>
	</tr>
	<tr class="hback">
		<td width="48%" height="32">　脚本解释引擎<span class="small2">：</span><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>　</td>
		<td width="52%" height="37">　WEB服务器的名称和版本：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	</tr>
	<tr class="hback">
		<td width="48%" height="32">　脚本超时时间：<%=Server.ScriptTimeout%> 秒</td>
		<td width="52%" height="32">　CDONTS组件支持<span class="small2">：</span>
			<%
		On Error Resume Next
		Server.CreateObject(G_CDONTS_NEWMAIL)
		if err=0 then 
			response.write("√")
		else
			response.write("×")
		end if	 
		err=0
	%>
		</td>
	</tr>
	<tr class="hback">
		<td width="48%" height="32">　虚拟路径：<%=Request.ServerVariables("SCRIPT_NAME")%></td>
		<td width="52%" height="32">　Jmail邮箱组件支持：
			<%If Not IsObjInstalled(theInstalledObjects(13)) Then%>
			×
			<%else%>
			√
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
		Str_Info="未知错误";
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
		<td height="25" colspan="2" class="xingmu">使用本软件，请确认的服务器和你的浏览器满足以下要求：</td>
	</tr>
	<tr class="hback">
		<td width="48%" height="25">　JRO.JetEngine(ACCESS&nbsp; 数据库<span class="small2">)：</span>
			<%
		On Error Resume Next
		Server.CreateObject(G_JRO_JETENGINE)
		if err=0 then 
		  response.write("√")
		else
		  response.write("×")
		end if	 
		err=0
	 %>
		</td>
		<td width="52%" height="25">　数据库使用:
			<%
		On Error Resume Next
		Server.CreateObject(G_FS_CONN)
		if err=0 then 
		  response.write("√,可以使用本系统")
		else
		  response.write("×,不能使用本系统")
		end if	 
		err=0
	%>
		</td>
	</tr>
	<tr class="hback">
		<td height="25">　<span class="small2">FSO</span>文本文件读写<span class="small2">：</span>
			<%
					On Error Resume Next
					Server.CreateObject(G_FS_FSO)
					if err=0 then 
					  response.write("√,可以使用本系统")
					else
					  response.write("×，不能使用此系统")
					end if	 
					err=0
				   %>
		</td>
		<td height="25">　Microsoft.XMLHTTP
			<%If Not IsObjInstalled(theInstalledObjects(22)) Then%>
			×
			<%else%>
			√
		  <%end if%>
			(非必须) 　Adodb.Stream
			<%If Not IsObjInstalled(theInstalledObjects(23)) Then%>
			×
			<%else%>
			√
			<%end if%>
		</td>
	</tr>
	<tr class="hback">
		<td height="25" colspan="2">　客户端浏览器版本：
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
			[需要IE5.5或以上,服务器建议采用Windows 2000或Windows 2003 Server]</td>
	</tr>
</table>
<table width="98%" height="132" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<tr class="hback">
		<td  colspan="4" class="xingmu">联系我们</td>
	</tr>
	<tr class="hback">
		<td height="20">
			<div align="center"> 产品开发</div>		</td>
		<td height="20">四川风讯科技发展有限公司</td>
		<td>&nbsp;</td>
	  <td>&nbsp;</td>
  </tr>
	<tr class="hback">
		<td width="13%" height="20">
			<div align="center">总机电话</div>		</td>
		<td width="30%" height="20">028-85336900 85336900 </td>
<td width="12%">
			<div align="center">产品咨询</div>		</td>
		<td width="45%">028-85336900-601\605\606\607</td>
  </tr>
	<tr class="hback">
		<td width="13%" height="20">
			<div align="center">传　　真</div>		</td>
		<td width="30%" height="20">028-85336900-603</td>
  <td width="12%">
			<div align="center">客服电话</div>		</td>
		<td width="45%">028-85336900-608</td>
  </tr>
	<tr class="hback">
		<td width="13%" height="20">
			<div align="center">官方网站</div>		</td>
		<td width="30%" height="20"><a href="http://www.Foosun.cn">Foosun.cn</a></td>
		<td width="12%">
			<div align="center">帮助中心</div>		</td>
		<td width="45%"><a href="http://Help.foosun.net" target="_blank">Help.foosun.net</a> 　论坛：<a href="http://bbs.foosun.net">bbs.foosun.net</a></td>
	</tr>
	<tr class="hback">
		<td height="20" colspan="4">
			<div align="center">&copy;2002-2008 CopyRight 
				Foosun Inc.　　All Rights Reserved</div>		</td>
	</tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>