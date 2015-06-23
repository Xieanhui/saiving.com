<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_checksysplus.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,strShowErr
MF_Default_Conn
MF_Session_TF
Dim startime
	 startime=timer()
Dim hx
Set hx = New cls_ActiveXCheck
Dim theObj(25,1)

	theObj(0,0) = "MSWC.AdRotator"
	theObj(1,0) = "MSWC.BrowserType"
	theObj(2,0) = "MSWC.NextLink"
	theObj(3,0) = "MSWC.Tools"
	theObj(4,0) = "MSWC.Status"
	theObj(5,0) = "MSWC.Counters"
	theObj(6,0) = "MSWC.PermissionChecker"
	theObj(7,0) = "WScript.Shell"
	theObj(8,0) = "Microsoft.XMLHTTP"
	theObj(9,0) = "Scripting.FileSystemObject"
	theObj(9,1) = "(FSO 文本文件读写)"
	theObj(10,0) = "ADODB.Connection"
	theObj(10,1) = "(ADO 数据对象)"
    
	theObj(11,0) = "SoftArtisans.FileUp"
	theObj(11,1) = "(SA-FileUp 文件上传)"
	theObj(12,0) = "SoftArtisans.FileManager"
	theObj(12,1) = "(SoftArtisans 文件管理)"
	theObj(13,0) = "LyfUpload.UploadFile"
	theObj(13,1) = "(Lyf的文件上传组件)"
	theObj(14,0) = "Persits.Upload"
	theObj(14,1) = "(ASPUpload 文件上传)"
	theObj(15,0) = "w3.upload"
	theObj(15,1) = "(Dimac 文件上传)"

	theObj(16,0) = "JMail.SmtpMail"
	theObj(16,1) = "(Dimac JMail 邮件收发)</a>"
	theObj(17,0) = "CDONTS.NewMail"
	theObj(17,1) = "(虚拟 SMTP 发信)"
	theObj(18,0) = "Persits.MailSender"
	theObj(18,1) = "(ASPemail 发信)"
	theObj(19,0) = "SMTPsvg.Mailer"
	theObj(19,1) = "(ASPmail 发信)"
	theObj(20,0) = "DkQmail.Qmail"
	theObj(20,1) = "(dkQmail 发信)"
	theObj(21,0) = "Geocel.Mailer"
	theObj(21,1) = "(Geocel 发信)"
	theObj(22,0) = "IISmail.Iismail.1"
	theObj(22,1) = "(IISmail 发信)"
	theObj(23,0) = "SmtpMail.SmtpMail.1"
	theObj(23,1) = "(SmtpMail 发信)"
	theObj(24,0) = "SoftArtisans.ImageGen"
	theObj(24,1) = "(SA 的图像读写组件)"
	theObj(25,0) = "W3Image.Image"
	theObj(25,1) = "(Dimac 的图像读写组件)"
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE></TITLE>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style>.PicBar { background-color: #0099CF; border: 1px solid #000000; height: 12px;}</style>
<SCRIPT language="JavaScript" runat="server">
	function getEngVerJs(){
		try{
			return ScriptEngineMajorVersion() +"."+ScriptEngineMinorVersion()+"."+ ScriptEngineBuildVersion() + " ";
		}catch(e){
			return "服务器不支持此项检测";
		}
		
	}
</SCRIPT>
<SCRIPT language="VBScript" runat="server">
	Function getEngVerVBS()
		getEngVerVBS=ScriptEngineMajorVersion() &"."&ScriptEngineMinorVersion() &"." & ScriptEngineBuildVersion() & " "
	End Function
</SCRIPT>
<script language="javascript">
<!--
	function Checksearchbox()
	{
	if(form1.classname.value == "")
	{
		alert("请输入你要检测的组件名！");
		form1.classname.focus();
		return false;
	}
	}
	function showsubmenu(sid)
	{
	whichEl = eval("submenu" + sid);
	if (whichEl.style.display == "none")
	{
	eval("submenu" + sid + ".style.display=\"\";");
	eval("txt" + sid + ".innerHTML=\"<a href='#' title='关闭此项'><font face='Wingdings' >x</font></a>\";");
	}
	else
	{
	eval("submenu" + sid + ".style.display=\"none\";");
	eval("txt" + sid + ".innerHTML=\"<a href='#' title='打开此项'><font face='Wingdings' >y</font></a>\";");
	}
	}
-->
</SCRIPT>
</HEAD>
<BODY leftmargin="50">
<a name=top></a>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td align="center" class="xingmu"><p align="left">ASP探针 </td>
  </tr>
  <tr>
    <td align="center" class="hback"> <div align="left">
        <%
		dim action
		action=request("action")
		if action="testzujian" then
		call ObjTest2
		end if
		
		Call menu
		Call SystemTest
		Call ObjTest
		Call CalculateTest
		Call DriveTest
		Call SpeedTest
		hx.ShowFooter
		Set hx= nothing
	%>
        <%Sub menu%>
        选项：<a href="#SystemTest">服务器有关参数</a> | <a href="#ObjTest">服务器组件情况</a> 
        | <a href="#CalcuateTest">服务器运算能力</a> | <a href="#DriveTest">服务器磁盘信息</a> 
        | <a href="#SpeedTest">服务器连接速度</a> 
        <%End Sub%>
        <%Sub smenu(i)%>
        <a href="#top" title="返回顶部"><font face='Webdings'>5</font></a> <span id=txt<%=i%> name=txt<%=i%>><a href='#' title='关闭此项'><font face='Wingdings'>x</font></a></span> 
        <%End Sub%>
        <%Sub SystemTest
			on error resume next
		%>
      </div></td>
  </tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td height="25" align="center"  class="xingmu" onClick="showsubmenu(0)"><div align="left"><strong>服务器有关参数</strong> 
        <%Call smenu(0)%>
        <a name="SystemTest"></a></div></td>
  </tr>
  <tr class="hback"> 
    <td style="display" id='submenu0'> <table width=100% border=0 align="center" cellpadding=5 cellspacing=1>
        <tr height=18> 
          <td width="130">&nbsp;服务器名</td>
          <td width="170" height="18">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
          <td width="130" height="18">&nbsp;服务器操作系统</td>
          <td width="170" height="18">&nbsp;<%=Request.ServerVariables("OS")%></td>
        </tr>
        <tr height=18> 
          <td>&nbsp;服务器IP</td>
          <td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
          <td>&nbsp;服务器端口</td>
          <td>&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
        </tr>
        <tr height=18> 
          <td>&nbsp;服务器时间</td>
          <td>&nbsp;<%=now%></td>
          <td>&nbsp;服务器CPU数量</td>
          <td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
        </tr>
        <tr height=18> 
          <td>&nbsp;IIS版本</td>
          <td height="18">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
          <td height="18">&nbsp;脚本超时时间</td>
          <td height="18">&nbsp;<%=Server.ScriptTimeout%> 秒</td>
        </tr>
        <tr height=18> 
          <td>&nbsp;Application变量</td>
          <td height="18">&nbsp; 
            <%Response.Write(Application.Contents.Count & "个 ")
		  if Application.Contents.count>0 then Response.Write("[<a href=""?action=showapp"">遍历Application变量</a>]")%>
          </td>
          <td height="18">&nbsp;Session变量<br> </td>
          <td height="18">&nbsp; 
            <%Response.Write(Session.Contents.Count&"个 ")
		  if Session.Contents.count>0 then Response.Write("[<a href=""?action=showsession"">遍历Session变量</a>]")%>
          </td>
        </tr>
        <tr height=18> 
          <td height="18">&nbsp;<a href="?action=showvariables">所有服务器参数</a></td>
          <td height="18">&nbsp; 
            <%Response.Write(Request.ServerVariables.Count&"个 ")
		  if Request.ServerVariables.Count>0 then Response.Write("[<a href=""?action=showvariables"">遍历服务器参数</a>]")%>
          </td>
          <td height="18">&nbsp;服务器环境变量</td>
          <td height="18">&nbsp; 
            <%
			dim WshShell,WshSysEnv
			Set WshShell = server.CreateObject(G_WSCRIPT_SHELL)
			Set WshSysEnv = WshShell.Environment
			if err then
				Response.Write("服务器不支持WScript.Shell组件")
				err.clear
			else
				Response.Write(WshSysEnv.count &"个 ")
				if WshSysEnv.count>0 then Response.Write("[<a href=""?action=showwsh"">遍历环境变量</a>]") 
		 	end if
		  %>
          </td>
        </tr>
        <tr height=18> 
          <td align=left>&nbsp;服务器解译引擎</td>
          <td height="18" colspan="3">&nbsp;JScript: <%= getEngVerJs() %> | VBScript: 
            <%=getEngVerVBS()%></td>
        </tr>
        <tr height=18> 
          <td align=left>&nbsp;本文件实际路径</td>
          <td height="8" colspan="3">&nbsp;<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
        </tr>
      </table>
      <%
if action="showapp" or action="showsession" or action="showvariables" or action="showwsh" then
	showvariable(action)
end if
%>
    </td>
  </tr>
</table>
<br>
<%
End Sub

Sub showvariable(action)
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <tr> 
    <td colspan="2" class="xingmu"><%
	on error resume next
	dim Item,xTestObj,outstr
	if action="showapp" then
		Response.Write("<font face='Webdings'>4</font> 遍历Application变量")
		set xTestObj=Application.Contents
	elseif action="showsession" then
		Response.Write("<font face='Webdings'>4</font> 遍历Session变量")
		set xTestObj=Session.Contents
	elseif action="showvariables" then
		Response.Write("<font face='Webdings'>4</font> 遍历服务器参数")
		set xTestObj=Request.ServerVariables
	elseif action="showwsh" then
		Response.Write("<font face='Webdings'>4</font> 遍历环境变量")
		dim WshShell
		Set WshShell = server.CreateObject(G_WSCRIPT_SHELL)
		set xTestObj=WshShell.Environment
	end if
		Response.Write "(<a href="&hx.FileName&">关闭</a>)"
	%>
    </td>
  </tr>
  <tr class="hback"> 
    <td width="130">变量名</td>
    <td width="470">值</td>
  </tr>
  <%
	if err then
		outstr = "<tr bgcolor=#FFFFFF><td colspan=2>没有符合条件的变量</td></tr>"
		err.clear
	else
		dim w
		if action="showwsh" then
			for each Item in xTestObj
				w=split(Item,"=")
				outstr = outstr & "<tr bgcolor=#FFFFFF>"
				outstr = outstr & "<td>" & w(0) & "</td>" 
				outstr = outstr & "<td>" & w(1) & "</td>" 
				outstr = outstr & "</tr>" 		
			next
		else
			dim i
			for each Item in xTestObj
				outstr = outstr & "<tr bgcolor=#FFFFFF>"
				outstr = outstr & "<td>" & Item & "</td>" 				
				outstr = outstr & "<td>"
				if IsArray(xTestObj(Item)) then		
					for i=0 to ubound(xTestObj(Item))-1
					outstr = outstr & hx.formatvariables(xTestObj(Item)(i)) & "<br>"
					next
				else
					outstr = outstr & hx.formatvariables(xTestObj(Item))
				end if			
				outstr = outstr & "</td>"
				outstr = outstr & "</tr>" 
			next
		end if
	end if
		Response.Write(outstr)	
		set xTestObj=nothing
		%>
</table>
<%End Sub%>
<%Sub ObjTest%>
<a name="ObjTest"></a>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td height="25" align="center" onClick="showsubmenu(1)" class="xingmu"><div align="left">服务器组件情况
        <%Call smenu(1)%>
      </div></td>
  </tr>
  <tr> 
    <td style="display" id='submenu1'><table border=0 width=100% cellspacing=1 cellpadding=3 class="table">
        <tr height=18> 
          <td colspan="2" class="xingmu">IIS自带的ASP组件</td>
        </tr>
        <tr  height=18 class="hback"> 
          <td width=450 align="center">组 件 名 称</td>
          <td width=150 align="center">支持及版本</td>
        </tr>
        <%hx.GetObjInfo 0,10%>
      </table>
      <table border=0 width=100% cellspacing=1 cellpadding=3  class="table">
        <tr class="xingmu" height=18> 
          <td colspan="2"  class="xingmu">常见的文件上传和管理组件 </td>
        </tr>
        <tr  class="hback" height=18> 
          <td width=450 align="center">组 件 名 称</td>
          <td width=150 align="center">支持及版本</td>
        </tr>
        <%hx.GetObjInfo 11,15%>
      </table>
      <table border=0 width=100% cellspacing=1 cellpadding=3 class="table">
        <tr  height=18> 
          <td colspan="2" class="xingmu">常见的收发邮件组件</td>
        </tr>
        <tr class="hback" height=18> 
          <td width=450 align="center">组 件 名 称</td>
          <td width=150 align="center">支持及版本</td>
        </tr>
        <%hx.GetObjInfo 16,23%>
      </table>
      <table border=0 width=100% cellspacing=1 cellpadding=3 class="table">
        <tr  height=18> 
          <td colspan="2" class="xingmu">图像处理组件</td>
        </tr>
        <tr class="hback" height=18> 
          <td width=450 align="center">组 件 名 称</td>
          <td width=150 align="center">支持及版本</td>
        </tr>
        <%hx.GetObjInfo 24,25%>
      </table>
      <table border=0 width=100% cellspacing=1 cellpadding=3  class="table">
        <tr> 
          <td  class="xingmu">其他组件支持情况检测 </td>
        </tr>
        <FORM action=?action=testzujian method=post id=form1 name=form1 onSubmit="JavaScript:return Checksearchbox();">
          <tr> 
            <td height=30 class="hback">输入你要检测的组件的ProgId或ClassId 
              <input class=input type=text value="" name="classname" size=40> 
              <INPUT type=submit value="确定" class=backc id=submit1 name=submit1> 
            </td></tr>
        </FORM>
      </table></td>
  </tr>
</table>
<br>
<%
End Sub
Sub ObjTest2
	Dim strClass
    strClass = Trim(Request.Form("classname"))
    If strClass <> "" then
    Response.Write "<br>您指定的组件的检查结果："
      If Not hx.IsObjInstalled(strClass) then 
        Response.Write "<br><font color=red>很遗憾，该服务器不支持" & strclass & "组件！</font>"
      Else
        Response.Write "<br><font color=green>"
		Response.Write " 恭喜！该服务器支持" & strclass & "组件。"
		If hx.getver(strclass)<>"" then
		Response.Write " 该组件版本是：" & hx.getver(strclass)
		End if
		Response.Write "</font>"
      End If
      Response.Write "<br>"
    end if
	
	Response.Write "<p><a href="&hx.FileName&">返回</a></p>"
	Response.End
End Sub
Sub CalculateTest
%><a name="CalcuateTest"></a>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr> 
    <td height="24" align="center" onClick="showsubmenu(2)" class="xingmu"><div align="left">服务器运算能力 
        <%Call smenu(2)%>
      </div></td>
  </tr>
  <tr> 
    <td style="display" id='submenu2' class="hback"> 
      <table border=0 width=100% cellspacing=1 cellpadding=3>
        <tr height=18> 
          <td colspan="3">让服务器执行50万次加法(整数运算)和20万次开方(浮点运算)，记录其所使用的时间。 </td>
        </tr>
        <%
	dim i,t1,t2,tempvalue,runtime1,runtime2
	'开始计算50万次加法所需时间
	t1=timer()
	for i=1 to 500000
		tempvalue= 1 + 1
	next
	t2=timer()
	runtime1=formatnumber((t2-t1)*1000,2)
	
	'开始计算20万次开方所需时间
	t1=timer()
	for i=1 to 200000
		tempvalue= 2^0.5
	next
	t2=timer()
	runtime2=formatnumber((t2-t1)*1000,2)
%>
        <tr height=25 class="hback"> 
          <td width="400" align=left >&nbsp;<font color=red>您正在使用的这台服务器</font>&nbsp; 
            <INPUT name="button" type="button" class=backc onClick="document.location.href='<%=hx.FileName%>'" value="重新测试"> 
          </td>
          <td width="100" >&nbsp;<font color=red><%=runtime1%> 毫秒</font></td>
          <td width="100" >&nbsp;<font color=red><%=runtime2%> 毫秒</font></td>
        </tr>
      </table></td></tr>
</table>
<br>
<%
End Sub
Sub DriveTest
	On Error Resume Next
	Dim fo,d,xTestObj
	set fo=Server.Createobject(G_FS_FSO)
	set xTestObj=fo.Drives
%>
<a name="DriveTest"></a>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
    <td height="25" align="center" onClick="showsubmenu(4)"><div align="left"><font color="#FFFFFF"></font> 
        <%Call smenu(4)%>
      </div></td>
  </tr>
  <tr>
    <td style="display" id='submenu4'> 
		<%if hx.IsObjInstalled(G_FS_FSO) then%>
      <table border=0 width=100% cellspacing=1 cellpadding=3 class="table">
        <tr height=18> 
          <td colspan="7" class="xingmu">服务器磁盘信息</td>
        </tr>
        <tr height=18 align=center> 
          <td width="90"  class="xingmu">磁盘类型</td>
          <td width="40"  class="xingmu">盘符</td>
          <td width="30"  class="xingmu">可用</td>
          <td width="100"  class="xingmu">卷标</td>
          <td width="80"  class="xingmu">文件系统</td>
          <td width="130"  class="xingmu">可用空间</td>
		  <td width="130"  class="xingmu">总空间</td>
        </tr>
		<%
	for each d in xTestObj
		Response.write "<tr height=18 class=""hback"">"
		Response.write "<td>&nbsp;"&hx.dtype(d.DriveType)&"</td>"
		Response.write "<td>&nbsp;"&d.DriveLetter&"</td>"		
		if d.DriveLetter = "A" then
		Response.Write "<td colspan=6>&nbsp;为防止影响服务器，不检查软驱</td>"
		else
		Response.write "<td align=center> "
		if d.isready then
			Response.Write "√"
			Response.write "</td>"
			Response.write "<td>&nbsp;"&d.VolumeName&"</td>"
			Response.write "</td>"		
			Response.write "<td>&nbsp;"&d.FileSystem&"</td>"	
			Response.write "<td>&nbsp;"&hx.formatdsize(d.FreeSpace)&"</td>"
			Response.write "<td>&nbsp;"&hx.formatdsize(d.TotalSize)&"</td>"
		else
			Response.Write "×"
			Response.Write "<td colspan=4>&nbsp;可能是磁盘有问题，或者程序没有读取权限</td>"			
		end if			 
		end if		 
	next%>
      </table>
	  <%
	Dim filePath,fileDir,fileDrive
	filePath = server.MapPath(".")
	set fileDir = fo.GetFolder(filePath)
	set fileDrive = fo.GetDrive(fileDir.Drive)
	  %>
      <table border=0 width=100% cellspacing=1 cellpadding=3 class="table">
        <tr height=18> 
          <td colspan="6" class="xingmu">当前文件夹信息 (<%=filePath%>)</td>
        </tr>
        <tr height=18 align=center class="xingmu"> 
          <td width="100"  class="xingmu">已用空间</td>
          <td width="100"  class="xingmu">可用空间</td>
          <td width="70"  class="xingmu">文件夹数</td>
          <td width="70"  class="xingmu">文件数</td>
          <td width="130"  class="xingmu">创建时间</td>
          <td width="130"  class="xingmu">修改时间</td>
        </tr>
        <%
		Response.write "<tr height=18  align=center  class=""hback""> "
		Response.write "<td>"&hx.formatdsize(fileDir.Size)&"</td>"
		Response.write "<td>"
		Response.write hx.formatdsize(fileDrive.AvailableSpace)
		if err then
		Response.write "没有权限读取"
		error.clear
		end if
		Response.write "</td>"
		Response.write "<td>"&fileDir.SubFolders.Count&"</td>"
		Response.write "<td>"&fileDir.Files.Count&"</td>"						
		Response.write "<td>"&fileDir.DateCreated&"</td> "
		Response.write "<td>"&fileDir.DateLastModified&"</td> "
	
	Dim i,t1,t2,runtime,TestFileName
	Dim tempfo
	t1= timer()
	TestFileName=server.mappath("Test.txt")
	for i=1 to 30
	set tempfo=fo.CreateTextFile(TestFileName,true)
		tempfo.WriteLine "It's a test file."
	set tempfo=nothing
	set tempfo=fo.OpenTextFile(TestFileName,8,0)
		tempfo.WriteLine "It's a test file."
	set tempfo=nothing
	set tempfo=fo.GetFile(TestFileName)
		tempfo.delete True
	set tempfo=nothing	
	next
	t2=	timer()
	runtime=formatnumber((t2-t1)*1000,2)		 
%>
      </table>
      <table border=0 width=100% cellspacing=1 cellpadding=3 class="table">
        <tr height=18> 
          <td colspan="2" class="xingmu"> 磁盘文件操作速度测试 (重复创建、写入、追加和删除文本文件30次，记录其所使用的时间)</td>
        </tr>
        <tr height=25 class="hback"> 
          <td width="400" align=left ><span class="tx">您正在使用的这台服务器</span> <INPUT name="button2" type="button" class=backc onClick="document.location.href='<%=hx.FileName%>'" value="重新测试"> 
          </td>
          <td width="200" >&nbsp;<font color=red><%=runtime%> 毫秒</font></td>
        </tr>
      </table>
      <%
	  	else
		Response.write "&nbsp;您的服务器或租用的空间不支持FSO组件，无法进行此项测试!"
	end if%>
	  </td>
  </tr>
</table>
<%
End Sub
Sub SpeedTest
Response.Flush()
%>
<a name="SpeedTest"></a>
<%	if action="SpeedTest" then%>
<div id="testspeed"> 
  <table width="200" border="0" cellspacing="0" cellpadding="0" class="divcenter">
    <tr> 
      <td height="30" align=center><p><font color="#000000"><span id=txt1>网速测试中，请稍候...</span></font></p></td>
    </tr>
  </table>
</div>
<%	end if%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr> 
    <td height="25" align="center" class="xingmu" onClick="showsubmenu(3)"><div align="left">服务器连接速度
        <%smenu(3)%>
      </div></td>
  </tr>
  <tr class="hback"> 
    <td bgcolor="#F8F9FC" style="display" id='submenu3'> 
	<table width="100%" border="0" cellspacing=1 cellpadding=3 class="table">
        <tr class="hback"> 
          <td width="80">接入设备</td>
          <td width="420">&nbsp;连接速度(理想值)</td>
          <td width="100">下载速度(理想值)</td>
        </tr>
<tr class="hback"> 
          <td>56k Modem</td>
          <td><img align=absmiddle class=PicBar width='1%'> 56 Kbps</td><td>&nbsp;7.0 k/s</td>
        </tr>
        <tr class="hback"> 
          <td>64k ISDN</td>
          <td><img align=absmiddle class=PicBar width='1%'> 64 Kbps</td><td>&nbsp;8.0 k/s</td>
        </tr>
        <tr class="hback"> 
          <td>512k ADSL</td>
          <td><img align=absmiddle class=PicBar width='5%'> 512 Kbps</td><td>&nbsp;64.0 k/s</td>
        </tr>
        <tr class="hback"> 
          <td height="19">1.5M Cable</td>
          <td><img align=absmiddle class=PicBar width='15%'> 1500 Kbps</td><td>&nbsp;187.5 k/s</td>
        </tr>
        <tr class="hback"> 
          <td>5M FTTP</td>
          <td><img width='50%' align=absmiddle class=PicBar style="background-color: #666633"> 
            5000 Kbps</td>
          <td>&nbsp;625.0 k/s</td>
        </tr>
        <tr class="hback"> 
          <td>当前连接速度</td>
          <%
	if action="SpeedTest" then
		dim i
		Response.Write("<script language=""JavaScript"">var tSpeedStart=new Date();</script>")	
		Response.Write("<!--") & chr(13) & chr(10)
		for i=1 to 1000
		Response.Write("####################################################################################################") & chr(13) & chr(10)
		next
		Response.Write("-->") & chr(13) & chr(10)
		Response.Write("<script language=""JavaScript"">var tSpeedEnd=new Date();</script>") & chr(13) & chr(10)		
		Response.Write("<script language=""JavaScript"">")
		Response.Write("var iSpeedTime=0;iSpeedTime=(tSpeedEnd - tSpeedStart) / 1000;")
		Response.Write("if(iSpeedTime>0) iKbps=Math.round(Math.round(100 * 8 / iSpeedTime * 10.5) / 10); else iKbps=10000 ;")
		Response.Write("var iShowPer=Math.round(iKbps / 100);")		
		Response.Write("if(iShowPer<1) iShowPer=1;  else if(iShowPer>82)   iShowPer=82;")
		Response.Write("</script>") & chr(13) & chr(10)		
		Response.Write("<script language=""JavaScript"">") 
		Response.Write("document.write('<td><img align=absmiddle class=PicBar width=""' + iShowPer + '%""> ' + iKbps + ' Kbps');")
		Response.Write("</script>") & chr(13) & chr(10)
		Response.Write("</td><td>&nbsp;<a href='?action=SpeedTest' title=测试连接速度><u>")
		Response.Write("<script language=""JavaScript"">")
		Response.Write("document.write(Math.round(iKbps/8*10)/10+ ' k/s');")		
		Response.Write("</script>") & chr(13) & chr(10)				
		Response.Write 	"</u></a></td>"
%>
<script>
txt1.innerHTML="网速测试完毕!"
testspeed.style.visibility="hidden"
</script>
<%
	else
		Response.Write "<td></td><td>&nbsp;<a href='?action=SpeedTest' title=测试连接速度><u>开始测试</u></a></td>"
	end if
%>
        </tr>
      </table></td>
  </tr>
</table>
<%End Sub%>
</BODY>
</HTML>






