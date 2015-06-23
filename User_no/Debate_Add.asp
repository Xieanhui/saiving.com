<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim GroupRs,DebateRs,UserGroupRs,DebateID,ClassID,GroupName,CroupContent,InfoType,ClassType,AddTime,GroupCreater,GroupManager,isSys,TempRs,ForIndex,GroupManagerArray,selectedTF,AppointUserGroup,AppointUserNumber,action,Debate_Title,Debate_Content,Debate_ID,GroupMembers,GroupID,AccessFile
ClassID=CintStr(Request.QueryString("classid"))
DebateID=CintStr(Request.QueryString("DebateID"))
action=NoSqlHack(Request.QueryString("act"))
Set GroupRs=Server.CreateObject(G_FS_RS)
Set DebateRs=Server.CreateObject(G_FS_RS)
'获得用户文件保存路径
Dim str_CurrPath
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")

GroupRs.open "select gdID,title,content,InfoType,ClassType,AddTime,PerPageNum,UserNumber,AdminName,ClassMember,isSys,hits,isLock from FS_ME_GroupDebateManage where gdID="&CintStr(ClassID),User_Conn,1,3
If Not GroupRs.eof then
	GroupMembers=GroupRs("ClassMember")
	if Instr(GroupMembers,session("FS_UserNumber"))=0 then
			Response.Redirect("lib/error.asp?ErrCodes=<li>请先加入该社群</li>")
			Response.End()
	end If
	on error resume next
	if GroupRs("isLock")=1 then
			Response.Redirect("lib/error.asp?ErrCodes=<li>该群已经被锁定</li>")
			Response.End()
	end If
End if

if DebateID<>"" then
	DebateRs.open "select DebateID,title,content,ParentID,ClassID,UserNumber,AddTime,AppointUserNumber,AppointUserGroup,AddIP,AccessFile from FS_ME_GroupDebate where UserNumber = '" & Fs_User.UserNumber & "' And DebateID="&CintStr(DebateID),User_Conn,1,3
else
	DebateRs.open "select DebateID,title,content,ParentID,ClassID,UserNumber,AddTime,AppointUserNumber,AppointUserGroup,AddIP,AccessFile from FS_ME_GroupDebate where ClassID="&CintStr(ClassID),User_Conn,1,3
end if

if action="addaction" then 
	DebateRs.addNew
	DebateRs("title")=NoSqlHack(Trim(Request.Form("title")))
	DebateRs("Content")=NoHtmlHackInput(Request.Form("Content"))
	DebateRs("AppointUserNumber")=NoSqlHack(Request.Form("AppointUserNumber"))
	if trim(Request.Form("userGroup"))<>"" or trim(Request.Form("corpGroup"))<>"" then
		DebateRs("AppointUserGroup")=NoSqlHack(Request.Form("userGroup"))&","&NoSqlHack(Request.Form("corpGroup"))
	End if
	DebateRs("ClassID")=NoSqlHack(ClassID)
	DebateRs("ParentID")=0
	DebateRs("AddTime")=Now()
	DebateRs("UserNumber")=session("FS_UserNumber")
	DebateRs("AddIP")=NoSqlHack(CheckIpSafe(request.ServerVariables("REMOTE_ADDR")))
	If Trim(request.Form("file"))<>"" then
		DebateRs("AccessFile")=NoSqlHack(Trim(request.Form("file")))
	End if
	DebateRs.update
	DebateRs.close
	if err.number=0 then 
		Response.Redirect("lib/success.asp?ErrCodes=<li>添加成功</li>&ErrorURL=../Group_unit.asp?GDID="&ClassID)
		Response.End()
	else
		Response.Redirect("lib/error.asp?ErrCodes=<li>"&err.description&"</li>")
		Response.End()
	end if
elseif action="edit" then
	Debate_ID=DebateRs("DebateID")
	Debate_Title=DebateRs("title")
	Debate_Content=DebateRs("content")
	AppointUserNumber=DebateRs("AppointUserNumber")
	If Not IsNull(DebateRs("AppointUserGroup")) Then AppointUserGroup=split(DebateRs("AppointUserGroup"),",")
	AccessFile=DebateRs("AccessFile")
elseif action="editaction" then 
	DebateID=DebateRs("ParentID")
	if DebateID=0 then 
		DebateID=DebateRs("DebateID")
	end if
	DebateRs("title")=NoSqlHack(Trim(Request.Form("title")))
	DebateRs("Content")=NoHtmlHackInput(Request.Form("Content"))
	DebateRs("AppointUserNumber")=NoSqlHack(Request.Form("AppointUserNumber"))
	if trim(Request.Form("userGroup"))<>"" or trim(Request.Form("corpGroup"))<>"" then
		DebateRs("AppointUserGroup")=NoSqlHack(Request.Form("userGroup"))&","&NoSqlHack(Request.Form("corpGroup"))
	End if
	DebateRs("AddTime")=Now
	DebateRs("UserNumber")=session("FS_UserNumber")
	DebateRs("AddIP")=CheckIpSafe(request.ServerVariables("REMOTE_ADDR"))
	If  Trim(request.Form("file"))<>"" then
		DebateRs("AccessFile")=NoSqlHack(Trim(request.Form("file")))
	End if
	DebateRs.update
	DebateRs.close
	if err.number=0 then 
		Response.Redirect("lib/success.asp?ErrCodes=<li>修改成功</li>&ErrorURL=../Debate_unit.asp?gdid="&ClassID&"***DebateID="&DebateID)
		Response.End()
	else
		Response.Redirect("lib/error.asp?ErrCodes=<li>"&err.description&"</li>")
		Response.End()
	end if
elseif action="revert" then
	Debate_Title="回复:"&Request.QueryString("title")
elseif action="revertaction" then
	DebateRs.addNew
	DebateRs("title")=NoSqlHack(Request.Form("title"))
	DebateRs("Content")=NoSqlHack(Request.Form("Content"))
	DebateRs("AppointUserNumber")=NoSqlHack(Request.Form("AppointUserNumber"))
	if trim(Request.Form("userGroup"))<>"" or trim(Request.Form("corpGroup"))<>"" then
		DebateRs("AppointUserGroup")=NoSqlHack(Request.Form("userGroup"))&","&NoSqlHack(Request.Form("corpGroup"))
	End if
	DebateRs("ClassID")=classID
	DebateRs("ParentID")=DebateID
	DebateRs("AddTime")=Now
	DebateRs("UserNumber")=session("FS_UserNumber")
	DebateRs("AddIP")=NoSqlHack(CheckIpSafe(request.ServerVariables("REMOTE_ADDR")))
	If Trim(request.Form("file"))<>"" then
		DebateRs("AccessFile")=NoSqlHack(request.Form("file"))
	End if
	DebateRs.update
	DebateRs.close
	if err.number=0 then 
		Response.Redirect("lib/success.asp?ErrCodes=<li>添加成功</li>&ErrorURL=../Debate_unit.asp?gdid="&ClassID&"***DebateID="&DebateID)
		Response.End()
	else
		Response.Redirect("lib/error.asp?ErrCodes=<li>"&err.description&"</li>")
		Response.End()
	end if
end if
if not DebateRs.eof then
	AppointUserNumber=DebateRs("AppointUserNumber")
	if DebateRs("AppointUserGroup")<>"" then
		AppointUserGroup=split(DebateRs("AppointUserGroup"),",")
	end if
else
	AppointUserGroup=null
end if

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head></head>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="javascript" src="lib/UserJs.js" type="text/javascript"></script>
<body> 
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table"> 
  <tr> 
    <td> <!--#include file="top.asp" --> </td> 
  </tr> 
</table> 
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table"> 
  <tr class="back"> 
    <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" --> </td> 
  </tr> 
  <tr class="back"> 
    <td width="18%" valign="top" class="hback"> <div align="left"> 
        <!--#include file="menu.asp" --> 
      </div></td> 
    <td width="82%" valign="top" class="hback"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"> 
        <tr> 
          <td width="72%"  valign="top"> <table width="98%" height="112" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
              <tr class="hback_1"> 
                <td height="33">&nbsp;&nbsp;&nbsp;<a href='GroupClass.asp?Act=Add'>创建社群</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='myGroup.asp?userNumber=<%=session("FS_UserNumber")%>'>我的社群</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='#'>社群帮助</a></td> 
              </tr>
              <%
						'-------------------------------------群介绍
						ClassID=GroupRs("gdID")
						GroupName=GroupRs("title")
						CroupContent=GroupRs("content")
						InfoType=GroupRs("InfoType")
						ClassType=GroupRs("ClassType")
						AddTime=GroupRs("AddTime")
						GroupCreater=GroupRs("UserNumber")
						GroupManager=GroupRs("AdminName")
						isSys=GroupRs("isSys")
						'----------------------获得管理员------------------
						GroupManagerArray=split(GroupManager,",")
						GroupManager=""
						for ForIndex=0 to ubound(GroupManagerArray)
							Set TempRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&GroupManagerArray(ForIndex)&"'")
							if not TempRs.eof then
								GroupManager=GroupManager&","&TempRs("UserName")
							end if
						next	
						GroupManager=DelHeadAndEndDot(GroupManager)												
						'-------------获得创建人-------------------------------
						if isSys=1 then
							GroupCreater="管理员"
						else
							Set TempRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&GroupCreater&"'")
							if not TempRs.eof then
								GroupCreater=TempRs("UserName")
							end if
						end if
						'-----------------获得社群所属行业----------------------
						Set TempRs=User_Conn.execute("select vClassName from FS_ME_VocationClass where vcid="&ClassType)
						if not TempRs.eof then
							ClassType=TempRs("vClassName")
						else
							ClassType="其他"
						end if
						'-------------------------------------
						Response.Write("<tr class='hback'><td>")
						Response.Write("<table width='100%' height='60' border='0' cellpadding='3' cellspacing='1' class='table'>"&Chr(10)&chr(13))
						Response.Write("<tr class='hback'><td colspan=9><a href='Group_unit.asp?GDID="&ClassID&"'><strong><img src=""images/GroupUser.gif"" border=""0""/> "&GroupName&"</strong></a></td></tr>"&Chr(10)&chr(13))
						Response.Write("<tr height='20'>"&Chr(10)&chr(13))						
						Response.Write("<td class='hback' align='right' width='8%'>创建时间:</td><td class='hback' align='left'>"&Datevalue(AddTime)&"</td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>创建人:</td><td class='hback' align='left'><a href=""ShowUser.asp?UserName="& GroupCreater &""" target=""_blank"">"&GroupCreater&"</a></td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>管理员:</td><td class='hback' align='left'><a href=""ShowUser.asp?UserName="& GroupManager &""" target=""_blank"">"&GroupManager&"</a></td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>所属行业:</td><td class='hback' align='left'>"&ClassType&"</td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>"&Chr(10)&chr(13))
						if Instr(GroupMembers,session("FS_UserNumber"))=0 then
							Response.Write("<a href='Group_unit.asp?act=join&GDID="&ClassID&"'>加入该社区</a>")
						else
							Response.Write("<a href='Group_unit.asp?act=exit&GDID="&ClassID&"'>退出该社区</a>")
						end if
						Response.Write("</td></tr>"&Chr(10)&chr(13))
						Response.Write("<tr height='30'>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' colspan=9>&nbsp;<img src=""images/GroupNews.gif""/>"&CroupContent&"</td>"&Chr(10)&chr(13))
						Response.Write("</tr>"&Chr(10)&chr(13))
						Response.Write("</table>")
						Response.Write("</td></tr>")
			  %>
		  </table></td> 
        </tr> 
		<tr>
		<td>		
			<%
				if action="new" then 
					Response.Write("<form action='?act=addaction&classid="&NoSqlHack(Request.QueryString("classid"))&"' method='post' name='addDebateForm' id='addDebateForm'>")
				elseif action="edit" then
					Response.Write("<form action='?act=editaction&classID="&NoSqlHack(Request.QueryString("classid"))&"&DebateID="&NoSqlHack(Debate_ID)&"' method='post' name='addDebateForm' id='addDebateForm'>")
				elseif action="revert" then
					Response.Write("<form action='?act=revertaction&classID="&NoSqlHack(Request.QueryString("classid"))&"&DebateID="&NoSqlHack(request.QueryString("DebateID"))&"' method='post' name='addDebateForm' id='addDebateForm'>")
				end if
			%>
			<table width="98%" height="110" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
			<tr>
			<td width="19%" align="right" class="hback">主题：</td>
			<td width="81%" class="hback"><input name="title" type="text" id="title" size="50" value="<%=Debate_Title%>"><span id="title_Alert"></span></td>
			</tr>
			<tr>
			  <td height="35" align="right" class="hback">上传附件：</td>
			  <td class="hback"><input name="file" type="text" size="50" value="<%=AccessFile%>"> <input type="button" name="Submit4" value="选择图片" onClick="var TempReturnValue=OpenWindow('Commpages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,290,window);if (TempReturnValue!='') document.addDebateForm.file.value=TempReturnValue;"></td>
			  </tr>
			<tr>
			  <td height="35" align="right" class="hback">内容：</td>
			  <td class="hback"><textarea name="content" cols="50" rows="10" id="content"><%=Debate_Content%></textarea><span id="content_Alert"></span></td>
			</tr>
			<tr>
			  <td height="35" align="right" class="hback">该主题个人查看权限：</td>
			  <td class="hback"><textarea name="AppointUserNumber" cols="50" id="AppointUserNumber"><%=AppointUserNumber%></textarea>
			    [填写允许查看人的编号，用,隔开]</td>
			  </tr>
			<tr>
			  <td align="right" class="hback">该主题组查看权限：<br></td>
			  <td class="hback">
			  <select name="userGroup" size="10" multiple>
          <option value="" style="background-color:#014952; color:#FFFFFF;">-个人用户组-</option>
		  <%
		  	Set UserGroupRs=User_Conn.Execute("Select GroupID,GroupName from FS_ME_Group where GroupType=1")
			while not UserGroupRs.eof
				selectedTF=false
				if  isArray(AppointUserGroup) then
					for ForIndex=0 to Ubound(AppointUserGroup)
						if not isnumeric(AppointUserGroup(ForIndex)) then exit for
						if Cint(trim(AppointUserGroup(ForIndex)))=UserGroupRs("GroupID") then
							selectedTF=true
							exit for
						end if
					next
				end if
				if selectedTF then
					Response.Write("<option value='"&UserGroupRs("GroupID")&"' selected>"&UserGroupRs("GroupName")&"</option>")
				else
					Response.Write("<option value='"&UserGroupRs("GroupID")&"'>"&UserGroupRs("GroupName")&"</option>")
				end if
				UserGroupRs.movenext
			wend
		  %>
        </select> 
          <select name="corpGroup" size="10" multiple>
            <option value="" style="background-color:#014952; color:#FFFFFF;">-企业用户组-</option>
			<%
			Set UserGroupRs=User_Conn.Execute("Select GroupID,GroupName from FS_ME_Group where GroupType=0")
			while not UserGroupRs.eof
				selectedTF=false
				if  not isNull(AppointUserGroup) then
					for ForIndex=0 to Ubound(AppointUserGroup)
						if not isnumeric(AppointUserGroup(ForIndex)) then exit for
						if Cint(NoSqlHack(AppointUserGroup(ForIndex)))=UserGroupRs("GroupID") then
							selectedTF=true
							exit for
						end if
					next
				end if
				if selectedTF then
					Response.Write("<option value='"&UserGroupRs("GroupID")&"' selected>"&UserGroupRs("GroupName")&"</option>")
				else
					Response.Write("<option value='"&UserGroupRs("GroupID")&"'>"&UserGroupRs("GroupName")&"</option>")
				end if
				UserGroupRs.movenext
			wend
			%>
            </select>
          [选择的组将允许查看]</td>
			</tr>
			<tr>
			  <td align="right" class="hback">&nbsp;</td>
			  <td class="hback"><input type="Button" name="Submit" onClick="MySubming()" value="提交">
&nbsp;&nbsp;
<input type="reset" name="Submit2" value="重置"></td>
			</tr>
			</table>
		</form>
		</td>
		</tr>
		
      </table></td> 
  </tr> 
  <tr class="back"> 
    <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
        <!--#include file="Copyright.asp" --> 
      </div></td> 
  </tr> 
</table> 
</body>
</html>
<script language="javascript">
function MySubming()
{
	var flag1=isEmpty("title","title_Alert")
	var flag2=isEmpty("content","content_Alert")
	if(flag1&&flag2)
		document.addDebateForm.submit()
}
</script>
<%
Set GroupRs=nothing
Set DebateRs=nothing
Set UserGroupRs=nothing
Set User_Conn=nothing
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






