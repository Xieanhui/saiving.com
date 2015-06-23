<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp"-->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim GroupRs,DebateRs,DebateID,GroupID,GroupName,CroupContent,InfoType,ClassType,AddTime,PerPageNum,GroupCreater,GroupManager,isSys,hits,GroupMembers,GroupMembersArray,TempRs,HotGroupNumber,ForwardNumber,ForIndex,GroupManagerArray,re,AppointUserNumber,AppointUserGroup
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
GroupID=CintStr(Request.QueryString("GDID"))
Set GroupRs=Server.CreateObject(G_FS_RS)
Set DebateRs=Server.CreateObject(G_FS_RS)

GroupRs.open "select gdID,title,content,InfoType,ClassType,AddTime,PerPageNum,UserNumber,ClassMemberNum,AdminName,ClassMember,isSys,hits from FS_ME_GroupDebateManage where gdID="&GroupID,User_Conn,1,3

DebateRs.open "select DebateID,title,content,ParentID,UserNumber,AddTime,AppointUserNumber,AppointUserGroup,AddIP from FS_ME_GroupDebate  where ClassID="&CintStr(GroupID)&" and ParentID=0 order by AddTime desc,DebateID desc",User_Conn,1,1

PerPageNum=GroupRs("PerPageNum")
'---------------------------------分页定义
int_RPP=PerPageNum '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"			'尾页
'--------------------------------------------------

if NoSqlHack(Request.QueryString("act"))="join" then 
	GroupMembers=GroupRs("ClassMember")
	GroupMembers=DelHeadAndEndDot(GroupMembers)&"," &session("FS_UserNumber")
	GroupMembersArray=split(GroupMembers,",")
	GroupRs("ClassMember")=NoSqlHack(GroupMembers)
	GroupRs("ClassMemberNum")=Ubound(GroupMembersArray)
	GroupRs.update
	GroupRs.close
	if err.number=0 then 
		Response.Redirect("lib/success.asp?ErrCodes=<li>成功加入该组</li>&ErrorURL=../Group_unit.asp?GDID="&GroupID)
		Response.End()
	else
		Response.Redirect("lib/error.asp?ErrCodes=<li>"&err.description&"</li>")
		Response.End()
	end if
elseif NoSqlHack(Request.QueryString("act"))="exit" then
	GroupMembers=GroupRs("ClassMember")
	Set re = New RegExp '正则匹配一个逗号
	re.Pattern = session("FS_UserNumber")&",*"
	GroupMembers=re.replace(GroupMembers,"")
	GroupMembersArray=split(GroupMembers,",")
	GroupRs("ClassMember")=GroupMembers
	GroupRs("ClassMemberNum")=Ubound(GroupMembersArray)
	GroupRs.update
	GroupRs.close
	if err.number=0 then 
		Response.Redirect("lib/success.asp?ErrCodes=<li>成功退出该组</li>&ErrorURL=../Group_unit.asp?GDID="&GroupID)
		Response.End()
	else
		Response.Redirect("lib/error.asp?ErrCodes=<li>"&err.description&"</li>")
		Response.End()
	end if
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head></head>
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
                <td height="33">&nbsp;&nbsp;&nbsp;
				<a href='GroupClass.asp?Act=Add'>创建社群</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='myGroup.asp?userNumber=<%=session("FS_UserNumber")%>'>我的社群</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='Group.asp'>社群首页</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='#'>社群帮助</a>
				</td> 
              </tr>
              <%
						'-------------------------------------群介绍
						GroupID=GroupRs("gdID")
						GroupName=GroupRs("title")
						CroupContent=GroupRs("content")
						InfoType=GroupRs("InfoType")
						ClassType=GroupRs("ClassType")
						AddTime=GroupRs("AddTime")
						GroupCreater=GroupRs("UserNumber")
						GroupManager=GroupRs("AdminName")
						isSys=GroupRs("isSys")
						hits=GroupRs("hits")
						GroupMembers=GroupRs("ClassMember")
						'----------------------获得管理员------------------
						GroupManagerArray=split(GroupManager,",")
						GroupManager=""
						for ForIndex=0 to ubound(GroupManagerArray)
							Set TempRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&NoSqlHack(GroupManagerArray(ForIndex))&"'")
							if not TempRs.eof then
								GroupManager=GroupManager&","&TempRs("UserName")
							end if
						next	
						GroupManager=DelHeadAndEndDot(GroupManager)												
						'-------------获得创建人-------------------------------
						if isSys=1 then
							GroupCreater="管理员"
						else
							Set TempRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&NoSqlHack(GroupCreater)&"'")
							if not TempRs.eof then
								GroupCreater=TempRs("UserName")
							end if
						end if
						'-----------------获得社群所属行业----------------------
						Set TempRs=User_Conn.execute("select vClassName from FS_ME_VocationClass where vcid="&CintStr(ClassType))
						if not TempRs.eof then
							ClassType=TempRs("vClassName")
						else
							ClassType="其他"
						end if
						'-----------------获得社群成员数----------------------
						if GroupMembers<>"" then
							GroupMembersArray=split(GroupMembers,",")
						end if
						'------------------------------------------
						Response.Write("<tr class='hback'><td>")
						Response.Write("<table width='100%' height='60' border='0' cellpadding='3' cellspacing='1' class='table'>"&Chr(10)&chr(13))
						Response.Write("<tr class='hback'><td colspan=9><a href='Group_unit.asp?GDID="&GroupID&"'><strong><img src=""images/GroupUser.gif"" border=""0""/> "&GroupName&"</strong></a></td></tr>"&Chr(10)&chr(13))
						Response.Write("<tr height='20'>"&Chr(10)&chr(13))						
						Response.Write("<td class='hback' align='right' width='8%'>创建时间:</td><td class='hback' align='left'>"&Datevalue(AddTime)&"</td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>创建人:</td><td class='hback' align='left'><a href=""ShowUser.asp?UserName="& GroupCreater &""" target=""_blank"">"&GroupCreater&"</a></td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>管理员:</td><td class='hback' align='left'><a href=""ShowUser.asp?UserName="& GroupManager &""" target=""_blank"">"&GroupManager&"</a></td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>所属行业:</td><td class='hback' align='left'>"&ClassType&"</td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>"&Chr(10)&chr(13))
						if Instr(GroupMembers,session("FS_UserNumber"))=0 then
							Response.Write("<a href='Group_unit.asp?act=join&GDID="&GroupID&"'>加入该社区</a>")
						else
							Response.Write("<a href='Group_unit.asp?act=exit&GDID="&GroupID&"'>退出该社区</a>")
						end if
						Response.Write("</td>")
						Response.Write("</tr>"&Chr(10)&chr(13))
						Response.Write("<tr height='30'>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' colspan=9><img src=""images/GroupNews.gif""/>"&CroupContent&"</td>"&Chr(10)&chr(13))
						Response.Write("</tr>"&Chr(10)&chr(13))
						Response.Write("</table>")
						Response.Write("</td></tr>")
						'-------------------------------------------------社群成员
						GroupMembers=""
						Set GroupMembersArray=nothing
						Response.Write("<tr class='hback'><td>"&Chr(10)&chr(13))
						Response.Write("<table width='100%' height='60' border='0' cellpadding='5' cellspacing='1' class='table'>"&Chr(10)&chr(13))

						Response.Write("<tr class='hback'><td align='left'><img src=""images/GroupMembers.gif"">社群成员</td></tr>")
						if GroupRs("ClassMember")<>"" then
							GroupMembersArray=split(DelHeadAndEndDot(GroupRs("ClassMember")),",")
							If IsArray(GroupMembersArray) Then
								Dim memberRs,userName
								for ForIndex =0 to Ubound(GroupMembersArray)
									Set memberRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&NoSqlHack(GroupMembersArray(ForIndex))&"'")
									If Not memberRs.eof Then
										userName="<a  href='ShowUser.asp?UserNumber="&GroupMembersArray(ForIndex)&"' title='点击查看该用户详情'>"&memberRs("UserName")&"</a>"
									Else
										userName=""
									End If
									If userName<>"" then
										GroupMembers=GroupMembers&"&nbsp;<img src='images/DebateUser.gif'/>&nbsp;"&userName&"</a>"
									End if
								Next
								If Not isNull(memberRs) Then memberRs.close()
							End if
						end if
						Response.Write("<tr><td  class='hback'>")
						Response.Write(GroupMembers)
						Response.Write("</td></tr>"&Chr(10)&chr(13))
						Response.Write("</table>"&Chr(10)&chr(13))
						Response.Write("</td></tr>"&Chr(10)&chr(13))
						'-------------------------------------------------社群讨论贴
						Response.Write("<tr class='hback'><td>"&Chr(10)&chr(13))
						Response.Write("<table width='100%' height='60' border='0' cellpadding='5' cellspacing='1' class='table'>"&Chr(10)&chr(13))

						Response.Write("<tr class='hback'><td align='left' colspan='2'>&nbsp;社群讨论<div align='right'><a href='Debate_Add.asp?act=new&ClassID="&GroupID&"'><img src=""images/newTopic.gif"" border=""0""/>发表主题</a></div></td></tr>"&Chr(10)&chr(13))
						
					If Not DebateRs.eof then
					'分页使用-----------------------------------
						DebateRs.PageSize=int_RPP
						cPageNo=NoSqlHack(Request.QueryString("page"))
						If cPageNo="" Then cPageNo = 1
						If not isnumeric(cPageNo) Then cPageNo = 1
						cPageNo = Clng(cPageNo)
						If cPageNo<=0 Then cPageNo=1
						If cPageNo>DebateRs.PageCount Then cPageNo=DebateRs.PageCount 
						DebateRs.AbsolutePage=cPageNo
					End if
					for i=0 to int_RPP
						if DebateRs.eof then exit for
						AppointUserNumber=DebateRs("AppointUserNumber")
						AppointUserGroup=DebateRs("AppointUserGroup")
						if AppointUserNumber<>"" then
							if instr(AppointUserNumber,Session("FS_UserNumber"))=0 then
								Response.Write("<tr height='20' class='hback'>")
								Response.Write("<td width='3%' align=""Center""><img src=""images/GroupTopic.gif""></td>")
								Response.Write("<td align='left'>该贴被作者隐藏</td>"&Chr(10)&chr(13))
								Response.Write("</tr>")								
							else
								Response.Write("<tr height='20' class='hback'>")
								Response.Write("<td width='3%' align=""Center""><img src=""images/GroupTopic.gif""></td>")
								Response.Write("<td align='left'><a href='Debate_unit.asp?DebateID="&DebateRs("DebateID")&"&gdid="&GroupID&"'>"&DebateRs("title")&"</a></td>"&Chr(10)&chr(13))
								Response.Write("</tr>")
							end if
						elseif AppointUserGroup<>"" then
							if instr(AppointUserGroup,Session("FS_Group"))=0 then
								Response.Write("<tr height='20' class='hback'>")
								Response.Write("<td width='3%' align=""Center""><img src=""images/GroupTopic.gif""></td>")
								Response.Write("<td align='left'>该贴被作者隐藏</td>"&Chr(10)&chr(13))
								Response.Write("</tr>")
							else			
								Response.Write("<tr height='20' class='hback'>")
								Response.Write("<td width='3%' align=""Center""><img src=""images/GroupTopic.gif""></td>")
								Response.Write("<td align='left'><a href='Debate_unit.asp?DebateID="&DebateRs("DebateID")&"&gdid="&GroupID&"'>"&DebateRs("title")&"</a></td>"&Chr(10)&chr(13))
								Response.Write("</tr>")
							end if
						else
							Response.Write("<tr height='20' class='hback'>")
							Response.Write("<td width='3%' align=""Center""><img src=""images/GroupTopic.gif""></td>")
							Response.Write("<td align='left'><a href='Debate_unit.asp?DebateID="&DebateRs("DebateID")&"&gdid="&GroupID&"'>"&DebateRs("title")&"</a></td>"&Chr(10)&chr(13))
							Response.Write("</tr>")
						end if
						DebateRs.movenext
					next
					Response.Write("</table>"&Chr(10)&chr(13))
					Response.Write("</td></tr>"&Chr(10)&chr(13))

			  %>
			  <%
				Response.Write("<tr>"&vbcrlf)
				Response.Write("<td align='right' colspan='5'  class=""hback"">"&fPageCount(DebateRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
				Response.Write("</tr>"&vbcrlf)
			%>
          </table></td> 
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
<%
GroupRs.close
DebateRs.close
Set GroupRs=nothing
Set DebateRs=nothing
Set User_Conn=nothing
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






