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
'---------------------------------��ҳ����
int_RPP=PerPageNum '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"			'βҳ
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
		Response.Redirect("lib/success.asp?ErrCodes=<li>�ɹ��������</li>&ErrorURL=../Group_unit.asp?GDID="&GroupID)
		Response.End()
	else
		Response.Redirect("lib/error.asp?ErrCodes=<li>"&err.description&"</li>")
		Response.End()
	end if
elseif NoSqlHack(Request.QueryString("act"))="exit" then
	GroupMembers=GroupRs("ClassMember")
	Set re = New RegExp '����ƥ��һ������
	re.Pattern = session("FS_UserNumber")&",*"
	GroupMembers=re.replace(GroupMembers,"")
	GroupMembersArray=split(GroupMembers,",")
	GroupRs("ClassMember")=GroupMembers
	GroupRs("ClassMemberNum")=Ubound(GroupMembersArray)
	GroupRs.update
	GroupRs.close
	if err.number=0 then 
		Response.Redirect("lib/success.asp?ErrCodes=<li>�ɹ��˳�����</li>&ErrorURL=../Group_unit.asp?GDID="&GroupID)
		Response.End()
	else
		Response.Redirect("lib/error.asp?ErrCodes=<li>"&err.description&"</li>")
		Response.End()
	end if
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
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
				<a href='GroupClass.asp?Act=Add'>������Ⱥ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='myGroup.asp?userNumber=<%=session("FS_UserNumber")%>'>�ҵ���Ⱥ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='Group.asp'>��Ⱥ��ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='#'>��Ⱥ����</a>
				</td> 
              </tr>
              <%
						'-------------------------------------Ⱥ����
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
						'----------------------��ù���Ա------------------
						GroupManagerArray=split(GroupManager,",")
						GroupManager=""
						for ForIndex=0 to ubound(GroupManagerArray)
							Set TempRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&NoSqlHack(GroupManagerArray(ForIndex))&"'")
							if not TempRs.eof then
								GroupManager=GroupManager&","&TempRs("UserName")
							end if
						next	
						GroupManager=DelHeadAndEndDot(GroupManager)												
						'-------------��ô�����-------------------------------
						if isSys=1 then
							GroupCreater="����Ա"
						else
							Set TempRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&NoSqlHack(GroupCreater)&"'")
							if not TempRs.eof then
								GroupCreater=TempRs("UserName")
							end if
						end if
						'-----------------�����Ⱥ������ҵ----------------------
						Set TempRs=User_Conn.execute("select vClassName from FS_ME_VocationClass where vcid="&CintStr(ClassType))
						if not TempRs.eof then
							ClassType=TempRs("vClassName")
						else
							ClassType="����"
						end if
						'-----------------�����Ⱥ��Ա��----------------------
						if GroupMembers<>"" then
							GroupMembersArray=split(GroupMembers,",")
						end if
						'------------------------------------------
						Response.Write("<tr class='hback'><td>")
						Response.Write("<table width='100%' height='60' border='0' cellpadding='3' cellspacing='1' class='table'>"&Chr(10)&chr(13))
						Response.Write("<tr class='hback'><td colspan=9><a href='Group_unit.asp?GDID="&GroupID&"'><strong><img src=""images/GroupUser.gif"" border=""0""/> "&GroupName&"</strong></a></td></tr>"&Chr(10)&chr(13))
						Response.Write("<tr height='20'>"&Chr(10)&chr(13))						
						Response.Write("<td class='hback' align='right' width='8%'>����ʱ��:</td><td class='hback' align='left'>"&Datevalue(AddTime)&"</td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>������:</td><td class='hback' align='left'><a href=""ShowUser.asp?UserName="& GroupCreater &""" target=""_blank"">"&GroupCreater&"</a></td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>����Ա:</td><td class='hback' align='left'><a href=""ShowUser.asp?UserName="& GroupManager &""" target=""_blank"">"&GroupManager&"</a></td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>������ҵ:</td><td class='hback' align='left'>"&ClassType&"</td>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' align='right'>"&Chr(10)&chr(13))
						if Instr(GroupMembers,session("FS_UserNumber"))=0 then
							Response.Write("<a href='Group_unit.asp?act=join&GDID="&GroupID&"'>���������</a>")
						else
							Response.Write("<a href='Group_unit.asp?act=exit&GDID="&GroupID&"'>�˳�������</a>")
						end if
						Response.Write("</td>")
						Response.Write("</tr>"&Chr(10)&chr(13))
						Response.Write("<tr height='30'>"&Chr(10)&chr(13))
						Response.Write("<td class='hback' colspan=9><img src=""images/GroupNews.gif""/>"&CroupContent&"</td>"&Chr(10)&chr(13))
						Response.Write("</tr>"&Chr(10)&chr(13))
						Response.Write("</table>")
						Response.Write("</td></tr>")
						'-------------------------------------------------��Ⱥ��Ա
						GroupMembers=""
						Set GroupMembersArray=nothing
						Response.Write("<tr class='hback'><td>"&Chr(10)&chr(13))
						Response.Write("<table width='100%' height='60' border='0' cellpadding='5' cellspacing='1' class='table'>"&Chr(10)&chr(13))

						Response.Write("<tr class='hback'><td align='left'><img src=""images/GroupMembers.gif"">��Ⱥ��Ա</td></tr>")
						if GroupRs("ClassMember")<>"" then
							GroupMembersArray=split(DelHeadAndEndDot(GroupRs("ClassMember")),",")
							If IsArray(GroupMembersArray) Then
								Dim memberRs,userName
								for ForIndex =0 to Ubound(GroupMembersArray)
									Set memberRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&NoSqlHack(GroupMembersArray(ForIndex))&"'")
									If Not memberRs.eof Then
										userName="<a  href='ShowUser.asp?UserNumber="&GroupMembersArray(ForIndex)&"' title='����鿴���û�����'>"&memberRs("UserName")&"</a>"
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
						'-------------------------------------------------��Ⱥ������
						Response.Write("<tr class='hback'><td>"&Chr(10)&chr(13))
						Response.Write("<table width='100%' height='60' border='0' cellpadding='5' cellspacing='1' class='table'>"&Chr(10)&chr(13))

						Response.Write("<tr class='hback'><td align='left' colspan='2'>&nbsp;��Ⱥ����<div align='right'><a href='Debate_Add.asp?act=new&ClassID="&GroupID&"'><img src=""images/newTopic.gif"" border=""0""/>��������</a></div></td></tr>"&Chr(10)&chr(13))
						
					If Not DebateRs.eof then
					'��ҳʹ��-----------------------------------
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
								Response.Write("<td align='left'>��������������</td>"&Chr(10)&chr(13))
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
								Response.Write("<td align='left'>��������������</td>"&Chr(10)&chr(13))
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






