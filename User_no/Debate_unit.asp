<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp"-->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim DebateRs,GroupRs,DebateSubRs,DebateID,GroupID,UserRs,CroupContent,InfoType,ClassType,AddTime,PerPageNum,GroupCreater,GroupManager,isSys,hits,GroupMembers,GroupMembersArray,TempRs,HotGroupNumber,ForwardNumber,ForIndex,GroupManagerArray,AppointUserNumber,AppointUserGroup
'lz_usernumber ¥���û����,
Dim lz_usernumber,creator,admin
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
'---------------------------------��ҳ����
int_RPP=15 '����ÿҳ��ʾ��Ŀ
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

DebateID=CintStr(Request.QueryString("DebateID"))
GroupID=CintStr(Request.QueryString("gdID"))
Set GroupRs=Server.CreateObject(G_FS_RS)
Set DebateRs=Server.CreateObject(G_FS_RS)
Set DebateSubRs=Server.CreateObject(G_FS_RS)
Set UserRs=Server.CreateObject(G_FS_RS)

GroupRs.open "select UserNumber,AdminName from FS_ME_GroupDebateManage where gdID="&CintStr(GroupID),User_Conn,1,3
if not GroupRs.eof then
	creator=GroupRs("UserNumber")
	admin=GroupRs("AdminName")
End if
DebateRs.open "select DebateID,title,content,ParentID,ClassID,UserNumber,AddTime,AppointUserNumber,AppointUserGroup,AddIP,AccessFile from FS_ME_GroupDebate  where DebateID="&CintStr(DebateID),User_Conn,1,1

DebateSubRs.open "select DebateID,title,content,ParentID,ClassID,UserNumber,AddTime,AppointUserNumber,AppointUserGroup,AddIP,AccessFile from FS_ME_GroupDebate  where ParentID="&CintStr(DebateID),User_Conn,1,1

Dim UserNumber,author,sHeadPic,HeadPicSize
UserRs.open "Select UserNumber,UserName,HeadPic,HeadPicSize from FS_ME_Users where UserNumber='"&DebateRs("UserNumber")&"'",User_Conn,1,1
author=""
UserNumber=""
sHeadPic=""
HeadPicSize=""
If Not UserRs.eof Then
	author=UserRs("UserName")
	UserNumber=UserRs("UserNumber")
	sHeadPic=UserRs("HeadPic")
	HeadPicSize=UserRs("HeadPicSize")
End If
If Trim(sHeadPic)="" Or IsNull(sHeadPic) Then
	sHeadPic="images/noHeadpic.gif"
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="lib/css/lightbox.css" type="text/css" media="screen" />
<script type="text/javascript" src="../FS_INC/prototype.js"></script>
<script type="text/javascript" src="lib/js/scriptaculous.js?load=effects"></script>
<script type="text/javascript" src="lib/js/lightbox.js"></script>

</head>
<body onLoad="initLightbox()"> 
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
			<tr class="xingmu"> 
                <td height="33">&nbsp;&nbsp;&nbsp;
				<a href='GroupClass.asp?Act=Add'>������Ⱥ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='myGroup.asp?userNumber=<%=session("FS_UserNumber")%>'>�ҵ���Ⱥ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='Group.asp'>��Ⱥ��ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='#'>��Ⱥ����</a>
				</td> 
              </tr>
              <tr class="xingmu"> 
                <td height="33"><img src="images/GroupUser.gif"/><a href='Group_unit.asp?GDID=<%=DebateRs("classID")%>'><%=User_Conn.execute("select Title from FS_ME_GroupDebateManage where GDID="&DebateRs("classID"))(0)%></a></td> 
              </tr>
              <%

				'-------------------------------------------------��Ⱥ������
				Response.Write("<tr class='xingmu'><td>"&Chr(10)&chr(13))
				Response.Write("<table width='100%' height='60' border='0' cellpadding='1' cellspacing='1' class='table'>"&Chr(10)&chr(13))

				Response.Write("<tr class='hback'><td align='left' colspan='7'>&nbsp;��Ⱥ����</td><td align='center'><a href='Debate_Add.asp?act=new&ClassID="&DebateRs("classID")&"'><img src=""images/newDebate.gif"" border=""0""/>��������</a></td></tr>"&Chr(10)&chr(13))
				Response.Write("<tr height='20' class='hback'>"&Chr(10)&chr(13))
				Response.Write("<td width='3%'><img src=""images/GroupTopic.gif""></td>")
				Response.Write("<td align='left'><strong>"&DebateRs("title")&"</strong></td>"&Chr(10)&chr(13))
				Response.Write("<td align='right'>�����ˣ�</td><td align='left'>"&author&"</td>")
				Response.Write("<td align='right'>����ʱ�䣺</td><td align='left'>"&DebateRs("AddTime")&"</td>")
				Response.Write("<td align=""center"">")
				lz_usernumber=DebateRs("UserNumber")
				if lz_usernumber=session("FS_UserNumber") then
					Response.Write("<a href='Debate_Add.asp?act=edit&classID="&DebateRs("ClassID")&"&DebateID="&DebateRs("DebateID")&"'>�༭</a>")
				end if
				if creator=session("FS_UserNumber") or instr(admin,session("FS_UserNumber"))>0 then
					Response.Write("&nbsp;|&nbsp;<a href=""#"" onClick=""if(confirm('ȷ�Ͻ���ɾ������')){location.href='Debate_action.asp?act=delete&classID="&DebateRs("ClassID")&"&DebateID="&DebateRs("DebateID")&"'}"">ɾ��</a>")
				end if
				Response.Write("</td>")
				Response.Write("<td align=""center""><a href='Debate_Add.asp?act=revert&title="&DebateRs("title")&"&classID="&DebateRs("ClassID")&"&DebateID="&DebateRs("DebateID")&"'>��������</a></td>")
				Response.Write("</tr>"&Chr(10)&chr(13))
				Response.Write("<tr height='30'><td colspan=8>"&Chr(10)&chr(13))
				Response.Write("<table width='100%' height='100' border='0' cellpadding='1' cellspacing='1' class='table'>"&Chr(10)&chr(13))
				Response.Write("<tr><td width=""80"" class=""hback""><img src="""&sHeadPic&""" width=""80"" height=""80""></td>")
				Response.Write("<td class='hback' align='left'><div>"&Encode(DebateRs("content"))&"</div>")
				If DebateRs("AccessFile")<>"" And NOt Isnull(DebateRs("AccessFile")) Then
					Response.write("<div><br><a href='"&DebateRs("AccessFile")&"' rel=""lightbox"" title="""&DebateRs("title")&"""><img src='"&DebateRs("AccessFile")&"' width='80' height='80' border=0/></a></div></td></tr>"&Chr(10)&chr(13))
				Else
					Response.write("</td></tr>"&Chr(10)&chr(13))
				End if
				Response.Write("</table>")
				Response.Write("</td></tr>")
				UserRs.close					
				If Not DebateSubRs.eof then
				'��ҳʹ��-----------------------------------
					DebateSubRs.PageSize=int_RPP
					cPageNo=NoSqlHack(Request.QueryString("page"))
					If cPageNo="" Then cPageNo = 1
					If not isnumeric(cPageNo) Then cPageNo = 1
					cPageNo = Clng(cPageNo)
					If cPageNo>DebateSubRs.PageCount Then cPageNo=DebateSubRs.PageCount 
					If cPageNo<=0 Then cPageNo=1
					DebateSubRs.AbsolutePage=cPageNo
				End if
				for i=0 to int_RPP
					if DebateSubRs.eof then exit For
					if not DebateSubRs.eof then 
						Set UserRs=User_Conn.execute("Select UserNumber,UserName,HeadPic,HeadPicSize from FS_ME_Users where UserNumber='"&DebateSubRs("UserNumber")&"'")
					end If
					If Not UserRs.eof Then
						author=UserRs("UserName")
						UserNumber=UserRs("UserNumber")
						sHeadPic=UserRs("HeadPic")
						HeadPicSize=UserRs("HeadPicSize")
					End If
					AppointUserNumber=DebateSubRs("AppointUserNumber")
					AppointUserGroup=DebateSubRs("AppointUserGroup")
					Response.Write("<tr height='20' class='hback'>"&Chr(10)&chr(13))
					Response.Write("<td width='3%'><img src=""images/Groupreplay.gif""></td>")
					Response.Write("<td align='left'><strong>"&DebateSubRs("title")&"</strong></td>"&Chr(10)&chr(13))
					Response.Write("<td align='right'>�����ˣ�</td><td lign='left'>"&author&"</td>")
					Response.Write("<td align='right'>����ʱ�䣺</td><td lign='left'>"&DebateSubRs("AddTime")&"</td>")
					Response.Write("<td align=""center"">")
					if DebateSubRs("UserNumber")=session("FS_UserNumber") then
						Response.Write("<a href='Debate_Add.asp?act=edit&classID="&DebateSubRs("ClassID")&"&DebateID="&DebateSubRs("DebateID")&"'>�༭</a>")
					end if
					if creator=session("FS_UserNumber") or instr(admin,session("FS_UserNumber"))>0 then
						Response.Write("&nbsp;|&nbsp;<a href=""#"" onClick=""if(confirm('ȷ�Ͻ���ɾ������')){location.href='Debate_action.asp?act=delete&classID="&DebateSubRs("ClassID")&"&DebateID="&DebateSubRs("DebateID")&"'}"">ɾ��</a>")
					end if
					Response.Write("</td>")
					Response.Write("<td>&nbsp;</td>")							
					Response.Write("</tr>")
					Response.Write("<tr height='30'><td  colspan=8>"&Chr(10)&chr(13))
					Response.Write("<table width='100%' height='100' border='0' cellpadding='1' cellspacing='1' class='table'>"&Chr(10)&chr(13))
					Response.Write("<tr><td width='80' class='hback' aling='center' valign='middle'><img src='"&sHeadPic&"' width='80' height='80'></td>")
					if AppointUserNumber<>"" then
						if instr(AppointUserNumber,Session("FS_UserNumber"))=0 then
							Response.Write("<td class='hback'>�����ѱ���������</td></tr>"&Chr(10)&chr(13))
						else
							Response.Write("<td class='hback'><div>"&Encode(DebateSubRs("content"))&"</div>")
							If DebateSubRs("AccessFile")<>"" Then
								Response.write("<div><br><a href='"&DebateSubRs("AccessFile")&"' rel=""lightbox"" title="""&DebateSubRs("title")&"""><img src='"&DebateSubRs("AccessFile")&"' width='80' height='80' border='0'/></a></div></td></tr>"&Chr(10)&chr(13))
							Else
								Response.write("</td></tr>"&Chr(10)&chr(13))
							End If
						End if
					elseif AppointUserGroup<>"" Then
						if instr(AppointUserGroup,Session("FS_Group"))=0 then
							Response.Write("<td class='hback'>�����ѱ���������</td></tr>"&Chr(10)&chr(13))
						else
							Response.Write("<td class='hback'><div>"&Encode(DebateSubRs("content"))&"</div>")
							If DebateSubRs("AccessFile")<>"" Then
								Response.write("<div><br><a href='"&DebateSubRs("AccessFile")&"' rel=""lightbox"" title="""&DebateSubRs("title")&"""><img src='"&DebateSubRs("AccessFile")&"' width='80' height='80' border='0'/></a></div></td></tr>"&Chr(10)&chr(13))
							Else
								Response.write("</td></tr>"&Chr(10)&chr(13))
							End If
						End if
					else
						Response.Write("<td class='hback'><div>"&Encode(DebateSubRs("content"))&"</div>")
						If DebateSubRs("AccessFile")<>"" Then
							Response.write("<div><br><a href="""&DebateSubRs("AccessFile")&""" rel=""lightbox"" title="""&DebateSubRs("title")&"""><img src='"&DebateSubRs("AccessFile")&"' width='80' height='80' border='0'/></a></div></td></tr>"&Chr(10)&chr(13))
						Else
							Response.write("</td></tr>"&Chr(10)&chr(13))
						End if
					end if
					Response.Write("</table>")
					Response.Write("</td></tr>")
					DebateSubRs.movenext
				next
				Response.Write("</table>"&Chr(10)&chr(13))
				Response.Write("</td></tr>"&Chr(10)&chr(13))

			  %>
			<%
				Response.Write("<tr>"&vbcrlf)
				Response.Write("<td align='right' colspan='5'  class=""hback"">"&fPageCount(DebateSubRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
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
Set UserRs=nothing
Set DebateRs=nothing
Set DebateSubRs=nothing
Set User_Conn=nothing
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






