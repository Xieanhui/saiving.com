<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp"-->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%

Dim GroupRs,GroupID,GroupName,CroupContent,InfoType,ClassType,AddTime,GroupCreater,GroupManager,isSys,hits,GroupMembers,GroupMembersArray,TempRs,HotGroupNumber,ForwardNumber,ForIndex,GroupManagerArray
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
'---------------------------------��ҳ����
int_RPP=8'����ÿҳ��ʾ��Ŀ
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
Set GroupRs=Server.CreateObject(G_FS_RS)
GroupRs.open "select gdID,title,content,InfoType,ClassType,AddTime,UserNumber,AdminName,ClassMember,isSys,hits from FS_ME_GroupDebateManage",User_Conn,1,1
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
                <td height="26"><a href="GroupClass.asp?Act=Add">������Ⱥ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='MyGroup.asp?userNumber=<%=session("FS_UserNumber")%>'>�ҵ���Ⱥ</a></td> 
              </tr>
              <%
					If Not GroupRs.eof then
					'��ҳʹ��-----------------------------------
						GroupRs.PageSize=int_RPP
						cPageNo=NoSqlHack(Request.QueryString("page"))
						If cPageNo="" Then cPageNo = 1
						If not isnumeric(cPageNo) Then cPageNo = 1
						cPageNo = Clng(cPageNo)
						If cPageNo>GroupRs.PageCount Then cPageNo=GroupRs.PageCount 
						If cPageNo<=0 Then cPageNo=1
						GroupRs.AbsolutePage=cPageNo
					End if
					for i=0 to int_RPP
						if GroupRs.eof then exit for
						'-------------------------------------
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
							Set TempRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&GroupManagerArray(ForIndex)&"'")
							if not TempRs.eof then
								GroupManager=GroupManager&","&TempRs("UserName")
							end if
						next	
						GroupManager=DelHeadAndEndDot(GroupManager)												
						'-------------��ô�����-------------------------------
						if isSys=1 then
							GroupCreater="����Ա"
						else
							Set TempRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&GroupCreater&"'")
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
						Response.Write("<tr><td>")
						Response.Write("<table width='100%' height='60' border='0' cellpadding='1' cellspacing='1' class='table'>"&vbcrlf)
						Response.Write("<tr class='xingmu'><td colspan=10><a href='Group_unit.asp?GDID="&GroupID&"' class=""Top_Navi""><img src=""images/GroupUser.gif"" border=""0""/><strong>"&GroupName&"</strong></a></td></tr>"&vbcrlf)
						Response.Write("<tr height='20'>"&vbcrlf)
						Response.Write("<td align='center' class='hback'>���г�Ա</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center' width='20%'>����ʱ��</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center' width='12%'>������</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center'>����Ա</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center'>������ҵ</td>"&vbcrlf)
						Response.Write("</tr>"&vbcrlf)
						Response.Write("<tr>"&vbcrlf)
						if instr(GroupMembers,",")>0 then
							Response.Write("<td align='center' class='hback'>"&Ubound(GroupMembersArray)+1&"��</td>"&vbcrlf)
						else
							Response.Write("<td align='center' class='hback'>0��</td>"&vbcrlf)
						end if
						Response.Write("<td class='hback' align='center' width='20%'>"&Datevalue(AddTime)&"</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center' width='12%'>"&GroupCreater&"</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center'>"&GroupManager&"</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center'>"&ClassType&"</td>"&vbcrlf)
						Response.Write("</tr>"&vbcrlf)
						Response.Write("<tr height='40'>"&vbcrlf)
						Response.Write("<td class='hback' colspan=5><img src=""images/GroupNews.gif""/>��"&CroupContent&"</td>"&vbcrlf)
						Response.Write("</tr>"&vbcrlf)
						Response.Write("</table>")
						Response.Write("</td></tr>")
						GroupRs.movenext
					next
			  %>
			  <%
				Response.Write("<tr>"&vbcrlf)
				Response.Write("<td align='right' colspan='5'  class=""hback"">"&fPageCount(GroupRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
				Response.Write("</tr>"&vbcrlf)
			%>
            </table></td> 
          <td width="32%" valign="top"> <table width="98%"  border="0" align="center" cellpadding="1" cellspacing="1" class="table"> 
              <tr> 
                <td height="27" class="xingmu"><span class="bigtitle"><strong>��������Ⱥ</strong></span></td> 
              </tr> 
				<%
					HotGroupNumber=5
					Set TempRs=User_Conn.execute("Select gdID,title from FS_ME_GroupDebateManage order by hits desc")
					for ForIndex=0 to HotGroupNumber
						if TempRs.eof then exit for
						Response.Write("<tr class='hback'>"&vbcrlf)
						Response.Write("<td  valign='top' align='left' height='20'>&nbsp;"&ForIndex+1&".<a href='Group_unit.asp?GDID="&TempRs("gdID")&"'>"&TempRs("title")&"</a></td>"&vbcrlf)
						Response.Write("</tr>"&vbcrlf)
						TempRs.movenext
					next
				%>
            </table> 
            <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
              <tr> 
                <td width="50%" height="27"  class="xingmu"><span class="bigtitle"><strong>����Ⱥ����</strong></span></td> 
              </tr> 
			  <%
					ForwardNumber=5
					Set TempRs=User_Conn.execute("Select gdID,title from FS_ME_GroupDebateManage order by ClassMemberNum desc")
					for ForIndex=0 to ForwardNumber
						if TempRs.eof then exit for
						Response.Write("<tr class='hback'>"&vbcrlf)
						Response.Write("<td  valign='top' align='left' height='20'>&nbsp;"&ForIndex+1&".<a href='Group_unit.asp?GDID="&TempRs("gdID")&"'>"&TempRs("title")&"</a></td>"&vbcrlf)
						Response.Write("</tr>"&vbcrlf)
						TempRs.movenext
					next
			  %>
            </table> 
            <table width="98%" height="124" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
              <tr> 
                <td width="50%" height="27"  class="xingmu"><span class="bigtitle"><strong>����Ⱥ����</strong></span></td> 
              </tr> 
              <tr> 
                <td height="94" valign="top" class="hback">���ţ����أ���Ʒ��������������ְ����Ƹ������ </td> 
              </tr> 
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
Set GroupRs=nothing
Set User_Conn=nothing
Set TempRs=nothing
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






