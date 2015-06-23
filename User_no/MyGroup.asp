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
'---------------------------------分页定义
int_RPP=15 '设置每页显示数目
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
Set GroupRs=Server.CreateObject(G_FS_RS)
GroupRs.open "select gdID,title,content,InfoType,ClassType,AddTime,UserNumber,AdminName,ClassMember,isSys,hits from FS_ME_GroupDebateManage where UserNumber like '"&session("FS_Usernumber")&"'",User_Conn,1,3
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
              <tr class="xingmu"> 
                <td height="33">&nbsp;&nbsp;&nbsp;
				<a href='GroupClass.asp?Act=Add'>创建社群</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='myGroup.asp?userNumber=<%=session("FS_UserNumber")%>'>我的社群</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='Group.asp'>社群首页</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a href='#'>社群帮助</a>
				</td> 
              </tr>
              <%
					If Not GroupRs.eof then
					'分页使用-----------------------------------
						GroupRs.PageSize=int_RPP
						cPageNo=NoSqlHack(Request.QueryString("page"))
						If cPageNo="" Then cPageNo = 1
						If not isnumeric(cPageNo) Then cPageNo = 1
						cPageNo = Clng(cPageNo)
						If cPageNo<=0 Then cPageNo=1
						If cPageNo>GroupRs.PageCount Then cPageNo=GroupRs.PageCount 
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
						'-----------------获得社群成员数----------------------
						if GroupMembers<>"" then
							GroupMembersArray=split(GroupMembers,",")
						end if
						'------------------------------------------
						Response.Write("<tr><td>")
						Response.Write("<table width='100%' height='60' border='0' cellpadding='1' cellspacing='1' class='table'>"&vbcrlf)
						Response.Write("<tr class='xingmu'><td colspan=10><a href='Group_unit.asp?GDID="&GroupID&"'><img src=""images/GroupUser.gif"" border=""0""/><strong> "&GroupName&"</strong></a></td></tr>"&vbcrlf)
						Response.Write("<tr height='20'>"&vbcrlf)
						Response.Write("<td align='center' class='hback'>现有成员</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center' width='20%'>创建时间</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center' width='12%'>创建人</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center'>管理员</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center'>所属行业</td>"&vbcrlf)
						Response.Write("</tr>"&vbcrlf)
						Response.Write("<tr>"&vbcrlf)
						if instr(GroupMembers,",")>0 then
							Response.Write("<td align='center' class='hback'>"&Ubound(GroupMembersArray)&"人</td>"&vbcrlf)
						else
							Response.Write("<td align='center' class='hback'>0人</td>"&vbcrlf)
						end if
						Response.Write("<td class='hback' align='center' width='20%'>"&Datevalue(AddTime)&"</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center' width='12%'>"&GroupCreater&"</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center'>"&GroupManager&"</td>"&vbcrlf)
						Response.Write("<td class='hback' align='center'>"&ClassType&"</td>"&vbcrlf)
						Response.Write("</tr>"&vbcrlf)
						Response.Write("<tr height='40'>"&vbcrlf)
						Response.Write("<td class='hback' colspan=5><img src=""images/GroupNews.gif""/>"&CroupContent&"</td>"&vbcrlf)
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
                <td height="27" class="xingmu"><span class="bigtitle"><strong>·热门社群</strong></span></td> 
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
            <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
              <tr> 
                <td width="50%" height="27"  class="xingmu"><span class="bigtitle"><strong>·社群排行</strong></span></td> 
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
                <td width="50%" height="27"  class="xingmu"><span class="bigtitle"><strong>·社群分类</strong></span></td> 
              </tr> 
              <tr> 
                <td height="94" valign="top" class="hback">新闻，下载，商品，房产，供求，求职，招聘，其它 </td> 
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->






