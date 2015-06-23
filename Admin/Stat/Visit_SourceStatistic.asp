<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,User_Conn,tmp_type,strShowErr
MF_Default_Conn
MF_User_Conn
MF_Session_TF
if not MF_Check_Pop_TF("SS_site") then Err_Show
if not MF_Check_Pop_TF("SS001") then Err_Show

%>
<html>
<head>
<title>访问者来源统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body  topmargin="2" leftmargin="2">
<table width="98%" height="56" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td height="28" class="xingmu"><div align="left"><strong>访问者来源统计</strong></div></td>
  </tr>
  <tr>
    <td height="28" class="hback">
      <%
	Dim RsIPObj,Sql,NumUnkown,page_size,page_no,page_total,record_all
	Set RsIPObj = Server.CreateObject(G_FS_RS)
	Sql="Select distinct Area,IP From FS_SS_Stat"
	RsIPObj.Open Sql,Conn,3,3
	Dim IP(),Area(),I
	I=1
	Redim IP(1),Area(1)
	Do While not RsIPObj.Eof
		IP(I)=RsIPObj("IP")
		Area(I)=RsIPObj("Area")
		RsIPObj.MoveNext
		I=I+1
		Redim Preserve IP(I),Area(I)
	Loop
	
	RsIPObj.Close
	Sql="Select Count(*) As RecordNum From FS_SS_Stat"
	RsIPObj.Open Sql,Conn,3,3
	Dim RecordNum
	RecordNum=RsIPObj("RecordNum")
	RsIPObj.Close
	
	Sql="Select LoginNum From FS_SS_Stat"
	RsIPObj.open sql,conn,3,3
	Dim AllNum
	AllNum=0
	do while not RsIPObj.eof
		AllNum=AllNum+Cint(RsIPObj("LoginNum"))
		RsIPObj.movenext
	loop
		
	page_size=50
	page_no=CintStr(request.querystring("page_no"))
	if page_no<=1 or page_no="" then page_no=1
	if ((UBound(IP)-1) mod page_size)=0 then
		page_total=(UBound(IP)-1)\page_size
	else
		page_total=(UBound(IP)-1)\page_size+1
	end if
%>
      <table width=100% border=0 align="center" cellpadding=2>
        <tr> 
          <td align=center>访问者来源统计图表</td>
        </tr>
        <tr> 
          <td align=center>
		  <table align=center>
              <tr valign=cente> 
                <td valign="middle" nowap> <div align="left">共</div></td>
                <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../Images/bar2.gif width="150" height=15></td>
                <td valign="middle" nowap>
                  <% = AllNum %>
                  次</td>
                <td valign="middle" nowap>&nbsp;</td>
                <%
					for I=(page_no*page_size-page_size+1) to page_no*page_size
						if I>UBound(IP) then exit for
						if IP(i)<>"" then
				%>
              <tr valign=bottom > 
                <td valign="middle" nowap><div align="left">
                    <% =IP(i) %>
                  </div></td>
                <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../Images/bar2.gif width="
						<% 
							RsIPObj.close
						   	Sql="select LoginNum from FS_SS_Stat where IP='"& NoSqlHack(IP(i))&"'"
						   	RsIPObj.open Sql,conn,3,3
						   	dim a
						   	a=0
							do while not RsIPObj.eof
						   		a=a+Cint(RsIPObj("LoginNum"))
						   		RsIPObj.movenext
						   	loop
						   response.Write(150*a/AllNum)
						%>
						" height=15></td>
                <td valign="middle" nowap> 
                  <%
					   response.Write(a)
					   %>
                  次&nbsp;&nbsp;</td>
                <td valign="middle" nowap><font color="red">(<%=Area(I)%>)</font></td>
              </tr>
              <%
					else
					end if
				Next
				if cint(page_no) = cint(page_total) then
				%>
              <tr valign=cente> 
                <td valign="middle" nowap><div align="left">未知</div></td>
                <td align=left background=../../Images/Visit/tu_back_2.gif valign=middle width=150><img src=../Images/bar2.gif width="
						<% if AllNum<>0  then
							Response.Write(150*NumUnkown/AllNum)
							else
							Response.Write(0)
							end if
						%>
					" height=15></td>
                <td valign="middle" nowap>
                  <% = NumUnkown %>
                  次</td>
                <td valign="middle" nowap>&nbsp;</td>
                <% end if %>
            </table>
            <%if page_total>1 then%>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td valign="middle"> <div align="right"></div>
                  <div align="right"></div>
                  <div align="right"> 
                    <% =  "NO.<b>"& page_no &"</b>,&nbsp;&nbsp;" %>
                    <% = "Totel:<b>"& page_total &"</b>,&nbsp;RecordCount:<b>" & record_all &"</b>&nbsp;&nbsp;&nbsp;"%>
                  </div></td>
                <td width="15%" valign="bottom"><div align="center"> 
<%
if page_total=1 then
		response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=首页></img>&nbsp;"
		response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=上一页></img>&nbsp;"
		response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=下一页></img>&nbsp;"
		response.Write "&nbsp;<img src=../images/endPage.gif border=0 alt=尾页></img>&nbsp;"
else
	if cint(page_no)<>1 and cint(page_no)<>page_total then
		response.Write "&nbsp;<a href=?page_no=1>首页</a>&nbsp;"
		response.Write "&nbsp;<a href=?page_no="&cstr(cint(page_no)-1)&">上一页</a>&nbsp;"
		response.Write "&nbsp;<a href=?page_no="&cstr(cint(page_no)+1)&">下一页</a>&nbsp;"
		response.Write "&nbsp;<a href=?page_no="& page_total &">最后一页</a>&nbsp;"
	elseif cint(page_no)=1 then
		response.Write "&nbsp;首页</a>&nbsp;"
		response.Write "&nbsp;上一页&nbsp;"
		response.Write "&nbsp;<a href=?page_no="&cstr(cint(page_no)+1)&">下一页</a>&nbsp;"
		response.Write "&nbsp;<a href=?page_no="& page_total &">最后一页</a>&nbsp;"
	else
		response.Write "&nbsp;<a href=?page_no=1>首页&nbsp;"
		response.Write "&nbsp;<a href=?page_no="&cstr(cint(page_no)-1)&">上一页</a>&nbsp;"
		response.Write "&nbsp;下一页</a>&nbsp;"
		response.Write "&nbsp;最后一页&nbsp;"
	end if
end if
%>
                    <select onChange="ChangePage(this.value);" style="width:50;" name="select">
                      <% for i=1 to page_total %>
                      <option <% if cint(page_no) = i then Response.Write("selected")%> value="<% = i %>"> 
                      <% = i %>
                      </option>
                      <% next %>
                    </select>
                  </div></td>
              </tr>
            </table>
            <%end if%>
      </table></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
function ChangePage(PageNum)
{
	window.location.href='?page_no='+PageNum;
}
function PriPage()
{
	var PageNum='<% = cint(page_no) - 1 %>';
	ChangePage(PageNum);
}
function NextPage()
{
	var PageNum='<% = cint(page_no) + 1 %>';
	ChangePage(PageNum);
}
</script> 
<%
set conn=nothing
RsIPObj.close
set RsIPObj=nothing
%>





