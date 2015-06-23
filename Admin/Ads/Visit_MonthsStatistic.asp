<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<%
Dim AdID,strShowErr,AdName
AdID=Request.QueryString("ID")
AdName=Request.QueryString("AdName")
If AdID="" or IsNull(AdID) Then
	strShowErr = "<li>参数错误!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
Else
	If Isnumeric(AdID)=False Then
		strShowErr = "<li>参数错误!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
	AdID=Clng(AdID)
End If
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("AS_site") then Err_Show
if not MF_Check_Pop_TF("AS002") then Err_Show
%>
<html>
<head>
<title>按月信息统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td>以下是[<font color="red"><%=AdName%></font>]的 按月信息统计  |  <a href="javascript:history.back();">返回上一级</a></td>
  </tr>
</table>
<%
Dim RsMonthObj,Sql
Set RsMonthObj = Server.CreateObject(G_FS_RS)
Dim MaxVisitCount,VisitTime,VisitMonth,CurrentMonth
CurrentMonth=Month(Now())
If G_IS_SQL_DB=0 then
	Sql="Select VisitTime From FS_AD_Source where DATEDIFF('m',VisitTime,Now()) < 12 And DATEDIFF('m',VisitTime,Now()) >=0 and AdID="&AdID&""
Else
	Sql="Select VisitTime From FS_AD_Source where DATEDIFF(month,VisitTime,GetDate()) < 12 And DATEDIFF(month,VisitTime,GetDate()) >=0 and AdID="&AdID&""
End If
RsMonthObj.Open Sql,Conn,3,3
MaxVisitCount=0
Dim VisitNum(12),I
for I=1 To 12
	VisitNum(I)=0
next
Do While not RsMonthObj.Eof 
	VisitTime = RsMonthObj("VisitTime")
	VisitMonth = Month(VisitTime)
	for I=1 To 12
		if I=VisitMonth then
			VisitNum(I)=VisitNum(I)+1
		end if
	next
	RsMonthObj.MoveNext
Loop
for I=1 To 12
	if VisitNum(I)>=MaxVisitCount  then
		MaxVisitCount=VisitNum(I)
	end if
next
%>
<% 
	Dim VisitSize(12)
	For I=1 To 12
	if MaxVisitCount<>0 then
	VisitSize(I)=100*VisitNum(I)/MaxVisitCount
	else
	VisitSize(I)=0
	end if
	Next
%>
<table border=0 align="center" cellpadding=2>
	<tr>
		<td align=center>最近12个月统计图表</td>
	</tr>
	<tr>
		
    <td align=center><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap> 
                  <% Response.Write(MaxVisitCount&"次")%>
               </td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>
				<% if MaxVisitCount>3 then  
					Response.Write(Round(MaxVisitCount*0.75)&"次") 
					elseif MaxVisitCount>1  then 
					Response.Write((MaxVisitCount-1)&"次") 
					else Response.Write("&nbsp;") 
					end if
				%> 
				</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>
				<% if MaxVisitCount>3 then  
					Response.Write(Round(MaxVisitCount*0.5)&"次") 
					elseif MaxVisitCount>2  then 
					Response.Write((MaxVisitCount-2)&"次") 
					else Response.Write("&nbsp;") 
					end if
				%>
				</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>
                <% if MaxVisitCount>3 then 
					 Response.Write(Round(MaxVisitCount*0.25)&"次") 
					else Response.Write("&nbsp;") 
					end if
				%>
				</td>
              </tr>
              <tr> 
                <td  height="31" valign=top  align="right" nowrap>0次</td>
              </tr>
            </table></td>
          <td valign="bottom">
<table align=center>
              <tr valign=bottom > 
			  <% if CurrentMonth =12 then %>
                <% For I=1 To 12 %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=/admin/Images/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
			<% else %>
                <% For I=CurrentMonth+1 To 12 %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=/admin/Images/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
                <% For I=1 To CurrentMonth  %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=/admin/Images/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
			<% end if %>
                <td>单位（月）</td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
	</tr>
</table>
<p>　</p>
<%
Dim RsMonthsObj
Set RsMonthsObj = Server.CreateObject(G_FS_RS)
Sql="Select VisitTime From FS_AD_Source where AdID="&AdID&""
RsMonthsObj.Open Sql,Conn,3,3
MaxVisitCount=0
for I=1 To 12
	VisitNum(I)=0
next
Do While not RsMonthsObj.Eof 
	VisitTime = RsMonthsObj("VisitTime")
	VisitMonth = Month(VisitTime)
	for I=1 To 12
		if I=VisitMonth then
			VisitNum(I)=VisitNum(I)+1
		end if
	next
	RsMonthsObj.MoveNext
Loop
for I=1 To 12
	if VisitNum(I)>=MaxVisitCount  then
		MaxVisitCount=VisitNum(I)
	end if
next
%>
<% 
	Dim VisitsSize(12),AllCount
	AllCount=0
	For I=1 To 12
		Allcount=AllCount+VisitNum(I)
	Next
	For I=0 To 12
		if VisitNum(I)<>0 then
			VisitsSize(I)=100*VisitNum(I)/AllCount
		else
			VisitsSize(I)=0
		end if
	Next
%>
<table border=0 align="center" cellpadding=2>
	<tr>
		<td align=center>访问量12个月分配图表</td>
	</tr>
	<tr>
		
    <td align=center><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap>100%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>75%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>50%</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>25%</td>
              </tr>
              <tr> 
                <td  height="31" valign=top  align="right" nowrap>0</td>
              </tr>
            </table></td>
          <td valign="bottom">
<table align=center>
              <tr valign=bottom >
                <td width="15" align=center nowrap background=../images/tu_back.gif><img src=/admin/Images/bar.gif width="15" height="100" id=htav><br>
                  总</td>
                <% For I=1 To 12 %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=/admin/Images/bar.gif width="15" height="<% =VisitsSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
                <td>单位（月）</td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
	</tr>
</table>
</body>
</html>
<%
Conn.Close
Set Conn=Nothing
%><!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





