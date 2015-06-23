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
<title>24小时信息统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td height="28" class="xingmu"><div align="left"><strong>24小时信息统计</strong></div></td>
  </tr>
  <tr>
    <td height="28" class="hback">
      <%
Dim RsHourObj,Sql
Set RsHourObj = Server.CreateObject(G_FS_RS)
Dim MaxVisitCount,VisitTime,VisitHour,CurrentHour
CurrentHour=hour(now())
If G_IS_SQL_DB=0 then
	Sql="Select VisitTime From FS_SS_Stat where DATEDIFF('h',VisitTime,Now()) < 24 And DATEDIFF('h',VisitTime,Now()) >=0 "
Else
	Sql="Select VisitTime From FS_SS_Stat where DATEDIFF(hour,VisitTime,GetDate()) < 24 And DATEDIFF(hour,VisitTime,GetDate()) >=0 "
End If
RsHourObj.Open Sql,Conn,3,3
MaxVisitCount=0
Dim VisitCount(23),I
for I=0 To 23
	VisitCount(I)=0
next
Do While not RsHourObj.Eof 
	VisitTime = RsHourObj("VisitTime")
	VisitHour = Hour(VisitTime)
	for I=0 To 23
		if I=VisitHour then
			VisitCount(I)=VisitCount(I)+1
		end if
	next
	RsHourObj.MoveNext
Loop
for I=0 To 23
	if VisitCount(I)>=MaxVisitCount  then
		MaxVisitCount=VisitCount(I)
	end if
next
%>
      <% 
	Dim VisitSize(23)
	For I=0 To 23
	if MaxVisitCount<>0 then
	VisitSize(I)=100*VisitCount(I)/MaxVisitCount
	else
	VisitSize(I)=0
	end if
	Next
%>
      <table border=0 align="center" cellpadding=2>
        <tr> 
          <td align=center>最近24小时统计图表</td>
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
                <td valign="bottom"> <table align=center>
                    <tr valign=bottom > 
                      <% For I=CurrentHour+1 To 23 %>
                      <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../Images//bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br> 
                        <% = I%>
                      </td>
                      <% Next %>
                      <% For I=0 To CurrentHour %>
                      <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../Images/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br> 
                        <% = I%>
                      </td>
                      <% Next %>
                      <td>单位（点）</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <%
Dim RsHoursObj
Set RsHoursObj = Server.CreateObject(G_FS_RS)
Sql="Select VisitTime From FS_SS_Stat"
RsHoursObj.Open Sql,Conn,3,3
MaxVisitCount=0
for I=0 To 23
	VisitCount(I)=0
next
Do While not RsHoursObj.Eof 
	VisitTime = RsHoursObj("VisitTime")
	VisitHour = Hour(VisitTime)
	for I=0 To 23
		if I=VisitHour then
			VisitCount(I)=VisitCount(I)+1
		end if
	next
	RsHoursObj.MoveNext
Loop
for I=0 To 23
	if VisitCount(I)>=MaxVisitCount  then
		MaxVisitCount=VisitCount(I)
	end if
next
%>
      <% 
	Dim VisitsSize(23),AllCount
	AllCount=0
	For I=0 To 23
		Allcount=AllCount+VisitCount(I)
	Next
	For I=0 To 23
		if VisitCount(I)<>0 then
			VisitsSize(I)=100*VisitCount(I)/AllCount
		else
			VisitsSize(I)=0
		end if
	Next
%>
      <table border=0 align="center" cellpadding=2>
        <tr> 
          <td align=center>访问量24小时分配图表</td>
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
                <td valign="bottom"> <table align=center>
                    <tr valign=bottom > 
                      <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../Images/bar.gif width="15" height="100" id=htav><br>
                        总</td>
                      <% For I=0 To 23 %>
                      <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../Images/bar.gif width="15" height="<% =VisitsSize(I) %>" id=htav><br> 
                        <% = I%>
                      </td>
                      <% Next %>
                      <td>单位（点）</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<%
Conn.Close
Set Conn=Nothing
%>






