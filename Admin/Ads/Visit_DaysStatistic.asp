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
	strShowErr = "<li>��������!</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
Else
	If Isnumeric(AdID)=False Then
		strShowErr = "<li>��������!</li>"
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
<title>������Ϣͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td>������[<font color="red"><%=AdName%></font>]�� ������Ϣͳ��  |  <a href="javascript:history.back();">������һ��</a></td>
  </tr>
</table>
<%
Dim RsDayObj,Sql
Set RsDayObj = Server.CreateObject(G_FS_RS)
Dim MaxVisitCount,VisitTime,VisitDay,CurrentDay
CurrentDay=Day(Now())
Dim DaysOfMonth
if Month(Now())=1 then
	DaysOfMonth=GetDayNum(Year(Now())-1, 12)
else
	DaysOfMonth=GetDayNum(Year(Now()), Month(Now())-1)
end if
If G_IS_SQL_DB=0 then
	Sql="Select VisitTime From FS_AD_Source where DATEDIFF('d',VisitTime,Now()) <= " & DaysOfMonth & " And DATEDIFF('d',VisitTime,Now()) >=0 and AdID="&CintStr(AdID)&""
Else
	Sql="Select VisitTime From FS_AD_Source where DATEDIFF(day,VisitTime,GetDate()) <= " & DaysOfMonth & " And DATEDIFF(day,VisitTime,GetDate()) >=0 and AdID="&CintStr(AdID)&""
End If
RsDayObj.Open Sql,Conn,3,3
MaxVisitCount=0
Dim VisitNum(31),I
for I=1 To DaysOfMonth
	VisitNum(I)=0
next
Do While not RsDayObj.Eof 
	VisitTime = RsDayObj("VisitTime")
	VisitDay = Day(VisitTime)
	for I=1 To DaysOfMonth
		if I=VisitDay then
			VisitNum(I)=VisitNum(I)+1
		end if
	next
	RsDayObj.MoveNext
Loop
for I=1 To DaysOfMonth
	if VisitNum(I)>=MaxVisitCount  then
		MaxVisitCount=VisitNum(I)
	end if
next
%>
<% 
	Dim VisitSize(31)
	For I=1 To DaysOfMonth
	if MaxVisitCount<>0 then
	VisitSize(I)=100*VisitNum(I)/MaxVisitCount
	else
	VisitSize(I)=0
	end if
	Next
%>
<table border=0 align="center" cellpadding=2 >
	<tr>
		<td align=center>���<% =DaysOfMonth %>��ͳ��ͼ��</td>
	</tr>
	<tr>
		
    <td align=center><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap> 
                  <% Response.Write(MaxVisitCount&"��")%>
               </td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>
				<% if MaxVisitCount>3 then  
					Response.Write(Round(MaxVisitCount*0.75)&"��") 
					elseif MaxVisitCount>1  then 
					Response.Write((MaxVisitCount-1)&"��") 
					else Response.Write("&nbsp;") 
					end if
				%> 
				</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>
				<% if MaxVisitCount>3 then  
					Response.Write(Round(MaxVisitCount*0.5)&"��") 
					elseif MaxVisitCount>2  then 
					Response.Write((MaxVisitCount-2)&"��") 
					else Response.Write("&nbsp;") 
					end if
				%>
				</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>
                <% if MaxVisitCount>3 then 
					 Response.Write(Round(MaxVisitCount*0.25)&"��") 
					else Response.Write("&nbsp;") 
					end if
				%>
				</td>
              </tr>
              <tr> 
                <td  height="31" valign=top  align="right" nowrap>0��</td>
              </tr>
            </table></td>
          <td valign="bottom">
<table align=center>
              <tr valign=bottom > 
                <% For I=CurrentDay+1 To DaysOfMonth %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=/admin/Images/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
                <% For I=1 To CurrentDay  %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=/admin/Images/bar.gif width="15" height="<% =VisitSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
                <td>��λ���գ�</td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
	</tr>
</table>
<p>��</p>
<%
Dim RsDaysObj
Set RsDaysObj = Server.CreateObject(G_FS_RS)
Sql="Select VisitTime From FS_AD_Source Where AdID="&CintStr(AdID)&""
RsDaysObj.Open Sql,Conn,3,3
MaxVisitCount=0
for I=1 To 30
	VisitNum(I)=0
next
Do While not RsDaysObj.Eof 
	VisitTime = RsDaysObj("VisitTime")
	VisitDay = Day(VisitTime)
	for I=1 To 31
		if I=VisitDay then
			VisitNum(I)=VisitNum(I)+1
		end if
	next
	RsDaysObj.MoveNext
Loop
for I=1 To 31
	if VisitNum(I)>=MaxVisitCount  then
		MaxVisitCount=VisitNum(I)
	end if
next
%>
<% 
	Dim VisitsSize(31),AllCount
	AllCount=0
	For I=1 To 31
		Allcount=AllCount+VisitNum(I)
	Next
	For I=0 To 31
		if VisitNum(I)<>0 then
			VisitsSize(I)=100*VisitNum(I)/AllCount
		else
			VisitsSize(I)=0
		end if
	Next
%>
<table border=0 align="center" cellpadding=2>
	<tr>
		<td align=center>�������������ͼ��</td>
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
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=/admin/Images/bar.gif width="15" height="100" id=htav><br>
                  ��</td>
                <% For I=1 To 31 %>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=/admin/Images/bar.gif width="15" height="<% =VisitsSize(I) %>" id=htav><br>
                  <% = I%>
                </td>
                <% Next %>
                <td>��λ���գ�</td>
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
Function GetDayNum(YearVar, MonthVar)
    Dim Temp, LeapYear, BigMonthArray, i, BigMonth
    BigMonthArray = Array(1, 3, 5, 7, 8, 10, 12)
    YearVar = CInt(YearVar)
    MonthVar = CInt(MonthVar)
    Temp = CInt(YearVar / 4)
    If YearVar = Temp * 4 Then
        LeapYear = True
    Else
        LeapYear = False
    End If
    For i = LBound(BigMonthArray) To UBound(BigMonthArray)
        If MonthVar = BigMonthArray(i) Then
            BigMonth = True
            Exit For
        Else
            BigMonth = False
        End If
    Next
    If BigMonth = True Then
        GetDayNum = 31
    Else
        If MonthVar = 2 Then
            If LeapYear = True Then
                GetDayNum = 29
            Else
                GetDayNum = 28
            End If
        Else
            GetDayNum = 30
        End If
    End If
End Function
Set Conn=Nothing
%><!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





