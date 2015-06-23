<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/WaterPrint_Function.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<%
Dim Conn,CollectConn
MF_Default_Conn
MF_Collect_Conn
MF_Session_TF
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.Charset="gb2312"
If Not MF_Check_Pop_TF("CS_collect") Then Err_Show
Dim Str_act
Str_act=Request.QueryString("Action")
Select Case Str_act
	Case "GetTaskList"
		Response.Write(GetTaskList())
	Case Else
		Main()
End Select
Function Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>定时采集监视器</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Prototype.js"></script>
<script language="JavaScript">
<!--
var TaskList=new Array();
var Autoget;
function OnLoad()
{
	autoGetTask();
	Autoget = setInterval("autoGetTask();",60000);
}
function GetThisTask(ObjArr)
{
	var ForNum1,ForNum2,TaskLen1,FindSite
	for (ForNum1=0; ForNum1<ObjArr.length; ForNum1++)
	{
		TaskLen1=TaskList.length;
		FindSite=-1;
		for (ForNum2=0; ForNum2<TaskLen1; ForNum2++)
		{
			if (TaskList[ForNum2][0]==ObjArr[ForNum1][0])
			{
				FindSite=ForNum2;
				break;
			}
		}
		if (FindSite>-1)
		{
			TaskList[FindSite]=ObjArr[ForNum1];
		}
		else
		{
			TaskList.push(new Array());
			for (ForNum2=0; ForNum2<ObjArr[ForNum1].length; ForNum2++)
			{
				TaskList[TaskLen1].push(ObjArr[ForNum1][ForNum2]);
			}
		}
	}
	UpdateList();
	Collection();
}
function Collection()
{
	for (var i=0; i<TaskList.length; i++)
	{
		if (TaskList[i][4]==0)
		{
			PrintLog("站点："+TaskList[i][1]+"在"+Date()+"开始采集。");
			GetInfo("AutoCollect_Action.asp?SiteID="+TaskList[i][0]+"&Num="+TaskList[i][3],"GET","",CollectionR);
		}
	}
}
function CollectionR(XMLH)
{
	var TRecv=XMLH.responseText;
	//alert(TRecv);
	if (TRecv.indexOf("||")>0)
	{
		TRecv=TRecv.split("||");
		switch (TRecv[0])
		{
			case "Next":
			GetInfo("AutoCollect_Action.asp?"+TRecv[2],"GET","",CollectionR);
			PrintLog("站点："+FindSiteName(TRecv[1])+"在"+Date()+"采集中。");
			break;
			case "End":
			SetStat(TRecv[1],1);
			PrintLog("站点："+FindSiteName(TRecv[1])+"在"+Date()+"采集完成。");
			break;
			case "Err":
			PrintLog("站点："+FindSiteName(TRecv[1])+"在"+Date()+"采集失败。");
			SetStat(TRecv[1],2);
			break;
		}
	}
	UpdateList();
}
function FindSiteName(SiteID)
{
	for (var i=0; i<TaskList.length; i++)
	{
		if (TaskList[i][0]==SiteID)
		{
			return TaskList[i][1];
		}
	}
}
function SetStat(SiteID,Stat)
{
	for (var i=0; i<TaskList.length; i++)
	{
		if (TaskList[i][0]==SiteID)
		{
			TaskList[i][4]=Stat;
			return true;
		}
	}
}
function PrintLog(Info)
{
	$("TaskLog").innerHTML+=Info+"<br />";
}

function UpdateList()
{
	var HtmlStr="",statstr="";
	for (var i=0; i<TaskList.length; i++)
	{
		switch (TaskList[i][4])
		{
			case 0:
			statstr="等待";
			break;
			case 1:
			statstr="完成";
			break;
			default :
			statstr="失败";
		}
		HtmlStr+="<div>站点："+TaskList[i][1]+"&nbsp;&nbsp;采集时间："+TaskList[i][2]+"&nbsp;&nbsp;状态："+statstr+"</div>";
	}
	if (HtmlStr=="")
		HtmlStr="&nbsp;";
	$("AllTaskList").innerHTML=HtmlStr;
}
function autoGetTask()
{
	GetInfo("?Action=GetTaskList","GET","",GetInfo_Receive);
}
function GetInfo(url,_method,_parameters,_onComplete){
	var myAjax = new Ajax.Request(
		url,
		{method:_method,
		parameters:_parameters,
		onComplete:_onComplete
		}
		);
}
function GetInfo_Receive(OriginalRequest){
	var Info="";
	var Arr_Info="";
	Info=OriginalRequest.responseText;
	if (Info.length>5)
	{
		Arr_Info=eval(Info);
	}else{
		Arr_Info=new Array();
	}
	GetThisTask(Arr_Info);
}
window.onload=OnLoad;
//-->
</script>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="0" class="table">
	<tr class="hback">
		<td>定时采集监视器</td>
	</tr>
</table>

<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<tr class="xingmu">
		<td  class="xingmu">任务列表</td>
		<td  class="xingmu">采集日志</td>
	</tr>
	<tr>
		<td id="AllTaskList" valign="top" class="hback">&nbsp;</td>
		<td id="TaskLog" valign="top" class="hback"></td>
	</tr>
</table>
</body>
</html>
<%
End Function

Function GetTaskList()
	Dim RsCollect,JsArray,LoopStr
	JsArray=""
	Set RsCollect=CollectConn.Execute("Select ID,SiteName,AutoCellectTime,CellectNewNum,LinkHeadSetting,LinkFootSetting,PagebodyHeadSetting,PagebodyFootSetting,PageTitleHeadSetting,PageTitleFootSetting From FS_Site Where ID > 0 And IsLock = 0 And AutoCellectTime <> 'no'")
	While Not RsCollect.eof
		If RsCollect("LinkHeadSetting") <> "" And RsCollect("LinkFootSetting") <> "" And RsCollect("PagebodyHeadSetting") <> "" And  RsCollect("PagebodyFootSetting") <> "" And  RsCollect("PageTitleHeadSetting") <> "" And  RsCollect("PageTitleFootSetting") <> "" Then
			LoopStr=GetAutoCS(RsCollect("AutoCellectTime"))
			If LoopStr<>False Then
				If JsArray="" Then
					JsArray="["&RsCollect("ID")&",'"&RsCollect("SiteName")&"','"&LoopStr&"',"&RsCollect("CellectNewNum")&",0]"
				Else
					JsArray=JsArray&",["&RsCollect("ID")&",'"&RsCollect("SiteName")&"','"&LoopStr&"',"&RsCollect("CellectNewNum")&",0]"
				End If
			End If
		End If
		RsCollect.movenext
	Wend
	RsCollect.close()
	Set RsCollect=Nothing
	If JsArray="" Then
		JsArray="new Array()"
	Else
		JsArray="["&JsArray&"]"
	End If
	GetTaskList = JsArray
End Function
Function GetAutoCS(TimeStr)
	Dim Mode,todytask,monthStr,thisDate,thisTime
	thisDate=Date()
	thisTime=Time
	thisTime=left(thisTime,InStrRev(thisTime,":",-1,0)-1)
	todytask=False
	If Not InStr(TimeStr,"$$$") > 0 Then
		GetAutoCS=False
		Exit Function
	End If
	Mode=Split(TimeStr,"$$$")
	Select Case mode(0)
		Case "day"
			If Left(thisTime,5)=mode(1) Then
				todytask=mode(1)
			Else
				todytask=False
			End If
		Case "week"
			weekStr=Split(mode(1),"|")
			If WeekDay(thisDate,2)=weekStr(0) Then
				If Left(thisTime,5)=weekStr(1) Then
					todytask=weekStr(1)
				Else
					todytask=False
				End If
			Else
				todytask=False
			End If
		Case "month"
			monthStr=Split(mode(1),"|")
			If Day(thisDate)=monthStr(0) Then
				If Left(thisTime,5)=monthStr(1) Then
					todytask=monthStr(1)
				Else
					todytask=False
				End If
			Else
				todytask=False
			End If
		Case Else
			todytask=False
	End Select
	GetAutoCS=todytask
End Function

%>






