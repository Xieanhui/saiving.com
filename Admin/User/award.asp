<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
'on error resume next
Dim Conn,User_Conn,awardRs,prizeIDs,EndDate,awardUser,AwardID,ArrayIndex,AwardsUserArray,UserInfoRs,AwardsUserRs
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
int_RPP=20 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"				'尾页

MF_Default_Conn
MF_User_Conn
MF_Session_TF
Set awardRs=Server.CreateObject(G_FS_RS)
awardRs.open "select AwardID,AwardName,AwardPic,StartDate,EndDate,PrizeIDS,Opened from FS_ME_award",User_Conn,1,1
'积分抽奖
function activeAward()
	Dim active_TF_Rs,sql_cmd,activeTF
	activeTF=false
	sql_cmd="select AwardID from FS_ME_award where opened=0"
	Set active_TF_Rs=User_Conn.execute(sql_cmd)
	if not active_TF_Rs.eof then
		activeTF=true
	End if
	activeAward=activeTF
	active_TF_Rs.close
	set active_TF_Rs=Nothing
End function
'积分兑换
Function activeAwardPoint
	Dim active_TF_Rs,sql_cmd,activeTF
	activeTF=false
	if  G_IS_SQL_DB=0 then
		sql_cmd="select AID from FS_ME_AnswerForPoint where DateDiff(d,Convert(nvarchar(10),AEndDate,120),#"&DateValue(Now)&"#)>0"
	Else
		sql_cmd="select AID from FS_ME_AnswerForPoint where DateDiff('d',Convert(nvarchar(10),AEndDate,120),'#"&DateValue(Now)&"#')>0"
	End if
	Set active_TF_Rs=User_Conn.execute(sql_cmd)
	if not active_TF_Rs then
		activeTF=true
	End if
	activeAwardPoint=activeTF
	active_TF_Rs.close
	set active_TF_Rs=nothing
End function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td class="xingmu">抽奖管理</td>
  </tr>
  <tr>
    <td class="hback">积分抽奖&nbsp;|&nbsp;<a href="ChangePrize.asp">积分兑换</a> | <a href="AnswerForPoint.asp">积分竞答</a> 
      | <a href="#" onClick="history.back()">后退</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form action="awardAction.asp?act=delete" method="post" name="awardForm" id="awardForm">
    <tr class="xingmu"> 
      <td width="16%" align="center">主题</td>
      <td width="30%" align="center">中奖会员</td>
      <td width="11%" align="center">状态</td>
      <td width="17%" align="center">截止时间</td>
      <td width="14%" align="center"><input type="checkbox" name="Delete_CheckAll" value="all" onClick="CheckAll(this,'DeleteAwards')"></td>
    </tr>
    <%
			if not awardRs.eof then
				awardRs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("Page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo>awardRs.PageCount Then cPageNo=awardRs.PageCount 
				If cPageNo<=0 Then cPageNo=1
				awardRs.AbsolutePage=cPageNo
			end if
			for i=0 to int_RPP
				if awardRs.eof then exit for
				Response.Write("<tr class='hback'>"&chr(10)&chr(13)) 
				Response.Write("<td width='30%' align='center'><a href='award_AddEdit.asp?act=edit&awardid="&awardRs("awardID")&"'>"&awardRs("awardName")&"</a></td>"&chr(10)&chr(13))
				Response.Write("<td width='20%'>")
				Response.Write("<select name='Grade_"&awardRs("AwardID")&"' onchange=""getAwardUser('"&awardRs("AwardID")&"',this.value)"">")
				prizeIDs=split(awardRs("prizeIDs"),",")
				for ArrayIndex=0 to Ubound(prizeIDs)
					Response.Write("<option value='"&prizeIDs(ArrayIndex)&"'>"&(ArrayIndex+1)&"等奖</option>"&chr(10)&chr(13))
				next
				Response.Write("</select>")
				Response.Write(" | <span id='PrizeUsers_"&awardRs("awardID")&"'>"&chr(10)&chr(13))
				Response.Write("<select name='AwardUsers_"&AwardID&"'>"&chr(10)&chr(13))
				Set AwardsUserRs=User_Conn.execute("Select UserNumber,winner From FS_ME_User_Prize where PrizeID="&CintStr(prizeIDs(0))&" And awardID="&awardRs("awardID")&" and winner=1")
				if not AwardsUserRs.eof then
					while not AwardsUserRs.eof  
						Set UserInfoRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&AwardsUserRs("UserNumber")&"'")
						Response.Write("<option value='"&AwardsUserRs("UserNumber")&"'>"&UserInfoRs("UserName")&"</option>"&Chr(10)&Chr(13))
						AwardsUserRs.movenext
					Wend
				ELse
					Response.Write("<option value='-1'>暂无中奖</option>"&Chr(10)&Chr(13))
				End if
				AwardsUserRs.close
				Set AwardsUserRs=nothing
				Set UserInfoRs=nothing
				Response.Write("</span>")
				Response.Write("</td>")
				EndDate=awardRs("EndDate")
				if awardRs("Opened")=1 then
					Response.Write("<td align='center'>已过期</td>")
				Elseif DateValue(EndDate)=DateValue(Now()) Then
					Response.Write("<td align='center'><button onClick=""openAward("&awardRs("AwardID")&")"">开  奖</button></td>")
				ElseIf DateValue(EndDate)<DateValue(Now()) And awardRs("Opened")=0 then
					Response.Write("<td align='center'><button onClick=""openAward("&awardRs("AwardID")&")"">已过期，请开奖</button>")
				Else
					Response.Write("<td align='center'><font color=""red"">未过期</font></td>")
				end If
				Response.Write("<td align='center'>"&EndDate&"</td>")
				Response.Write("<td align='center'><input type='checkbox' name='DeleteAwards' value='"&awardRs("AwardID")&"'></td>")
				Response.Write("</tr>")
				awardRs.movenext
			next
		%>
  </form>
  <tr> 
    <td align="right" colspan="6" class="hback">
	<%
		Dim displayTF
		if activeAward then
			displayTF="disabled"
		End if
	%>
	<input name="AddAward" type="button" value="添 加" onClick="location='award_AddEdit.asp?act=add'" <%=displayTF%>> 
      <input type="Button" name="deleteAward" onClick="AlertBeforeSubmite('DeleteAwards')" value="删 除"> 
      <%
	response.Write fPageCount(awardRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	%>
    </td>
  </tr>
</table> 
</body>
<%
if Request.QueryString("Act")="addGroup" then
	AddGroupRs.close
	set AddGroupRs=nothing
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
end if
%>
<script language="JavaScript" type="text/JavaScript">
var request=true;
var result;
try
{
	request=new XMLHttpRequest();
}catch(trymicrosoft)
{
try
{
	request=new ActiveXObject("Msxml2.XMLHTTP")
}catch(othermicrosoft)
{
try
{
	request=new ActiveXObject("Microsoft.XMLHTTP")
}catch(filed)
{
	request=false;
}
}
}
if(!request) alert("Error initializing XMLHttpRequest!");
function getAwardUser(Obj1,Obj2)
{
	var url="getAwardUser.asp?AwardId="+Obj1+"&PrizeID="+Obj2+"&r="+Math.random();//构造url
	request.open("GET",url,true);//建立连接
	request.onreadystatechange = getResult;
	request.send(null);//传送数据，因为数据通过url传递了，所以这里传递的是nulla
}
function getResult(Obj)//当服务器响应的时候就使用这个方法
{
	if(request.readyState ==4)//根据HTTP 就绪状态判断响应是否完成
	{
		if(request.status == 200)//判断请求是否成功
		{
			result=request.responseText;//获得响应的结果，也就是新的<select>
			var contaner=result.substring(0,result.indexOf("*"));
			var selectContent=result.substring(result.indexOf("*")+1,result.length);
			document.getElementById(contaner).innerHTML=selectContent;

		}
	}
}
function CheckAll(Obj,TargetName)
{
	var CheckBoxArray;
	CheckBoxArray=document.getElementsByName(TargetName);
	for(var i=0;i<CheckBoxArray.length;i++)
	{
		if(Obj.checked)
		{
			CheckBoxArray[i].checked=true;
		}
		else
		{
			CheckBoxArray[i].checked=false;
		}
	}
}
function AlertBeforeSubmite(TargetName)
{
	var checkGroup=document.getElementsByName(TargetName);
	var flag=false;
	for(var i=0;i<checkGroup.length;i++)
	{
		if(checkGroup[i].checked)
		{
			flag=true;
		}
	}
	if(flag)
	{
		if(confirm("确认要删除该记录?该操作将会删除用户中奖记录！"))
		{
			document.awardForm.submit();
		}
	}
	else
	{
		alert("请选择要删除的记录")
	}
}
function openAward(awardID)
{
	location="awardAction.asp?act=open&awardID="+awardID+"&rnd="+Math.random();
}
</script>
</html>






