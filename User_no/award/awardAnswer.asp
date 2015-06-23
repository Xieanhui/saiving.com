<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="lib/cls_award.asp"-->
<%
'no cache
response.expires=0 
response.addHeader "pragma" , "no-cache" 
response.addHeader "cache-control" , "private" 
'-------------------------------------------
Dim AnswerRs,awardObj,currentDate,activeTF
activeTF=false
currentDate=DateValue(Now())
if G_IS_SQL_DB=0 then
	Randomize
	Set AnswerRs=User_Conn.execute("Select top 1 AID from FS_ME_AnswerForPoint where datediff('d',dateValue(AstartDate),'"&NoSqlHack(currentDate)&"')>=0 And datediff('d',dateValue(AEndDate),'"&NoSqlHack(currentDate)&"')<0  order by Rnd(-(AID+"&Rnd()&"))")
ELse
	Set AnswerRs=User_Conn.execute("Select top 1 AID from FS_ME_AnswerForPoint where datediff(d,convert(nvarchar(10),AstartDate,120),'"&NoSqlHack(currentDate)&"')>=0 And datediff(d,convert(nvarchar(10),AEndDate,120),'"&NoSqlHack(currentDate)&"')<0 ORDER BY NEWID()")
End if
if not AnswerRs.eof then
	Set awardObj=New cls_award
	activeTF=true
End if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=GetUserSystemTitle%></title>
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<script language="javascript" src="../../FS_Inc/ProtoType.js"></script>
</head>
<body class="hback">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
	<tr>
		<td class="xingmu" height="20"><img src="../images/award.gif" alt="积分抽奖" border="0">积分问答 </td>
	</tr>
	<%if not activeTF then%>
	<tr>
		<td class="hback"><img src="../images/alert.gif" />暂无问答活动</td>
	</tr>
	<%Else%>
	<tr>
	  <td class="hback">
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
		<tr>
		<td>
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
				<%
					Dim UserRs,Integral'当前会员积分,
					Dim onceMoreTF,Rs
					onceMoreTF=false
					'获得当前会员积分--------------------------------
					Set UserRs=User_Conn.execute("Select Integral from FS_ME_Users where UserNumber='"&session("FS_UserNumber")&"'")
					if not UserRs.eof then
						Integral=UserRs("Integral")
					Else
						Integral=0
					End if
					UserRs.close
					set UserRs=nothing
					'------------------------------------------------
					
					Dim PrizeArray,i,action,answerArray
					if not AnswerRs.eof then
						Set Rs=User_Conn.execute("Select ID From FS_ME_Answer_User where questionID="&AnswerRs("AID")&" And usernumber='"&session("FS_UserNumber")&"'")
						if not Rs.eof then
							onceMoreTF=true
						End if
						awardObj.getAnswerForPoint(AnswerRs("AID"))
							Response.Write("<tr>"&vbcrlf)
						if Clng(Integral)<Clng(awardObj.answer_NeedPoint) then
							action="<img src=""../images/alert.gif""/><font color=""red"">积分不足</font>"
						Else
							action="<a href=""#"" onClick=""changeAward("&AnswerRs("AID")&","&awardObj.answer_NeedPoint&")""><img src=""../images/bottomduihuan.gif"" border=""0"" alt=""兑换奖品""/></a>"
						End if
						answerArray=split(awardObj.AnswerIDS,",")
						Response.Write("<td class=""hback"">"&vbcrlf)
						Response.Write("<table border=""0""  cellpadding=""1"" cellspacing=""1"" width=""100%"" class=""table"">"&vbcrlf)
						Response.Write("<tr>"&vbcrlf)
							Response.Write("<td colspan=""2"" class=""hback""><img src=""../images/question.gif""/><strong>"&awardObj.ATopic&"</strong>&nbsp;|需要积分：<img src=""../images/moneyOrPoint.gif""/>"&awardObj.answer_NeedPoint&"&nbsp;|奖励积分：<img src=""../images/moneyOrPoint.gif""/><font color=""red"">"&awardObj.PrizePoint&"</font></td>")
						Response.Write("</tr>"&vbcrlf)
						Response.Write("<tr>"&vbcrlf)
						Response.Write("<td width=""10%"" class=""hback""><img src="""&awardObj.APic&""" width=""80"" height=""80"" border=""0""></td>")
						Response.Write("<td valign=""top"" class=""hback"">"&awardObj.ADesc&"</td>")
						Response.Write("</tr>"&vbcrlf)
						for i=0 to Ubound(answerArray)
							awardObj.getAnswer(answerArray(i))
							Response.Write("<tr>"&vbcrlf)
								Response.Write("<td colspan=""2"" class=""hback"" id=""td_"&answerArray(i)&""" onClick=""checkIT($('td_"&answerArray(i)&"'))""><span onClick=""emptyClick()""><input type=""radio"" name=""answer_"&AnswerRs("AID")&""" value="""&answerArray(i)&"""></span><img src=""../images/answer.gif""/>"&awardObj.AnswerDesc&"</td>")
							Response.Write("</tr>"&vbcrlf)
						next
						Response.Write("<tr>"&vbcrlf)
						if Clng(Integral)<Clng(awardObj.answer_NeedPoint) then
							Response.Write("<td colspan=""2"" class=""hback"">&nbsp;&nbsp;<img src=""../images/alert.gif""/><font color=""red"">积分不足</font></td>")
						Elseif onceMoreTF Then
							Response.Write("<td colspan=""2"" class=""hback"">&nbsp;&nbsp;<img src=""../images/alert.gif""/><font color=""red"">已参与过</font></td>")
						Else
							Response.Write("<td colspan=""2"" class=""hback"">&nbsp;&nbsp;&nbsp;&nbsp;<button onClick=""makeAnswer("&AnswerRs("AID")&","&awardObj.answer_NeedPoint&")"">提  交 答 案</button></td>")
						End if
						Response.Write("</tr>"&vbcrlf)
						Response.Write("</table>"&vbcrlf)
						Response.Write("</td>")
						Response.Write("</tr>"&vbcrlf)
					End if
					AnswerRs.close
					set AnswerRs=nothing
				%>
			</table>
		</td>
		</tr>
	  </table>
	  </td>
  </tr>
	<%End if%>
</table>
</body>
</html>
<script language="javascript">
<!--
var tf=true;;
function emptyClick()
{
	tf=false;
}
function checkIT(Obj)
{
	var answer=Obj.firstChild.firstChild
	if(tf)
	{
		answer.checked=!answer.checked
	}
	tf=true;
}
function makeAnswer(answerID,Integral)
{
	var url="awardAction.asp"
	var element=document.all("answer_"+answerID)
	var questionID=0;
	for(var i=0;i<element.length;i++)
	{
		if(element[i].checked)
		{
			questionID=element[i].value;
			break;
		}
	}
	var pars="action=answer&answerID="+answerID+"&questionID="+questionID+"&rnd="+Math.random();
	if(confirm("确定要进行该操作\n该操作将消费积分："+Integral))
	{
		 var myAjax = new Ajax.Request(url,{method: 'get', parameters: pars, onComplete: showResponse});

	}
	function showResponse(originalRequest)
	{
		var result=originalRequest.responseText;
		alert(result);
		location="awardAnswer.asp?rnd="+Math.random();
	}
}
-->
</script>
<%
Set Conn=nothing
Set User_Conn=nothing
Set Fs_User = Nothing
%>





