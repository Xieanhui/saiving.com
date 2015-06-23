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
Dim awardRs,awardObj,currentDate,currentAwardID
currentAwardID=0
currentDate=DateValue(Now())
Set awardRs=User_Conn.execute("Select awardID from FS_ME_award where opened=0")
if not awardRs.eof then
	currentAwardID=awardRs("awardID")
	Set awardObj=New cls_award
	awardObj.getAwardInfo(currentAwardID)
End if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=GetUserSystemTitle%></title>
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../../FS_Inc/ProtoType.js"></script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<body class="hback">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
	<tr>
		<td class="xingmu" height="20"><img src="../images/award.gif" alt="积分抽奖" border="0">积分抽奖</td>
	</tr>
	<%if currentAwardID=0 then%>
	<tr>
		<td class="hback"><img src="../images/alert.gif" />暂无抽奖活动</td>
	</tr>
	<%Else%>
	<tr>
	  <td class="hback">
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
	  	<tr>
			<td class="hback" width="400" height="120"><div id="awardPic" align="center"><img src="<%=awardObj.AwardPic%>" width="400" height="120" alt="第 <%=awardObj.awardid%> 期积分抽奖主题图片" /></div></td>
			<td class="hback" valign="top">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr>
                <td class="hback" align="center"><strong>第 <%=awardObj.awardid%> 期</strong></td>
              </tr>
              <tr>
                <td class="hback" align="center"><strong>主题：<%=awardObj.awardName%></strong></td>
              </tr>
              <tr>
                <td class="hback" align="center">活动时间从<%=awardObj.award_startDate%>到<%=awardObj.award_EndDate%></td>
              </tr>
			<tr>
                <td class="hback" align="center">离结束时间还有<font color="#FF0000"><%=(DateValue(awardObj.award_EndDate)-DateValue(Now()))%></font>天</td>
              </tr>
            </table></td>
		</tr>
		<tr>
		<td colspan="2">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
				<%
					Dim UserRs,Integral'当前会员积分
					Dim onceMoreTF
					onceMoreTF=false
					'获得当前会员积分--------------------------------
					Set UserRs=User_Conn.execute("Select Integral from FS_ME_Users where UserNumber='"&session("FS_UserNumber")&"'")
					if not UserRs.eof then
						Integral=UserRs("Integral")
					Else
						Integral=0
					End if
					'------------------------------------------------
					
					Dim PrizeArray,tr_count,i,joinNumber,Rs
					tr_count=0
					PrizeArray=split(awardObj.PrizeIDS,",")
					if isArray(PrizeArray) then
						for i=0 to Ubound(PrizeArray)
							if not isNumeric(PrizeArray(i)) then exit for
							awardObj.getPrizeInfo(PrizeArray(i))
							joinNumber=User_Conn.execute("Select count(prizeID) from  FS_ME_User_Prize where prizeID="&PrizeArray(i))(0)
							if  tr_count Mod 4=0 then
								Response.Write("<tr>"&vbcrlf)
							End if
							Set Rs=User_Conn.execute("Select ID From FS_ME_User_Prize where prizeid="&PrizeArray(i)&" And usernumber='"&session("FS_UserNumber")&"'")
							onceMoreTF=false
							if not Rs.eof then
								onceMoreTF=true
							End if
							Rs.close
							Set Rs=nothing
							Response.Write("<td class=""hback"">"&vbcrlf)
							Response.Write("<table bodor=""0""  cellpadding=""1"" cellspacing=""1"" >"&vbcrlf)
							Response.Write("<tr>"&vbcrlf)
							Response.Write("<td><div><img src="""&awardObj.PrizePic&""" width=""80"" height=""80"" border=""0""></div></td>")
							Response.Write("<td><div><img src=""../images/award.gif""/><font color=""red"">"&awardObj.PrizeName&"</font><img src=""../images/award.gif""/></div><div><font color=""red"">"&awardObj.PrizeGrade&"</font> 等奖</div><div><img src=""../images/moneyOrPoint.gif"" alt=""抽取该奖品所需要的积分""/> "&awardObj.prize_NeedPoint&" 积分</div><div>参加人数："&joinNumber&"</div></td>")
							Response.Write("</tr>"&vbcrlf)
							Response.Write("<tr>"&vbcrlf)
							if Clng(Integral)<Clng(awardObj.prize_NeedPoint) then
								Response.Write("<td class=""hback"" clospan=""2""><img src=""../images/alert.gif""/><font color=""red"">积分不足</font></td>"&vbcrlf)
							Elseif onceMoreTF Then
								Response.Write("<td colspan=""2"" class=""hback"">&nbsp;&nbsp;<img src=""../images/alert.gif""/><font color=""red"">已参与过</font></td>")
							Else
								Response.Write("<td class=""hback"" clospan=""2""><a href=""#"" onclick=""joinAward("&PrizeArray(i)&","&awardObj.prize_NeedPoint&","&awardRs("awardID")&")""><img src=""../images/joinaward.bmp"" border=""0""/></a></td>"&vbcrlf)
							End if
							Response.write("</tr>"&vbcrlf)
							Response.Write("</table>"&vbcrlf)
							Response.Write("</td>")
							tr_count=tr_count+1
							if  tr_count Mod 4=0 then
								Response.Write("</tr>"&vbcrlf)
							End if
						next
						if tr_count Mod 4<>0 then
							for i=(tr_count Mod 4)+1 to 4
								Response.Write("<td class=""hback"">"&vbcrlf)
								Response.Write("<table bodor=""0""  cellpadding=""1"" cellspacing=""1"">"&vbcrlf)
								Response.Write("<tr>"&vbcrlf)
								Response.Write("<td></td>")
								Response.Write("</tr>"&vbcrlf)
								Response.Write("</table>"&vbcrlf)
								Response.Write("</td>")
							next
							Response.Write("</tr>"&vbcrlf)
						End if
					End if
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
function checkIT(Obj)
{
	Obj.firstChildNode.checked=true;
}
function joinAward(id,Integral,awardID)
{
	var url="awardAction.asp"
	var pars="action=join&awardID="+awardID+"&Integral="+Integral+"&prizeID="+id+"&rnd="+Math.random();
	if(confirm("确定要进行该操作\n该次抽奖将消费积分 :"+Integral))
	{
		 var myAjax = new Ajax.Request(url,{method: 'get', parameters: pars, onComplete: showResponse});

	}
	function showResponse(originalRequest)
	{
		var result=originalRequest.responseText;
		var joinNumber;
		alert(result);
		location="awardPane.asp?rnd="+Math.random();
	}
}
-->
</script>
<%
Set Conn=nothing
Set User_Conn=nothing
Set Fs_User = Nothing
%>





