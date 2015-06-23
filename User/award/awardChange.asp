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
Dim ChangeRs,awardObj,currentDate,activeTF
activeTF=false
currentDate=DateValue(Now())
if G_IS_SQL_DB=0 then
	Set ChangeRs=User_Conn.execute("Select PrizeID from FS_ME_Prize where datediff('d',startDate,'"&NoSqlHack(currentDate)&"')>=0 And datediff('d',EndDate,'"&NoSqlHack(currentDate)&"')<0 And isChange=1")
ELse
	Set ChangeRs=User_Conn.execute("Select PrizeID from FS_ME_Prize where datediff(d,startDate,'"&NoSqlHack(currentDate)&"')>=0 And datediff(d,EndDate,'"&NoSqlHack(currentDate)&"')<0 And isChange=1")
End If
if not ChangeRs.eof then
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
		<td class="xingmu" height="20"><img src="../images/award.gif" alt="积分抽奖" border="0">积分兑换 </td>
	</tr>
	<%if not activeTF then%>
	<tr>
		<td class="hback"><img src="../images/alert.gif" />暂无兑换活动</td>
	</tr>
	<%Else%>
	<tr>
	  <td class="hback">
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
		<tr>
		<td>
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
				<%
					Dim UserRs,Integral'当前会员积分
					'获得当前会员积分--------------------------------
					Set UserRs=User_Conn.execute("Select Integral from FS_ME_Users where UserNumber='"&session("FS_UserNumber")&"'")
					if not UserRs.eof then
						Integral=UserRs("Integral")
					Else
						Integral=0
					End if
					'------------------------------------------------
					
					Dim PrizeArray,tr_count,i,joinNumber,action,provider
					tr_count=0
					while not ChangeRs.eof
						awardObj.getPrizeInfo(ChangeRs("prizeID"))
						joinNumber=User_Conn.execute("Select count(prizeID) from  FS_ME_User_Prize where prizeID="&ChangeRs("prizeID"))(0)
						if trim(awardObj.provider)<>"" then
							provider=awardObj.provider
						Else
							provider="本站"
						End if
						if  tr_count Mod 4=0 then
							Response.Write("<tr>"&vbcrlf)
						End if
						if Clng(Integral)<Clng(awardObj.prize_NeedPoint) then
							action="<img src=""../images/alert.gif""/><font color=""red"">积分不足</font>"
						Elseif Clng(joinNumber)>Clng(awardObj.storage) or Clng(joinNumber)=Clng(awardObj.storage) then
							action="<img src=""../images/alert.gif""/><font color=""red"">兑换完毕</font>"
						Else
							action="<a href=""#"" onClick=""changeAward("&ChangeRs("prizeID")&","&awardObj.prize_NeedPoint&")""><img src=""../images/bottomduihuan.gif"" border=""0"" alt=""兑换奖品""/></a>"
						End if
						Response.Write("<td class=""hback"">"&vbcrlf)
						Response.Write("<table bodor=""0""  cellpadding=""1"" cellspacing=""1"" >"&vbcrlf)
						Response.Write("<tr>"&vbcrlf)
						Response.Write("<td><div><img src="""&awardObj.PrizePic&""" width=""80"" height=""80"" border=""0"" alt=""该奖品由"&provider&"提供""></div></td>")
						Response.Write("<td><div><img src=""../images/award.gif""/><font color=""red"">"&awardObj.PrizeName&"</font><img src=""../images/award.gif""/></div><div><img src=""../images/moneyOrPoint.gif"" alt=""抽取该奖品所需要的积分""/>&nbsp;"&awardObj.prize_NeedPoint&" 积分</div><div>参加人数："&joinNumber&"</div><div>"&action&"</div></td>")
						Response.Write("</tr>"&vbcrlf)
						Response.Write("</table>"&vbcrlf)
						Response.Write("</td>")
						tr_count=tr_count+1
						if  tr_count Mod 4=0 then
							Response.Write("</tr>"&vbcrlf)
						End if
						ChangeRs.movenext
					Wend
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
function changeAward(id,Integral)
{
	var url="awardAction.asp"
	var pars="action=change&Integral="+Integral+"&prizeID="+id+"&rnd="+Math.random();
	if(confirm("确定要进行该操作\n该次操作将消费积分："+Integral))
	{
		 var myAjax = new Ajax.Request(url,{method: 'get', parameters: pars, onComplete: showResponse});

	}
	function showResponse(originalRequest)
	{
		var result=originalRequest.responseText;
		alert(result);
		location="awardChange.asp?rnd="+Math.random();
	}
}
-->
</script>
<%
Set Conn=nothing
Set User_Conn=nothing
Set Fs_User = Nothing
%>





