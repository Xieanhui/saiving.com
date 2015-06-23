<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
session("audit")=NoSqlHack(request.QueryString("audit"))
session("classid")=CintStr(request.QueryString("classid"))
session("myclassid")=CintStr(request.QueryString("myclassid"))
Dim contrRs,classid,sql_info_cmd,AuditTF
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
if session("classid")<>"0" and session("classid")<>"" then'如果classid不为根目录
	if  session("audit")="1" then '如果已经审核
		if session("myclassid")<>"" then
			sql_info_cmd="Select ContID,ContTitle,addtime,classid,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,IsLock,isTF,MainID,untread,AuditTF,type from FS_ME_InfoContribution where classid="&session("myclassid")&" and  AuditTF=1 and usernumber='"&session("FS_UserNumber")&"'"
			if session("classid")<>"" then
				sql_info_cmd=sql_info_cmd&"and MainID="&session("classid")
			end if
		else
			sql_info_cmd="Select ContID,ContTitle,addtime,classid,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,IsLock,isTF,MainID,untread,AuditTF,type from FS_ME_InfoContribution where  AuditTF=1 and usernumber='"&session("FS_UserNumber")&"'"
			if session("classid")<>"" then
				sql_info_cmd=sql_info_cmd&"and MainID="&session("classid")
			end if
		End if
	else'如果未审核
		if session("myclassid")<>"" then
			sql_info_cmd="Select ContID,ContTitle,addtime,classid,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,IsLock,isTF,MainID,untread,AuditTF,type from FS_ME_InfoContribution where classid="&session("myclassid")&"  and usernumber='"&session("FS_UserNumber")&"'"
			if session("classid")<>"" then
				sql_info_cmd=sql_info_cmd&"and MainID="&session("classid")
			end if
		else
			sql_info_cmd="Select ContID,ContTitle,addtime,classid,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,IsLock,isTF,MainID,untread,AuditTF,type from FS_ME_InfoContribution where usernumber='"&session("FS_UserNumber")&"'"
			if session("classid")<>"" then
				sql_info_cmd=sql_info_cmd&"and MainID="&session("classid")
			end if
		End if
	End if
else'如果classid为空
	'增加session("myclassid")为0的判断,09-07-08 by SamJun
	if session("audit")="1" then'如果已经审核  
		if session("myclassid")<>"" and session("myclassid")<>"0" then
			sql_info_cmd="Select ContID,ContTitle,addtime,classid,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,IsLock,isTF,MainID,untread,AuditTF,type from FS_ME_InfoContribution where classid="&session("myclassid")&" and  AuditTF=1  and usernumber='"&session("FS_UserNumber")&"' order by addtime DESC,ContID DESC"
		else
			sql_info_cmd="Select ContID,ContTitle,addtime,classid,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,IsLock,isTF,MainID,untread,AuditTF,type from FS_ME_InfoContribution where  AuditTF=1 and usernumber='"&session("FS_UserNumber")&"' order by addtime DESC,ContID DESC"
		End if
	else'如果未审核 
		if session("myclassid")<>"" and session("myclassid")<>"0" then
			sql_info_cmd="Select ContID,ContTitle,addtime,classid,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,IsLock,isTF,MainID,untread,AuditTF,type from FS_ME_InfoContribution where classid="&session("myclassid")&"  and usernumber='"&session("FS_UserNumber")&"' order by addtime DESC,ContID DESC"
		else
			sql_info_cmd="Select ContID,ContTitle,addtime,classid,InfoType,ContSytle,UserNumber,IsPublic,AdminLock,IsLock,isTF,MainID,untread,AuditTF,type from FS_ME_InfoContribution where usernumber='"&session("FS_UserNumber")&"' order by addtime DESC,ContID DESC"
		End if
	End if
End IF
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>欢迎用户<%=Fs_User.UserName%>来到<%=GetUserSystemTitle%></title>
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../../FS_Inc/PublicJS.js"></script>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
</head>

<body style="margin-left: 0px;margin-top: 0px;margin-right: 0px;margin-bottom: 0px;">
<table width="100%" height="25" border="0" align="center" cellpadding="1" cellspacing="1" class="table" id="listContainer">
<tr>
<td class="xingmu" align="center" colspan="6">新 闻</td>
</tr>
<tr>
<td class="hback" align="center">标题</td>
<td class="hback" align="center">分类</td>
<td class="hback" align="center">添加时间</td>
<td class="hback" align="center">状态</td>
<td class="hback" align="center">操作</td>
<td class="hback" align="center"><input type="checkbox" name="contrList" value="" onclick="selectAll(document.all('contrList'))"/></td>
</tr>
<%	
	Dim contrType,ContSytle,InfoType,classRs,className,AuditStatus,constrlock
	Set contrRs=Server.CreateObject(G_FS_Rs)
	contrRs.open sql_info_cmd,User_Conn,1,3
	If Not contrRs.eof then
'分页使用-----------------------------------
		contrRs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<=0 Then cPageNo=1
		If cPageNo>contrRs.PageCount Then cPageNo=contrRs.PageCount 
		contrRs.AbsolutePage=cPageNo
	End if
	for i=0 to int_RPP
		if contrRs.eof then exit for
		select case contrRs("type")
			case "0" contrType="[新闻]"
			case "1" contrType="[下载]"
			case "2" contrType="[商品]"
			case else contrType="[新闻]"
		end select
		select case contrRs("ContSytle")
			case "0" ContSytle="[原创]"
			case "1" ContSytle="[转载]"
			case "3" ContSytle="[代理]"
			case else ContSytle="[原创]"
		End select
		select case contrRs("InfoType")
			case "0" InfoType="[普通]"
			case "1" InfoType="<font color='#A4D234'>[优先]</font>"
			case "2" InfoType="<font color='red'>[加急]</font>"
			case else InfoType="[普通]"
		End select
		'获得分类名------------------------------------------------------------*/
		Set classRs=Conn.execute("select ClassID,ClassName from FS_NS_NewsClass where isConstr=1 and id="&contrRs("MainID"))
		if not classRs.eof then
			className=classRs("ClassName")
		Else
			ClassName="无"
		End if
		classRs.close
		Set classRs=nothing
		if contrRs("AuditTF")="1" then
			AuditStatus="<font color='#A4D234'>已审核</font>"
		Else
			AuditStatus="未审核"
		End if
		if contrRs("IsLock")="1" then
			constrlock="<a href='解锁' onClick=""contrAction('unlock','"&contrRs("ContID")&"','span_lock_"&contrRs("ContID")&"');return false;""><font color='red'>解锁</font></a>"
		Else
			constrlock="<a href='锁定' onClick=""contrAction('lock','"&contrRs("ContID")&"','span_lock_"&contrRs("ContID")&"');return false;"">锁定</a>"
		End if
		Dim delStr,modStr,lockStr
		delStr="<a href='删除' onClick=""if(confirm('确定要删除？'))contrAction('delete','"&contrRs("ContID")&"','span_lock_"&contrRs("ContID")&"');return false;"">删除</a>"
		modStr="<a href='contrEdit.asp?action=edit&id="&contrRs("ContID")&"' target='_parent'>修改</a>"
		lockStr="<span id='span_lock_"&contrRs("ContID")&"'>"&constrlock&"</span>"
		'End获得分类名------------------------------------------------------------*/
		Response.Write("<tr onMouseOver=overColor(this) onMouseOut=outColor(this) height='20'>"&vbcrlf)
		Response.Write("<td class='hback'><a href='contrEdit.asp?action=edit&id="&contrRs("ContID")&"' target='_parent'>{"&InfoType&contrType&ContSytle&"} "&contrRs("ContTitle")&"</a></td>"&vbcrlf)
		Response.Write("<td class='hback' align='center'>"&className&"</td>"&vbcrlf)
		Response.Write("<td class='hback' align='center'>"&contrRs("addtime")&"</td>"&vbcrlf)
		Response.Write("<td class='hback' align='center'>"&AuditStatus&"</td>"&vbcrlf)
		Response.Write("<td class='hback' align='center'>"&modStr&" |"&delStr&"| "&lockStr&"</td>"&vbcrlf)
		Response.Write("<td class='hback' align='center'><input type='checkbox' name='contrList' value='"&contrRs("ContID")&"'/></td>"&vbcrlf)
		contrRs.movenext
		Response.Write("</tr>"&vbcrlf)
	next
%>
<tr>
<td class="hback" align="right" colspan="6">
<button name="bnt_lock" onclick="parent.location='contrEdit.asp?action=add'">投 稿</button>&nbsp;
<button name="bnt_lock" onclick="contrBatAction('lock')">批量锁定</button>&nbsp;
<button name="bnt_unlock" onclick="contrBatAction('unlock')" >批量解锁</button>&nbsp;
<button name="bnt_delete" onclick="contrBatAction('delete')">批量删除</button>&nbsp;
</td>
</tr>
<%
Response.Write("<tr>"&vbcrlf)
Response.Write("<td align='right' colspan='10'  class=""hback"">"&fPageCount(contrRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td>"&vbcrlf)
Response.Write("</tr>"&vbcrlf)
%>
</table>
</body>
</html>
<%
	Conn.close
	User_Conn.close
	Set Conn=nothing
	Set contrRs=nothing
%>
<script language="javascript">
<!--
parent.hd_audit.value="<%=session("audit")%>";
parent.hd_classid.value="<%=session("classid")%>";
parent.hd_myclassid.value="<%=session("myclassid")%>";
function contrAction(act,id,container)
{
	$(container).innerHTML="<img src='../../sys_images/small_loading.gif'/>";
	if(act=="lock")
	{	
		var ajax1=new Ajax.Updater(container,"contrAction.asp",{method:'get',parameters:"action=lock&id="+id+"&and"+Math.random()});
	}else if(act=="unlock")
	{
		var ajax2=new Ajax.Updater(container,"contrAction.asp",{method:'get',parameters:"action=unlock&id="+id+"&and"+Math.random()});
	}else if(act=="delete")
	{
			var ajax3=new Ajax.Request("contrAction.asp",{method:'get',parameters:"action=delete&id="+id+"&and"+Math.random(),onComplete:showResponse});
	}
	function showResponse(originalRequest)
	{
		var result=originalRequest.responseText;
		if(result=="ok")
		{
			$('listContainer').firstChild.removeChild($(container).parentNode.parentNode)
			$('recordcount').innerHTML=parseInt($('recordcount').innerText)-1;
		}else
		{
			alert("发生异常，请联系技术人员！");
		}
	}
}
function contrBatAction(act)
{
	var elements=document.all("contrList");
	if(act=="lock")
	{
		for(var i=1;i<elements.length;i++)
		{
			if(elements[i].checked)
				contrAction('lock',elements[i].value,"span_lock_"+elements[i].value);
		}
	}else if(act=="unlock")
	{
		for(var i=1;i<elements.length;i++)
		{
			if(elements[i].checked)
				contrAction('unlock',elements[i].value,"span_lock_"+elements[i].value);
		}
	}else if(act=="delete")
	{
		if(confirm("确定要删除？"))
		{
			for(var i=1;i<elements.length;i++)
			{
				if(elements[i].checked)
					contrAction('delete',elements[i].value,"span_lock_"+elements[i].value);
			}
		}
	}
}
-->
</script>