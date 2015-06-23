<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim Conn
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo

int_RPP=20'设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("MF_Log") then Err_Show
if not MF_Check_Pop_TF("MF018") then Err_Show
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<%
Dim obj_Log_Rs,strpage,select_count,select_pagecount,i,Tmp_adminname,Tmp_super,Tmp_Lock,tmp_my,SQL,tmp_admin_name,tmp_SubType
strpage=NoSqlHack(request("page"))
'if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
Set  obj_Log_Rs = server.CreateObject(G_FS_RS)
if tRIM(Request("Admin_Name"))<>"" then:tmp_admin_name=" and Admin_Name='"& NoSqlHack(Trim(Request("Admin_Name"))) &"'":else:tmp_admin_name="":end if
if tRIM(Request("SubType"))<>"" then:tmp_SubType=" and Logtype='"& NoSqlHack(tRIM(Request("SubType"))) &"'":else:tmp_SubType="":end if
SQL = "Select ID,LogTitle,LogContent,LogTime,Admin_Name,Logtype  from FS_MF_Oper_Log where ID>0  "& tmp_SubType & tmp_admin_name &" Order by id desc"
obj_Log_Rs.Open SQL,Conn,1,1
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr class="xingmu">
    <td class="xingmu"><a href="#" class="sd"><strong>管理员操作日志</strong></a>
      </td>
  </tr>
  <tr class="hback">
    <td><a href="Sys_Oper_Log.asp">首页</a>｜<a href="Sys_Oper_Log.asp?SubType=MF">主系统</a>｜<a href="Sys_Oper_Log.asp?SubType=NS">新闻</a>｜<%if Request.Cookies("FoosunSUBCookie")("FoosunSUBMS")=1 then%><a href="Sys_Oper_Log.asp?SubType=MS">商城</a>｜<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBDS")=1 then%><a href="Sys_Oper_Log.asp?SubType=DS">下载</a>｜<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBCS")=1 then%><a href="Sys_Oper_Log.asp?SubType=CS">采集</a>｜<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBAP")=1 then%><a href="Sys_Oper_Log.asp?SubType=AP">人才</a>｜<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBHS")=1 then%><a href="Sys_Oper_Log.asp?SubType=HS">房产</a>｜<%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBSS")=1 then%><a href="Sys_Oper_Log.asp?SubType=SS">统计</a>｜<%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBSD")=1 then%><a href="Sys_Oper_Log.asp?SubType=SD">供求</a>｜<%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBME")=1 then%><a href="Sys_Oper_Log.asp?SubType=ME">会员</a>｜<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBVS")=1 then%><a href="Sys_Oper_Log.asp?SubType=VS">投票</a>｜<%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBAS")=1 then%><a href="Sys_Oper_Log.asp?SubType=AS">广告</a>｜<%End if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBWS")=1 then%><a href="Sys_Oper_Log.asp?SubType=WS">留言本</a>｜<%end if%><%if Request.Cookies("FoosunSUBCookie")("FoosunSUBFL")=1 then%><a href="Sys_Oper_Log.asp?SubType=FL">友情连接</a><%End if%>　<a href="Sys_Login_Log.asp"><strong>安全日志</strong></a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="LogForm" method="post" action="">
    <tr class="hback"> 
      <td width="17%" height="25" class="xingmu"> <div align="left">标题(点标题查看操作描述)</div></td>
      <td width="13%" height="25" class="xingmu"> <div align="center">日期</div></td>
      <td width="18%" height="25" class="xingmu"> <div align="center">管理员</div></td>
      <td width="11%" height="25" class="xingmu"> <div align="center">子系统</div></td>
    </tr>
    <%
if obj_Log_Rs.eof then
   obj_Log_Rs.close
   set obj_Log_Rs=nothing
   Response.Write"<TR  class=""hback""><TD colspan=""6""  class=""hback"" height=""40"">没有操作。</TD></TR>"
else
	obj_Log_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo>obj_Log_Rs.PageCount Then cPageNo=obj_Log_Rs.PageCount 
	If cPageNo<=0 Then cPageNo=1
	obj_Log_Rs.AbsolutePage=cPageNo
	for i=1 to obj_Log_Rs.pagesize
		if obj_Log_Rs.eof Then exit For 
%>
    <tr class="hback"> 
      <td height="25"> <div  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(Logid<% = obj_Log_Rs("ID")%>);"  language=javascript>·<a href="#"><% = obj_Log_Rs("LogTitle")%></a></div></td>
      <td height="25"><div align="center"> 
          <% = obj_Log_Rs("LogTime")%>
        </div></td>
      <td height="25"><div align="center"> <a href="Sys_Oper_Log.asp?Admin_Name=<% = obj_Log_Rs("Admin_Name")%>"> 
          <% = obj_Log_Rs("Admin_Name")%>
          </a> </div></td>
      <td height="25"><div align="center"> 
          <% = obj_Log_Rs("Logtype")%>
        </div></td>
    </tr>
    <tr valign="top" class="hback" id="Logid<% = obj_Log_Rs("ID")%>" style="display:none"> 
      <td height="32" colspan="4"> 　 
        <% = obj_Log_Rs("LogContent")%>
      </td>
    </tr>
    <%
		obj_Log_Rs.movenext
	Next
	%>
    <tr class="hback"> 
      <td height="25" colspan="4"><div align="right"> 
          <input name="Action" type="hidden" id="Action">
          <input type="button" name="Submit222" value="删除所有日志，只能删除最近<% = G_HOLD_LOG_DAY_NUM %>天以前的日志"   onClick="document.LogForm.Action.value='Del';{if(confirm('确定清除日志吗？')){this.document.LogForm.submit();return true;}return false;}">
        </div></td>
    </tr>
    <tr class="hback">
      <td height="25" colspan="4"><span class="tx">子系统类型说明</span><br>
        MF-主系统 NS-新闻系统 MS-商城 DS-下载 CS-采集系统 AP-人才系统 HS-房产系统 <br>
        SS-统计系统 ME-会员系统 VS-投票系统 AS-广告系统 WS-留言系统 FL-友情联接</td>
    </tr>
  </form>
</table>

<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr height="1" class="hback"> 
    <td height="25"> 
      <%
		response.Write "<p>"&  fPageCount(obj_Log_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
    End if%>
    </td>
  </tr>
</table>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
</html>
<%
set obj_Log_Rs = nothing
if Request.Form("Action")="Del" then
	if not MF_Check_Pop_TF("MF020") then Err_Show
	Dim obj_DLog_rs,strShowErr,tmp_i,DatAllowDate,f_Sql
	Set  obj_DLog_rs = server.CreateObject(G_FS_RS)
	DatAllowDate=dateadd("d",-G_HOLD_LOG_DAY_NUM,date())
	If G_IS_SQL_DB=1 then
		f_Sql = "Select ID from FS_MF_Oper_Log Where logtime<='" & DatAllowDate & " 23:59:59' Order by  id desc"
	Else
		f_Sql = "Select ID from FS_MF_Oper_Log Where  logtime<=#" & DatAllowDate & " 23:59:59# Order by  id desc"
	End If
	obj_DLog_rs.Open f_Sql,Conn,1,3
	tmp_i = 0 
	Do while not obj_DLog_rs.eof 
		Conn.execute("Delete From FS_MF_Oper_Log where ID="&CintStr(obj_DLog_rs("ID")))
		obj_DLog_rs.movenext
		tmp_i = tmp_i +1 
	Loop
	if tmp_i>0 then
		Call MF_Insert_oper_Log("删除操作日志","删除所有操作日志,删除共"& tmp_i &"个日志",now,session("admin_name"),"MF")
	End if
	strShowErr = "<li>日志删除成功,共删除"& tmp_i &"个日志</li>"
	Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
	obj_DLog_rs.close:set obj_DLog_rs = nothing
End if
Set Conn = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





