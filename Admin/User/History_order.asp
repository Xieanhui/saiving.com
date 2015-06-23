<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
	Dim Conn,User_Conn,strShowErr
	MF_Default_Conn
	MF_User_Conn
	'session判断
	MF_Session_TF 
	
	if not MF_Check_Pop_TF("ME_Horder") then Err_Show 
	if not MF_Check_Pop_TF("ME033") then Err_Show 

	Function GetFriendName(f_strNumber)
		Dim RsGetFriendName
		Set RsGetFriendName = User_Conn.Execute("Select UserName From FS_ME_Users Where UserNumber = '"& f_strNumber &"'")
		If  Not RsGetFriendName.eof  Then 
			GetFriendName = RsGetFriendName("UserName")
		Else
			GetFriendName = 0
		End If 
		set RsGetFriendName = nothing
	End Function 
	if Request.Form("Action")="Del" then
		if trim(Request.Form("Id"))="" then
			strShowErr = "<li>请选择至少一项</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Delete From FS_ME_Log where LogId in ("&FormatIntArr(Request.Form("Id"))&")")
			Call MF_Insert_oper_Log("删除交易明晰","ID("& NoSqlHack(Replace(Request.Form("Id")," ",""))&")",now,session("admin_name"),"ME")
			strShowErr = "<li>删除成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/History_Order.asp")
			Response.end
	end if
	if Request.Form("Action")="Delall" then
		User_Conn.execute("Delete From FS_ME_Log ")
		Call MF_Insert_oper_Log("删除交易明晰","删除所有用户的交易明晰",now,session("admin_name"),"ME")
	end if
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i
	int_RPP=30 '设置每页显示数目
	int_showNumberLink_=8 '数字导航显示数目
	showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
	str_nonLinkColor_="#999999" '非热链接颜色
	toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
	toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
	toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
	toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
	toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
	toL_="<font face=webdings title=""最后一页"">:</font>"
	strpage=request("page")
	if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>日志-网站内容管理系统</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="hback"><strong>交易明晰</strong> | <strong><a href="../News/Constr_Manage.asp">稿件管理</a></strong></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td colspan="8" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="44%"> <strong> 
            <%
				  dim strTmp,strLogType,strTmp1
				  strLogType = NoSqlHack(Trim(Request.QueryString("LogTye")))
			     if Request.QueryString("LogTye")<>"" then
			  		strTmp =  " and LogType='"& strLogType &"'"
			     Else
			  		strTmp =  " "
			    End if
				if Request("date1") <>"" and  Request("date2")<>"" then
					if isdate(Request("date1"))=false or isdate(Request("date2"))=false then
						strShowErr = "<li>您输入的日期格式不正确</li>"
						Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
						Response.end
					else
						if G_IS_SQL_User_DB =0 then
							strTmp1 = " and datevalue(Logtime)>=#"&datevalue(Request("date1"))&"#  and datevalue(Logtime)<=#"&datevalue(Request("date2"))&"#"
						Else
							strTmp1 = " and convert(varchar(10),Logtime,120)>='"&datevalue(Request("date1"))&"'  and convert(varchar(10),Logtime,120)<='"&datevalue(Request("date2"))&"'"
						End if
					End if
				End if
				Dim RsUserListObj,RsUserSQL
				Dim strSQLs,StrOrders
				strpage=request("page")
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsUserListObj = Server.CreateObject(G_FS_RS)
				RsUserSQL = "Select LogId,LogType,UserNumber,points,moneys,LogTime,LogContent,Logstyle From Fs_ME_Log  where LogID>0 "& strTmp & strTmp1 &" order by  LogID desc"
				RsUserListObj.Open RsUserSQL,User_Conn,1,3
				Response.Write "<Font color=red>" & RsUserListObj.RecordCount&"</font>"
				%>
            </strong> 个日志　类型：<a href="History_Order.asp">所有</a>｜<a href="history_order.asp?LogTye=%D7%A2%B2%E1">注册</a>｜<a href="history_order.asp?LogTye=%B5%C7%C2%BD">登陆</a>｜<a href="history_order.asp?LogTye=购买">购买</a>｜<a href="history_order.asp?LogTye=%D4%DA%CF%DF%D6%A7%B8%B6">冲值</a>｜<a href="history_order.asp?LogTye=兑换">兑换</a>｜<a href="history_order.asp?LogTye=其他">其他</a></td>
          <form action="history_order.asp"  method="post" name="myform" id="myform">
            <td width="56%"><div align="left"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="63%" valign="top">从 
                      <input name="date1" type="text" id="date1" value="<%=datevalue(date())-1%>" size="10">
                      到 
                      <input name="date2" type="text" id="date2" value="<%=datevalue(date())%>" size="10">
                      的记录 
                      <input type="submit" name="Submit" value="搜索">
                      日期格式请用1977-6-7格式</td>
                  </tr>
                </table>
              </div></td>
          </form>
        </tr>
      </table></td>
  </tr class="hback">
  <form action="history_order.asp"  method="post" name="form1" id="form1">
    <tr class="hback"> 
      <td width="16%" class="xingmu"><div align="left"><strong> 类型</strong></div></td>
      <td width="7%" class="xingmu"><div align="center"><strong>点数</strong></div></td>
      <td width="16%" class="xingmu"><div align="center"><strong>用户</strong></div></td>
      <td width="9%" class="xingmu"><div align="center"><strong>金币</strong></div></td>
      <td width="20%" class="xingmu"><div align="center"><strong>日期</strong></div></td>
      <td width="20%" class="xingmu"><div align="center"><strong>说明</strong></div></td>
      <td width="9%" class="xingmu"><div align="center"><strong>增加/减少</strong></div></td>
      <td width="3%" class="xingmu">&nbsp;</td>
    </tr>
    <%
		if RsUserListObj.eof then
		   RsUserListObj.close
		   set RsUserListObj=nothing
		   Response.Write"<tr  class=""hback""><td colspan=""8""  class=""hback"" height=""40"">没有管理员。</td></tr>"
		else
			RsUserListObj.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("Page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then cPageNo=1
			If cPageNo>RsUserListObj.PageCount Then cPageNo=RsUserListObj.PageCount 
			RsUserListObj.AbsolutePage=cPageNo
			for i=1 to RsUserListObj.pagesize
				if RsUserListObj.eof Then exit For 
		%>
    <tr class="hback"> 
      <td class="hback"><div align="left"><a href=history_Order.asp?LogTye=<% = RsUserListObj("LogType")%>> 
          <% = RsUserListObj("LogType")%>
          </a></div></td>
      <td class="hback"> 
        <div align="center"><% = RsUserListObj("points")%>
        </div></td>
      <td class="hback"><div align="center"><a href="../../<%=G_USER_DIR%>/ShowUser.asp?UserNumber=<% = RsUserListObj("UserNumber")%>" target="_blank"> 
          <% = GetFriendName(RsUserListObj("UserNumber"))%>
          </a></div></td>
      <td class="hback"> <div align="center">
          <% = FormatNumber(RsUserListObj("moneys"),2,-1)%>
        </div></td>
      <td class="hback"><div align="center"> 
          <% = RsUserListObj("LogTime")%>
        </div></td>
      <td class="hback"><div align="center"> 
          <% = RsUserListObj("LogContent")%>
        </div></td>
      <td class="hback"><div align="center"> 
          <%
			  if RsUserListObj("Logstyle") = 0 then
				  Response.Write("<font color=red>增加</font>")
			  Else
				  Response.Write("减少")
			  End if
			  %>
        </div></td>
      <td class="hback"><input name="ID" type="checkbox" id="ID" value="<% = RsUserListObj("LogID")%>"></td>
    </tr>
    <%
			  RsUserListObj.MoveNext
		  Next
		  %>
    <tr class="hback"> 
      <td colspan="8" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="80%"> <span class="top_navi"> 
              <%
			response.Write "<p>"&  fPageCount(RsUserListObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
			%>
              <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
              选中所有短信 
              <input name="Action" type="hidden" id="Action" value="">
              <input type="button" name="Submit2" value="删除选中的明晰"  onClick="document.form1.Action.value='Del';{if(confirm('确定清除您所选择的记录吗？')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit22" value="删除全部明晰"  onClick="document.form1.Action.value='Delall';{if(confirm('确定清除您所选择的记录吗？')){this.document.form1.submit();return true;}return false;}">
              </SPAN></td>
          </tr>
          <%end if%>
        </table></td>
    </tr>
  </FORM>
</table>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form1.elements.length;i++)  
    {  
    var e = form1.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form1.chkall.checked;  
    }  
  }
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





