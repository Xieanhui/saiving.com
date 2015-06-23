<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
	dim Conn,User_Conn,rs,str_c_isp,str_c_user,str_c_pass,str_c_url,str_domain,rs_param,str_c_gurl,strShowErr
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	if not MF_Check_Pop_TF("ME_Log") then Err_Show 
	if not MF_Check_Pop_TF("ME039") then Err_Show 

	Function GetFriendName(f_strNumber)
		Dim RsGetFriendName
		Set RsGetFriendName = User_Conn.Execute("Select UserName From FS_ME_Users Where UserNumber = '"& NoSqlHack(f_strNumber) &"'")
		If  Not RsGetFriendName.eof  Then 
			GetFriendName = RsGetFriendName("UserName")
		Else
			GetFriendName = 0
		End If 
		set RsGetFriendName = nothing
	End Function 
	if Request("Action")="Del" then
		if trim(Request("Id"))="" then
			strShowErr = "<li>请选择至少一项</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Delete From FS_ME_Infoilog where iLogID in ("&FormatIntArr(Request("Id"))&")")
			strShowErr = "<li>删除日志成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog.asp")
			Response.end
	end if
	if Request("Action")="DelAll" then
		User_Conn.execute("Delete From FS_ME_Infoilog")
			Call MF_Insert_oper_Log("删除日志","删除了所有日志",now,session("admin_name"),"ME")
			strShowErr = "<li>删除所有日志成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog.asp")
			Response.end
	end if
	if Request("Action")="UnLock" then
		if trim(Request("Id"))="" then
			strShowErr = "<li>请选择至少一项</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Update FS_ME_Infoilog set AdminLock=0 where IsLock=0 and iLogID in ("&FormatIntArr(Request("Id"))&")")
			strShowErr = "<li>解锁成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog.asp")
			Response.end
	end if
	if Request("Action")="Lock" then
		if trim(Request("Id"))="" then
			strShowErr = "<li>请选择至少一项</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
			User_Conn.execute("Update FS_ME_Infoilog set AdminLock=1 where iLogID in ("&FormatIntArr(Request("Id"))&")")
			strShowErr = "<li>锁定成功</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog.asp")
			Response.end
	end if
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i
	int_RPP=25 '设置每页显示数目
	int_showNumberLink_=8 '数字导航显示数目
	showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
	str_nonLinkColor_="#999999" '非热链接颜色
	toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
	toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
	toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
	toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
	toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
	toL_="<font face=webdings title=""最后一页"">:</font>"
	strpage=CintStr(request("page"))
	if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
%>
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="100%" class="xingmu">日志网摘管理</td>
    </tr>
    <tr> 
      
    <td class="hback"><a href="iLog.asp">日志管理</a>┆<a href="iLog_Templet.asp">模板设置</a>┆<a href="iLog_Class.asp">系统栏目</a>┆<a href="iLog_SetParam.asp">参数设置</a></td>
    </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form action="iLog.asp"  method="post" name="form1" id="form1">
    <tr class="hback"> 
      <td width="31%" class="xingmu"><div align="left"><strong> 标题</strong></div></td>
      <td width="13%" class="xingmu"><div align="center"><strong>类型</strong></div></td>
      <td width="12%" class="xingmu"><div align="center"><strong>发表人</strong></div></td>
      <td width="15%" class="xingmu"><div align="center"><strong>日期</strong></div></td>
      <td width="13%" class="xingmu"><div align="center"><strong>状态(用户)</strong></div></td>
      <td width="13%" class="xingmu"><div align="center"><strong>状态(管理员)</strong></div></td>
      <td width="3%" class="xingmu">&nbsp;</td>
    </tr>
    <%
		dim rs_ilogsql,rs_ilog,str_type,str_isLock,iLogStyle,AdminLock
		strpage=CintStr(request("page"))
		if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
		if trim(Request.QueryString("iLogStyle"))<>"" then:iLogStyle=" and iLogStyle="&clng(Request.QueryString("iLogStyle"))&"":else:iLogStyle="":end if
		if trim(Request.QueryString("AdminLock"))<>"" then:AdminLock=" and AdminLock="&clng(Request.QueryString("AdminLock"))&"":else:AdminLock="":end if
		Set rs_ilog = Server.CreateObject(G_FS_RS)
		rs_ilogsql = "Select * From FS_ME_Infoilog  where 1=1 "& iLogStyle & AdminLock &" order by  isTop desc, Addtime desc, iLogID desc"
		rs_ilog.Open rs_ilogsql,User_Conn,1,1
		if rs_ilog.eof then
		   rs_ilog.close
		   set rs_ilog=nothing
		   Response.Write"<tr  class=""hback""><td colspan=""7""  class=""hback"" height=""40"">没有记录。</td></tr>"
		else
			rs_ilog.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("Page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo>rs_ilog.PageCount Then cPageNo=rs_ilog.PageCount 
			If cPageNo<=0 Then cPageNo=1
			rs_ilog.AbsolutePage=cPageNo
			for i=1 to int_RPP
				if rs_ilog.eof Then exit For 
		%>
    <tr class="hback"> 
      <td class="hback"><div align="left" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(rid<%=rs_ilog("iLogID")%>);" language=javascript> 
          <a href="#"> 
          <% = rs_ilog("Title")%>
          </a></div></td>
      <td class="hback"><div align="center">
         <a href="iLog.asp?iLogStyle=<%=rs_ilog("iLogStyle")%>"><%
		  if rs_ilog("iLogStyle")=0 then:response.Write"日记":else:response.Write"网摘":end if
		  %></a>
        </div></td>
      <td class="hback"><div align="center"><a href="../../<%=G_USER_DIR%>/ShowUser.asp?UserNumber=<% = rs_ilog("UserNumber")%>" target="_blank"> 
          <% = GetFriendName(rs_ilog("UserNumber"))%>
          </a> </div></td>
      <td class="hback"><div align="center"> 
          <% = rs_ilog("addtime")%>
        </div></td>
      <td class="hback"><div align="center"> <%if rs_ilog("isLock")=0 then:response.Write"开放":else:response.Write"锁定":end if%>
        </div></td>
      <td class="hback"><div align="center"> 
          <%
		  if rs_ilog("adminLock")=0 then
			  response.Write"<a href=iLog.asp?id="&rs_iLog("iLogId")&"&Action=Lock>开放</a>"
		  elseif rs_ilog("adminLock")=1 then
			  response.Write"<a href=iLog.asp?id="&rs_iLog("iLogId")&"&Action=UnLock><span class=""tx"">锁定</span></a>"
		  end if
		  %>
        </div></td>
      <td class="hback"><div align="center"> 
          <input name="ID" type="checkbox" id="ID" value="<% = rs_ilog("iLogID")%>">
        </div></td>
    </tr>
    <tr class="hback" id="rid<%=rs_ilog("iLogID")%>" style="display:none"> 
      <td height="31" colspan="7" class="hback"> <strong>日志内容:</strong> 
        <% = rs_ilog("Content")%>
      </td>
    </tr>
    <%
		  rs_ilog.MoveNext
	  Next
	  %>
    <tr class="hback"> 
      <td colspan="7" class="hback"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="80%"> <span class="top_navi"> 
              <%
			response.Write "<p>"&  fPageCount(rs_ilog,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
			%>
              <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
              全选 
              <input name="Action" type="hidden" id="Action" value="">
              <input type="button" name="Submit2" value="删除"  onClick="document.form1.Action.value='Del';{if(confirm('确定清除您所选择的记录吗？')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit22" value="批量解锁"  onClick="document.form1.Action.value='UnLock';{if(confirm('您确定要批量解锁吗？')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit23" value="批量锁定"  onClick="document.form1.Action.value='Lock';{if(confirm('您确定要锁定？')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit232" value="删除所有"  onClick="document.form1.Action.value='DelAll';{if(confirm('您确定要删除所有日志吗？\n   删除后将不能恢复!!!')){this.document.form1.submit();return true;}return false;}">
              </SPAN></td>
          </tr>
          <%end if%>
        </table></td>
    </tr>
  </FORM>
</table>
</body>
</html>
<%
Conn.close:set conn=nothing
User_Conn.close:set User_Conn=nothing
%><script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form1.elements.length;i++)  
    {  
    var e = form1.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form1.chkall.checked;  
    }  
  }
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
</script> 





