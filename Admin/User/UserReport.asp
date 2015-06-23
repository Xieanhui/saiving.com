<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
	dim Conn,User_Conn,strShowErr
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	if not MF_Check_Pop_TF("ME_Jubao") then Err_Show 
	if not MF_Check_Pop_TF("ME037") then Err_Show 

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
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Delete From FS_ME_Report where ReportID in ("&FormatIntArr(Request.Form("Id"))&")")
			strShowErr = "<li>删除成功</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../UserRePort.asp")
			Response.end
	end if
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i
	
	int_RPP=15 '设置每页显示数目
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
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td width="100%" class="xingmu">举报管理</td>
  </tr class="hback">
  <tr class="hback">
    <td class="hback"><a href="UserReport.asp">返回</a></td>
  </tr class="hback">
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form action="UserReport.asp"  method="post" name="form1" id="form1">
    <tr class="hback"> 
      <td width="17%" class="xingmu"><div align="left"><strong> 标题</strong></div></td>
      <td width="15%" class="xingmu"><div align="left"><strong>举报人</strong></div></td>
      <td width="11%" class="xingmu"><div align="left"><strong>被举报人</strong></div></td>
      <td width="20%" class="xingmu"><div align="center"><strong>日期</strong></div></td>
      <td width="10%" class="xingmu"><div align="center"><strong>类型</strong></div></td>
      <td width="2%" class="xingmu">&nbsp;</td>
    </tr>
    <%
		dim rs_reportsql,rs_report
		strpage=request("page")
		if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
		Set rs_report = Server.CreateObject(G_FS_RS)
		if Request.QueryString("type")<>"" then
			rs_reportsql = "Select * From FS_ME_Report  where ReportType="&cint(Request.QueryString("type"))&" order by  ReportID desc"
		else
			rs_reportsql = "Select * From FS_ME_Report   order by  ReportID desc"
		end if
		rs_report.Open rs_reportsql,User_Conn,1,1
		if rs_report.eof then
		   rs_report.close
		   set rs_report=nothing
		   Response.Write"<tr  class=""hback""><td colspan=""6""  class=""hback"" height=""40"">没有记录。</td></tr>"
		else
			rs_report.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("Page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo>rs_report.PageCount Then cPageNo=rs_report.PageCount 
			If cPageNo<=0 Then cPageNo=1
			rs_report.AbsolutePage=cPageNo
			for i=1 to int_RPP
				if rs_report.eof Then exit For 
					dim s_b,s_b_
					if rs_report("isRead")=0 then
						s_b="<b>"
						s_b_="</b>"
					end if
		%>
    <tr class="hback"> 
      <td class="hback"><div align="left"> <a href="UserReport.asp?id=<% = rs_report("ReportID")%>&Read=1"> 
          <%=s_b%><% = left(rs_report("Content"),20)&"..."%><%=s_b_%>
          </a></div></td>
      <td class="hback"><div align="left"> <%=s_b%><a href="../../<%=G_USER_DIR%>/ShowUser.asp?UserNumber=<% = rs_report("UserNumber")%>" target="_blank">
          <% = GetFriendName(rs_report("UserNumber"))%><%=s_b_%>
          </a> </div></td>
      <td class="hback"><div align="left"> <%=s_b%><a href="../../<%=G_USER_DIR%>/ShowUser.asp?UserNumber=<% = rs_report("f_UserNumber")%>" target="_blank">
          <% = GetFriendName(rs_report("f_UserNumber"))%><%=s_b_%>
          </a> </div></td>
      <td class="hback"><div align="center"> 
          <%=s_b%><% = rs_report("addtime")%><%=s_b_%>
        </div></td>
      <td class="hback"><div align="center"> 
          <%=s_b%><%
			select case rs_report("ReportType")
				case 0
					response.Write "<a href=UserReport.asp?type=0>骗子</a>"
				case 1
					response.Write "<a href=UserReport.asp?type=1>广告</a>"
				case 2
					response.Write "<a href=UserReport.asp?type=2>攻击别人</a>"
				case 3
					response.Write "<a href=UserReport.asp?type=3>非法言论</a>"
				case else
					response.Write "<a href=UserReport.asp?type=4>其他</a>"
			end select
		%><%=s_b_%>
        </div></td>
      <td class="hback"><input name="ID" type="checkbox" id="ID" value="<% = rs_report("ReportID")%>"></td>
    </tr>
    <%
			  rs_report.MoveNext
		  Next
		  %>
    <tr class="hback"> 
      <td colspan="6" class="hback"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="80%"> <span class="top_navi"> 
              <%
			response.Write "<p>"&  fPageCount(rs_report,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
			%>
              <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
              选中所有短信 
              <input name="Action" type="hidden" id="Action" value="">
              <input type="button" name="Submit2" value="删除选中信息"  onClick="document.form1.Action.value='Del';{if(confirm('确定清除您所选择的记录吗？')){this.document.form1.submit();return true;}return false;}">
              </SPAN></td>
          </tr>
          <%end if%>
        </table>
</td>
    </tr>
  </FORM>
</table>
		  <%if Request.QueryString("Read")="1" then%>
		<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr>
            
    <td height="55" class="hback"> 
      <%
			dim rs
			set rs= Server.CreateObject(G_FS_RS)
			rs.open "select Content,isRead From FS_ME_Report  where ReportID="&NoSqlHack(Request.QueryString("Id")),User_Conn,1,3
			if rs.eof then
				response.Write("找不到记录")
			else
				rs("isRead")=1
				rs.update
				response.Write rs("Content")
			end if
			rs.close:set rs=nothing
			%>
    </td>
          </tr>
        </table>
		<%end if%>
		</body>
</html>
<%
Conn.close:set conn=nothing
User_Conn.close:set User_Conn=nothing
%>
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






