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
			strShowErr = "<li>��ѡ������һ��</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Delete From FS_ME_Report where ReportID in ("&FormatIntArr(Request.Form("Id"))&")")
			strShowErr = "<li>ɾ���ɹ�</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../UserRePort.asp")
			Response.end
	end if
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i
	
	int_RPP=15 '����ÿҳ��ʾ��Ŀ
	int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
	showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
	str_nonLinkColor_="#999999" '����������ɫ
	toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
	toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
	toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
	toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
	toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
	toL_="<font face=webdings title=""���һҳ"">:</font>"
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
    <td width="100%" class="xingmu">�ٱ�����</td>
  </tr class="hback">
  <tr class="hback">
    <td class="hback"><a href="UserReport.asp">����</a></td>
  </tr class="hback">
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form action="UserReport.asp"  method="post" name="form1" id="form1">
    <tr class="hback"> 
      <td width="17%" class="xingmu"><div align="left"><strong> ����</strong></div></td>
      <td width="15%" class="xingmu"><div align="left"><strong>�ٱ���</strong></div></td>
      <td width="11%" class="xingmu"><div align="left"><strong>���ٱ���</strong></div></td>
      <td width="20%" class="xingmu"><div align="center"><strong>����</strong></div></td>
      <td width="10%" class="xingmu"><div align="center"><strong>����</strong></div></td>
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
		   Response.Write"<tr  class=""hback""><td colspan=""6""  class=""hback"" height=""40"">û�м�¼��</td></tr>"
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
					response.Write "<a href=UserReport.asp?type=0>ƭ��</a>"
				case 1
					response.Write "<a href=UserReport.asp?type=1>���</a>"
				case 2
					response.Write "<a href=UserReport.asp?type=2>��������</a>"
				case 3
					response.Write "<a href=UserReport.asp?type=3>�Ƿ�����</a>"
				case else
					response.Write "<a href=UserReport.asp?type=4>����</a>"
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
              ѡ�����ж��� 
              <input name="Action" type="hidden" id="Action" value="">
              <input type="button" name="Submit2" value="ɾ��ѡ����Ϣ"  onClick="document.form1.Action.value='Del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.form1.submit();return true;}return false;}">
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
				response.Write("�Ҳ�����¼")
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






