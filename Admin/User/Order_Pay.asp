<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<%
Dim Conn,User_Conn,strShowErr
MF_Default_Conn
MF_User_Conn
'session�ж�
MF_Session_TF 

if not MF_Check_Pop_TF("ME_Order") then Err_Show 

Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
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
if Request.Form("Action") = "Del" then
	if not MF_Check_Pop_TF("ME031") then Err_Show
	'Response.Write("Delete From FS_ME_Order  where OrderId in ("& Request.Form("Id")&")")
	'Response.End() 
	if Request.Form("Id")="" then
	strShowErr = "<li>���������ѡ��һ�����ɾ��</li>"
	Response.Redirect("../error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Order_Pay.asp")
	Response.end
	Else
	User_Conn.execute("Delete From FS_ME_Order  where OrderId in ("& FormatIntArr(Request.Form("Id"))&")")
	Call MF_Insert_oper_Log("ɾ������֧������","����ID��("& Request.Form("Id")&")",now,session("admin_name"),"ME")
	set User_Conn=nothing
	set conn=nothing
	strShowErr = "<li>���������ɹ�!</li>"
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Order_Pay.asp")
	Response.end
End if
 end if
if Request.QueryString("type_1")="success" then
	User_Conn.execute("Update FS_ME_Order set isSuccess=1,M_state=1,isLock=0 where OrderId ="& Request.QueryString("Id"))
	Call MF_Insert_oper_Log("�������߶���Ϊ[��֧��]","�����ţ�("& Request.QueryString("OrderNumber")&")",now,session("admin_name"),"ME")
	strShowErr = "<li>���³ɹ�!</li>"
	set User_Conn=nothing
	set conn=nothing
	Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/Order_Pay.asp")
	Response.end
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>����-��վ���ݹ���ϵͳ</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head><body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback">
    <td class="hback"><a href="Order_Pay.asp">����֧������</a></td>
  </tr>
  <tr class="hback">
    <td class="hback"><a href="Order_Pay.asp">ȫ��</a>��<a href="Order_Pay.asp?success=1">�ɹ�����</a>��<a href="Order_Pay.asp?success=0">ʧ�ܶ���</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="form1" method="post" action="Order_Pay.asp">
    <tr class="hback">
      <td colspan="7" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="44%"><strong>
              <%
				  dim strTmp,strLogType,strTmp1
				  strLogType = NoSqlHack(Trim(Request.QueryString("LogTye")))
			     if Request.QueryString("LogTye")<>"" then
			  		strTmp =  " and LogType='"& strLogType &"'"
			     Else
			  		strTmp =  " "
			    End if
				Dim RsOrderObj,RsOrderSQL
				Dim strpage,strSQLs,StrOrders,strmp_or,IsSuccess
				strpage=request("page")
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsOrderObj = Server.CreateObject(G_FS_RS)
				if trim(Request("keyword"))<>"" then
					if request("type")="0" then
						strmp_or = " and OrderNumber Like '%"&trim(Request("keyword"))&"%'"
					else
						strmp_or = " and UserNumber Like '%"&trim(Request("keyword"))&"%'"
					end if
				else
					strmp_or=""
				end if
				if Request("success")="1" then
					IsSuccess = " and IsSuccess=1"
				elseif Request("success")="0" then
					IsSuccess = " and IsSuccess=0"
				else
					IsSuccess = ""
				end if
				RsOrderSQL = "Select * From FS_ME_Order  where  OrderType=3 "& strmp_or & IsSuccess &" order by  OrderID desc"
				RsOrderObj.Open RsOrderSQL,User_Conn,1,1
				Response.Write RsOrderObj.recordcount
				%>
              </strong> ������</td>
            <td width="56%"><div align="left">
                <input name="keyword" type="text" id="keyword">
                <select name="type" id="type">
                  <option value="0" selected>������</option>
                  <option value="1">�û����</option>
                </select>
                <input type="submit" name="Submit2" value="����">
              </div></td>
          </tr>
        </table></td>
    </tr class="hback">
    <tr class="hback">
      <td width="20%" class="xingmu"><div align="left"><strong> ������(�㶨���鿴����)</strong></div></td>
      <td width="9%" class="xingmu"><div align="center"><strong>���״̬</strong></div></td>
      <td width="16%" class="xingmu"><div align="center"><strong>�û���</strong></div></td>
      <td width="18%" class="xingmu"><div align="center"><strong>֧���ɹ�����</strong></div></td>
      <td width="19%" class="xingmu"><div align="center"><strong>����</strong></div></td>
      <td width="8%" class="xingmu"><div align="center"><strong>����</strong></div></td>
      <td width="10%" class="xingmu"><div align="center"><strong>֧��</strong></div></td>
    </tr>
    <%
		Dim select_count,select_pagecount,i
		if RsOrderObj.eof then
			   RsOrderObj.close
			   set RsOrderObj=nothing
			   set conn=nothing
			   set user_conn=nothing
			   Response.Write"<TR><TD colspan=""7""  class=""hback"">û�м�¼��</TD></TR>"
		else
				if Request("CountPage")="" or len(Request("CountPage"))<1 then
					RsOrderObj.pagesize = 20
				Else
					RsOrderObj.pagesize = Request("CountPage")
				End if
				RsOrderObj.absolutepage=cint(strpage)
				select_count=RsOrderObj.recordcount
				select_pagecount=RsOrderObj.pagecount
				for i=1 to RsOrderObj.pagesize
					if RsOrderObj.eof Then exit For 
		 %>
    <tr class="hback">
      <td class="hback"><div align="left">
          <% = RsOrderObj("OrderNumber")%>
        </div></td>
      <td class="hback"><div align="center">
          <%
					if RsOrderObj("isLock")=1 then
						Response.Write"<span class=tx>�����...<span>"
					Else
						Response.Write"�����..."
					End if
					%>
        </div></td>
      <td class="hback"><div align="center"> <a href="../../<%=G_USER_DIR%>/ShowUser.asp?UserNumber=<% = RsOrderObj("UserNumber")%>" target="_blank">
          <% = GetFriendName(RsOrderObj("UserNumber"))%>
          </a> </div></td>
      <td class="hback"><div align="center">
          <% = RsOrderObj("M_PayDate")%>
        </div></td>
      <td class="hback"><div align="center">
          <% = RsOrderObj("AddTime")%>
        </div></td>
      <td class="hback"><div align="center">
          <%
			if RsOrderObj("OrderType")=0 then
				Response.Write("��Ա��")
			Elseif RsOrderObj("OrderType")=1 then
				Response.Write("��Ʒ")
			Elseif RsOrderObj("OrderType")=2 then
				Response.Write("�㿨")
			Elseif RsOrderObj("OrderType")=3 then
				Response.Write("����֧��")
			Else
				Response.Write("����")
			End if
			%>
        </div></td>
      <td class="hback"><div align="center">
          <%
		  if RsOrderObj("IsSuccess")=0 then
		  	if RsOrderObj("M_state")=0 then
		  %>
          <a href="Order_Pay.asp?type_1=success&id=<%=RsOrderObj("OrderID")%>&OrderNumber=<% = RsOrderObj("OrderNumber")%>" title="����Ϊ�ɹ�֧������" onClick="{if(confirm('�Ƿ����Ϊ[[֧���ɹ�]]��')){return true;}return false;}">ʧ��</a>
          <%
		  	else
				response.Write "�Ѵ���" 
			end if
		  Else%>
          <span class="tx">�ɹ�</span>
          <%End if%>
          <input name="id" type="checkbox" id="id" value="<%=RsOrderObj("OrderID")%>">
        </div></td>
    </tr>
    <%
			  RsOrderObj.MoveNext
		  Next
		  %>
    <tr class="hback">
      <td colspan="7" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="hback">
            <td width="40%"><% 	Response.Write("ÿҳ:"& RsOrderObj.pagesize &"��,")
							Response.write"&nbsp;��<b>"& select_pagecount &"</b>ҳ<b>&nbsp;" & select_count &"</b>����¼����ҳ�ǵ�<b>"& strpage &"</b>ҳ��"
							if int(strpage)>1 then
								Response.Write"&nbsp;<a href=?page=1&keyword="&Request("keyword")&"&type="&request("type")&">��һҳ</a>&nbsp;&nbsp;"
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&keyword="&Request("keyword")&"&type="&request("type")&">��һҳ</a>&nbsp;&nbsp;"
							End if
							If int(strpage)<select_pagecount then
								Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&keyword="&Request("keyword")&"&type="&request("type")&">��һҳ</a>&nbsp;"
								Response.Write"&nbsp;<a href=?page="& select_pagecount &"&keyword="&Request("keyword")&"&type="&request("type")&">���һҳ</a>&nbsp;&nbsp;"
							End if
								Response.Write"<br>"
								RsOrderObj.close
								Set RsOrderObj=nothing
				%>
            </td>
            <td width="40%" class="hback"><input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
              ѡ�����ж���
              <input name="Action" type="hidden" id="Action" value="">
              <input type="button" name="Submit" value="ɾ��ѡ�еĶ���"  onClick="document.form1.Action.value='Del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.form1.submit();return true;}return false;}">
            </td>
          </tr>
        </table></td>
    </tr>
    <%End if%>
  </form>
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->






