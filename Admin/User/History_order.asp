<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
	Dim Conn,User_Conn,strShowErr
	MF_Default_Conn
	MF_User_Conn
	'session�ж�
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
			strShowErr = "<li>��ѡ������һ��</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Delete From FS_ME_Log where LogId in ("&FormatIntArr(Request.Form("Id"))&")")
			Call MF_Insert_oper_Log("ɾ����������","ID("& NoSqlHack(Replace(Request.Form("Id")," ",""))&")",now,session("admin_name"),"ME")
			strShowErr = "<li>ɾ���ɹ�</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/History_Order.asp")
			Response.end
	end if
	if Request.Form("Action")="Delall" then
		User_Conn.execute("Delete From FS_ME_Log ")
		Call MF_Insert_oper_Log("ɾ����������","ɾ�������û��Ľ�������",now,session("admin_name"),"ME")
	end if
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i
	int_RPP=30 '����ÿҳ��ʾ��Ŀ
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
<title>��־-��վ���ݹ���ϵͳ</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="hback"><strong>��������</strong> | <strong><a href="../News/Constr_Manage.asp">�������</a></strong></td>
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
						strShowErr = "<li>����������ڸ�ʽ����ȷ</li>"
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
            </strong> ����־�����ͣ�<a href="History_Order.asp">����</a>��<a href="history_order.asp?LogTye=%D7%A2%B2%E1">ע��</a>��<a href="history_order.asp?LogTye=%B5%C7%C2%BD">��½</a>��<a href="history_order.asp?LogTye=����">����</a>��<a href="history_order.asp?LogTye=%D4%DA%CF%DF%D6%A7%B8%B6">��ֵ</a>��<a href="history_order.asp?LogTye=�һ�">�һ�</a>��<a href="history_order.asp?LogTye=����">����</a></td>
          <form action="history_order.asp"  method="post" name="myform" id="myform">
            <td width="56%"><div align="left"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="63%" valign="top">�� 
                      <input name="date1" type="text" id="date1" value="<%=datevalue(date())-1%>" size="10">
                      �� 
                      <input name="date2" type="text" id="date2" value="<%=datevalue(date())%>" size="10">
                      �ļ�¼ 
                      <input type="submit" name="Submit" value="����">
                      ���ڸ�ʽ����1977-6-7��ʽ</td>
                  </tr>
                </table>
              </div></td>
          </form>
        </tr>
      </table></td>
  </tr class="hback">
  <form action="history_order.asp"  method="post" name="form1" id="form1">
    <tr class="hback"> 
      <td width="16%" class="xingmu"><div align="left"><strong> ����</strong></div></td>
      <td width="7%" class="xingmu"><div align="center"><strong>����</strong></div></td>
      <td width="16%" class="xingmu"><div align="center"><strong>�û�</strong></div></td>
      <td width="9%" class="xingmu"><div align="center"><strong>���</strong></div></td>
      <td width="20%" class="xingmu"><div align="center"><strong>����</strong></div></td>
      <td width="20%" class="xingmu"><div align="center"><strong>˵��</strong></div></td>
      <td width="9%" class="xingmu"><div align="center"><strong>����/����</strong></div></td>
      <td width="3%" class="xingmu">&nbsp;</td>
    </tr>
    <%
		if RsUserListObj.eof then
		   RsUserListObj.close
		   set RsUserListObj=nothing
		   Response.Write"<tr  class=""hback""><td colspan=""8""  class=""hback"" height=""40"">û�й���Ա��</td></tr>"
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
				  Response.Write("<font color=red>����</font>")
			  Else
				  Response.Write("����")
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
              ѡ�����ж��� 
              <input name="Action" type="hidden" id="Action" value="">
              <input type="button" name="Submit2" value="ɾ��ѡ�е�����"  onClick="document.form1.Action.value='Del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit22" value="ɾ��ȫ������"  onClick="document.form1.Action.value='Delall';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.form1.submit();return true;}return false;}">
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





