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
			strShowErr = "<li>��ѡ������һ��</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Delete From FS_ME_Infoilog where iLogID in ("&FormatIntArr(Request("Id"))&")")
			strShowErr = "<li>ɾ����־�ɹ�</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog.asp")
			Response.end
	end if
	if Request("Action")="DelAll" then
		User_Conn.execute("Delete From FS_ME_Infoilog")
			Call MF_Insert_oper_Log("ɾ����־","ɾ����������־",now,session("admin_name"),"ME")
			strShowErr = "<li>ɾ��������־�ɹ�</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog.asp")
			Response.end
	end if
	if Request("Action")="UnLock" then
		if trim(Request("Id"))="" then
			strShowErr = "<li>��ѡ������һ��</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		User_Conn.execute("Update FS_ME_Infoilog set AdminLock=0 where IsLock=0 and iLogID in ("&FormatIntArr(Request("Id"))&")")
			strShowErr = "<li>�����ɹ�</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog.asp")
			Response.end
	end if
	if Request("Action")="Lock" then
		if trim(Request("Id"))="" then
			strShowErr = "<li>��ѡ������һ��</li>"
			Response.Redirect("../Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
			User_Conn.execute("Update FS_ME_Infoilog set AdminLock=1 where iLogID in ("&FormatIntArr(Request("Id"))&")")
			strShowErr = "<li>�����ɹ�</li>"
			Response.Redirect("../Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=User/iLog.asp")
			Response.end
	end if
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i
	int_RPP=25 '����ÿҳ��ʾ��Ŀ
	int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
	showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
	str_nonLinkColor_="#999999" '����������ɫ
	toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
	toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
	toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
	toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
	toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
	toL_="<font face=webdings title=""���һҳ"">:</font>"
	strpage=CintStr(request("page"))
	if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
%>
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr> 
      <td width="100%" class="xingmu">��־��ժ����</td>
    </tr>
    <tr> 
      
    <td class="hback"><a href="iLog.asp">��־����</a>��<a href="iLog_Templet.asp">ģ������</a>��<a href="iLog_Class.asp">ϵͳ��Ŀ</a>��<a href="iLog_SetParam.asp">��������</a></td>
    </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form action="iLog.asp"  method="post" name="form1" id="form1">
    <tr class="hback"> 
      <td width="31%" class="xingmu"><div align="left"><strong> ����</strong></div></td>
      <td width="13%" class="xingmu"><div align="center"><strong>����</strong></div></td>
      <td width="12%" class="xingmu"><div align="center"><strong>������</strong></div></td>
      <td width="15%" class="xingmu"><div align="center"><strong>����</strong></div></td>
      <td width="13%" class="xingmu"><div align="center"><strong>״̬(�û�)</strong></div></td>
      <td width="13%" class="xingmu"><div align="center"><strong>״̬(����Ա)</strong></div></td>
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
		   Response.Write"<tr  class=""hback""><td colspan=""7""  class=""hback"" height=""40"">û�м�¼��</td></tr>"
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
		  if rs_ilog("iLogStyle")=0 then:response.Write"�ռ�":else:response.Write"��ժ":end if
		  %></a>
        </div></td>
      <td class="hback"><div align="center"><a href="../../<%=G_USER_DIR%>/ShowUser.asp?UserNumber=<% = rs_ilog("UserNumber")%>" target="_blank"> 
          <% = GetFriendName(rs_ilog("UserNumber"))%>
          </a> </div></td>
      <td class="hback"><div align="center"> 
          <% = rs_ilog("addtime")%>
        </div></td>
      <td class="hback"><div align="center"> <%if rs_ilog("isLock")=0 then:response.Write"����":else:response.Write"����":end if%>
        </div></td>
      <td class="hback"><div align="center"> 
          <%
		  if rs_ilog("adminLock")=0 then
			  response.Write"<a href=iLog.asp?id="&rs_iLog("iLogId")&"&Action=Lock>����</a>"
		  elseif rs_ilog("adminLock")=1 then
			  response.Write"<a href=iLog.asp?id="&rs_iLog("iLogId")&"&Action=UnLock><span class=""tx"">����</span></a>"
		  end if
		  %>
        </div></td>
      <td class="hback"><div align="center"> 
          <input name="ID" type="checkbox" id="ID" value="<% = rs_ilog("iLogID")%>">
        </div></td>
    </tr>
    <tr class="hback" id="rid<%=rs_ilog("iLogID")%>" style="display:none"> 
      <td height="31" colspan="7" class="hback"> <strong>��־����:</strong> 
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
              ȫѡ 
              <input name="Action" type="hidden" id="Action" value="">
              <input type="button" name="Submit2" value="ɾ��"  onClick="document.form1.Action.value='Del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit22" value="��������"  onClick="document.form1.Action.value='UnLock';{if(confirm('��ȷ��Ҫ����������')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit23" value="��������"  onClick="document.form1.Action.value='Lock';{if(confirm('��ȷ��Ҫ������')){this.document.form1.submit();return true;}return false;}">
              <input type="button" name="Submit232" value="ɾ������"  onClick="document.form1.Action.value='DelAll';{if(confirm('��ȷ��Ҫɾ��������־��\n   ɾ���󽫲��ָܻ�!!!')){this.document.form1.submit();return true;}return false;}">
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





