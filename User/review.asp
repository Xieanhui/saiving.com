<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
Dim DleIDStr
if Trim(Request.QueryString("Url"))<>"" then
	response.Write NoSqlHack(request.QueryString("Url"))&"-" &NoSqlHack(request.QueryString("type"))
	response.end
end if
if request("Action")="del" then
	if Request("id")="" then
		strShowErr = "<li>����Ĳ�����</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		DleIDStr = NoHtmlHackInput(NoSqlHack(Trim(Request("id"))))
		User_Conn.execute("Delete from FS_ME_Review where UserNumber = '" & Fs_User.UserNumber & "' And ReviewID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>ɾ���ɹ���</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Review.asp")
		Response.end
	end if
end if
if request("Action")="Lock" then
	if Request("id")="" then
		strShowErr = "<li>����Ĳ�����</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		DleIDStr = NoHtmlHackInput(NoSqlHack(Trim(Request("id"))))
		User_Conn.execute("Update FS_ME_Review set islock=1 where UserNumber = '" & Fs_User.UserNumber & "' And ReviewID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>�����ɹ���</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Review.asp")
		Response.end
	end if
end if
if request("Action")="UnLock" then
	if Request("id")="" then
		strShowErr = "<li>����Ĳ�����</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		DleIDStr = NoHtmlHackInput(NoSqlHack(Trim(Request("id"))))
		User_Conn.execute("Update FS_ME_Review set islock=0 where UserNumber = '" & Fs_User.UserNumber & "' And ReviewID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>�����ɹ���</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Review.asp")
		Response.end
	end if
end if
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,rs,i
int_RPP=20 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=5 '���ֵ�����ʾ��Ŀ
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
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" -->
    </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback">
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td  valign="top">���λ�ã�<a href="../">��վ��ҳ</a> &gt;&gt; <a href="main.asp">��Ա��ҳ</a> 
            &gt;&gt; <a href="Review.asp">���۹���</a> &gt;&gt; </td>
        </tr>
        <tr class="hback"> 
          <td  valign="top"><a href="Review.asp">ȫ������</a>��<a href="Review.asp?type=0">��������</a>��<%if IsExist_SubSys("DS") Then%><a href="Review.asp?type=1">��������</a>��<%End if%><%if IsExist_SubSys("MS") Then%><a href="Review.asp?type=2">��Ʒ����</a>��<%end if%><%if IsExist_SubSys("HS") Then%><a href="Review.asp?type=3">��������</a>��<%end if%><%if IsExist_SubSys("SD") Then%><a href="Review.asp?type=4">��������</a>��<%end if%><a href="Review.asp?type=5">�ռ�����</a>��<a href="Review.asp?type=6">�������</a></td>
        </tr>
      </table>
     
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="myForm" method="post" action="">
		  <tr> 
            <td width="6%" class="hback_1"><div align="center"> 
                <input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="���ѡ�����л��߳�������ѡ��">
              </div></td>
            <td width="27%" class="hback_1"><div align="left"><strong>����</strong></div></td>
            <td width="14%" class="hback_1"><div align="center"><strong>����</strong></div></td>
            <td width="10%" class="hback_1"><div align="center"><strong>��������Ϣ</strong></div></td>
            <td width="17%" class="hback_1"><div align="center"><strong>����</strong></div></td>
            <td width="6%" class="hback_1"><div align="center"><strong>״̬</strong></div></td>
            <td width="5%" class="hback_1"><strong>ͨ��</strong></td>
            <td width="15%" class="hback_1"><strong>����</strong></td>
          </tr>
          <%
		  	dim o_type
		  	select case NoSqlHack(Request.QueryString("type"))
				case "1"
					o_type = " and ReviewTypes=1"
				case "2"
					o_type = " and ReviewTypes=2"
				case "3"
					o_type = " and ReviewTypes=3"
				case "4"
					o_type = " and ReviewTypes=4"
				case "5"
					o_type = " and ReviewTypes=5"
				case "6"
					o_type = " and ReviewTypes=6"
				case "0"
					o_type = " and ReviewTypes=0"
				case else
					o_type = ""
			end select
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select * from FS_ME_Review where UserNumber='"& Fs_User.UserNumber&"' "& o_type &" order by Addtime desc,ReviewID desc",User_Conn,1,3
			if rs.eof then
			   rs.close
			   set rs=nothing
			   set conn=nothing
			   set fs_user=nothing
			   Response.Write"<tr  class=""hback""><td colspan=""8""  class=""hback"" height=""40"">û�м�¼��</td></tr>"
			else
				rs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("Page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo<=0 Then cPageNo=1
				If cPageNo>rs.PageCount Then cPageNo=rs.PageCount 
				rs.AbsolutePage=cPageNo
				for i=1 to rs.pagesize
					if rs.eof Then exit For 
		  %>
          <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="center"> 
                <input name="id" type="checkbox" id="id" value="<%=rs("ReviewID")%>">
              </div></td>
            <td class="hback" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(rid<%=rs("ReviewID")%>);" language=javascript><%=rs("title")%></td>
            <td class="hback"><div align="center">
			<%'0Ϊ�������ۣ�1Ϊ�������ۣ�2Ϊ��Ʒ��3Ϊ�������ۣ�4Ϊ�������ۣ�5�ռ�,6���
			select case rs("ReviewTypes")
				case 0
					response.Write"<a href=Review.asp?type=0>��������</a>"
				case 1
					response.Write"<a href=Review.asp?type=1>��������</a>"
				case 2
					response.Write"<a href=Review.asp?type=2>��Ʒ����</a>"
				case 3
					response.Write"<a href=Review.asp?type=3>��������</a>"
				case 4
					response.Write"<a href=Review.asp?type=4>��������</a>"
				case 5
					response.Write"<a href=Review.asp?type=5>��־����</a>"
				case 6
					response.Write"<a href=Review.asp?type=6>�������</a>"
				case else
					response.Write"<a href=Review.asp>-</a>"
			end select
			%></div></td>
            <td class="hback"><div align="center">
			<%'0Ϊ�������ۣ�1Ϊ�������ۣ�2Ϊ��Ʒ��3Ϊ�������ۣ�4Ϊ�������ۣ�5�ռ�,6���
			select case rs("ReviewTypes")
				case 0
					response.Write"<a href=Public_info.asp?type=NS&Url="&rs("InfoID")&" target=_blank>�鿴</a>"
				case 1
					response.Write"<a href=Public_info.asp?type=DS&Url="&rs("InfoID")&" target=_blank>�鿴</a>"
				case 2
					response.Write"<a href=Public_info.asp?type=MS&Url="&rs("InfoID")&" target=_blank>�鿴</a>"
				case 3
					response.Write"<a href=Public_info.asp?type=HS&Url="&rs("InfoID")&" target=_blank>�鿴</a>"
				case 4
					response.Write"<a href=Public_info.asp?type=SD&Url="&rs("InfoID")&" target=_blank>�鿴</a>"
				case 5
					response.Write"<a href=Public_info.asp?type=LS&Url="&rs("InfoID")&" target=_blank>�鿴</a>"
				case 6
					response.Write"<a href=Public_info.asp?type=PH&Url="&rs("InfoID")&" target=_blank>�鿴</a>"
				case else
					
			end select
			%>
		</div></td>
            <td class="hback"><div align="center"><%=rs("AddTime")%></div></td>
            <td class="hback"><div align="center"> 
                <%if rs("isLock")=1 then:response.Write"<span class=tx>����</span>":else:response.Write"����":end if%>
              </div></td>
            <td class="hback"> 
              <div align="center"><b><%if rs("AdminLock")=1 then:response.Write"<span class=tx>��</span>":else:response.Write"��":end if%></b></div>
            </td>
            <td class="hback"><div align="center"><a href="Review.asp?Action=del&id=<%=rs("ReviewID")%>" onClick="{if(confirm('ȷ��Ҫɾ����?')){return true;}return false;}">ɾ��</a>��<a href="Review.asp?Action=UnLock&id=<%=rs("ReviewID")%>">����</a>��<a href="Review.asp?Action=Lock&id=<%=rs("ReviewID")%>" onClick="{if(confirm('ȷ�����������𣿣�\n�����󽫲�����ʾ')){return true;}return false;}">����</a></div></td>
          </tr>
          <tr  class="hback" id="rid<%=rs("ReviewID")%>" style="display:none"> 
            <td height="40"><div align="center">����:</div></td>
            <td height="40" colspan="7"><%=rs("Content")%></td>
          </tr>
          <%
			  rs.movenext
		  next
		  %>
          <tr  class="hback"> 
            <td colspan="8"><div align="right"> 
                <%
			response.Write "<p>"&  fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
			%>
                <input name="Action" type="hidden" id="Action">
                <input type="button" name="Submit" value="ɾ��"  onClick="document.myForm.Action.value='del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.myForm.submit();return true;}return false;}">
                <input type="button" name="Submit2" value="��������"  onClick="document.myForm.Action.value='UnLock';{if(confirm('ȷ������������')){this.document.myForm.submit();return true;}return false;}">
                <input name="Submit3" type="button"  onClick="document.myForm.Action.value='Lock';{if(confirm('ȷ�����������𣿣�\n�����󽫲�����ʾ')){this.document.myForm.submit();return true;}return false;}" value="��������">
              </div></td>
          </tr>
          <% 
		  rs.close:set rs=nothing
		  end if
		  %>
        </form>
      </table>
    </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="Copyright.asp" -->
        </div></td>
    </tr>
</table>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = myForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = myForm.chkall.checked;  
    }  
	}
</script>

<%
Set Fs_User = Nothing
set user_conn=nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





