<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
dim rs_mysys
If CheckBlogOpen=False Then 
	Response.write("<script language=""javascript"">alert('��־������ͣʹ��,����Ҫʹ������ϵ����Ա.');history.back();</script>")
	Response.End()
End If
set rs_mysys = User_Conn.execute("select id From FS_ME_InfoiLogParam where UserNumber='"& Fs_User.UserNumber&"'")
if rs_mysys.eof then
	Response.write("<br>Ҫ������־���뿪ͨ������־�ռ�,5���ת��...")
	Response.Write("<meta http-equiv=""refresh"" content=""5;url=PublicParam.asp"">")
	response.end
end if
if request("Action")="del" then
	if Request("id")="" then
		strShowErr = "<li>����Ĳ�����</li>"
		Response.Redirect("../lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Delete from FS_ME_Infoilog where iLogID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>ɾ���ɹ���</li>"
		Response.Redirect("../lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../i_Blog/index.asp")
		Response.end
	end if
end if
if request("Action")="Lock" then
	if Request("id")="" then
		strShowErr = "<li>����Ĳ�����</li>"
		Response.Redirect("../lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Update FS_ME_Infoilog set islock=1 where iLogID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>�����ɹ���</li>"
		Response.Redirect("../lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../i_Blog/index.asp")
		Response.end
	end if
end if
if request("Action")="UnLock" then
	if Request("id")="" then
		strShowErr = "<li>����Ĳ�����</li>"
		Response.Redirect("../lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Update FS_ME_Infoilog set islock=0 where iLogID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>�����ɹ���</li>"
		Response.Redirect("../lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../i_Blog/index.asp")
		Response.end
	end if
end if

Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,rs,i
int_RPP=20 '����ÿҳ��ʾ��Ŀ
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
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head></head>
<body>

<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" -->
    </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback">
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td  valign="top">���λ�ã�<a href="../../">��վ��ҳ</a> &gt;&gt; <a href="../main.asp">��Ա��ҳ</a> 
            &gt;&gt; <a href="index.asp">��־����</a> &gt;&gt;��־����</td>
        </tr>
        <tr class="hback"> 
          <td  valign="top"><a href="index.asp">��־��ҳ</a>��<a href="PublicLog.asp">������־</a>��<a href="index.asp?type=box">�ݸ���</a>��<a href="../PhotoManage.asp">������</a>��<a href="PublicParam.asp">��������</a>��<a href="../Review.asp">���۹���</a></td>
        </tr>
      </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="myform" method="post" action="">
          <tr> 
            <td width="4%" class="xingmu"><div align="center"> 
                <input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="���ѡ�����л��߳�������ѡ��">
              </div></td>
            <td width="31%" class="xingmu"><div align="center">����</div></td>
            <td width="15%" class="xingmu"><div align="center">����</div></td>
            <td width="17%" class="xingmu"><div align="center">����</div></td>
            <td width="14%" class="xingmu"><div align="center">״̬</div></td>
            <td width="19%" class="xingmu"><div align="center">����</div></td>
          </tr>
          <%
		  	dim o_class,o_draff
		  	if request.QueryString("classid")<>"" then
				o_class= " and ClassId="&CintStr(request.QueryString("classid"))&""
			else
				o_class= ""
			end if
		  	if request.QueryString("type")="box" then
				o_draff= " and isDraft=1"
			else
				o_draff= ""
			end if
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select * from FS_ME_Infoilog where UserNumber='"& Fs_User.UserNumber&"' "&o_draff&o_class&" order by isTop desc,AddTime desc,iLogID desc",User_Conn,1,3
			if rs.eof then
			   rs.close
			   set rs=nothing
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
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="center"> 
                <input name="id" type="checkbox" id="id" value="<%=rs("iLogID")%>">
              </div></td>
            <td class="hback"><a href="PublicLogEdit.asp?id=<%=rs("iLogID")%>"><%=rs("Title")%></a></td>
            <td class="hback"><div align="center">
			<%
			if rs("ClassID")=0 then
				response.Write"<a href=index.asp?classid=0>δ����</a>"
			else
				dim c_rs
				set c_rs= Server.CreateObject(G_FS_RS)
				c_rs.open "select ClassID,ClassCName From FS_ME_InfoClass where UserNumber='"&Fs_User.UserNumber&"' and ClassTypes=7 and ClassID="&rs("ClassID"),User_Conn,1,3
				if not c_rs.eof then
					Response.Write "<a href=index.asp?classid="&rs("ClassID")&">"&c_rs("ClassCName")&"</a>"
					c_rs.close:set c_rs=nothing
				else
					response.Write"<a href=index.asp?classid=0>δ����</a>"
					c_rs.close:set c_rs=nothing
				end if
			end if
			%> </div></td>
            <td class="hback"><div align="center"><%=rs("addtime")%> </div></td>
            <td class="hback"> 
              <div align="center">
			  <%
			if rs("adminLock")=1 then
				Response.Write("<span class=tx>����Ա�����.���û�����</span>")
			else
				if rs("islock")=1 then
					response.Write("�û�����")
				else
					response.Write("����")
				end if
			end if
			%>
              </div></td>
            <td class="hback"><div align="center"><a href="PublicLogEdit.asp?id=<%=rs("iLogID")%>">�޸�</a>��<a href="index.asp?id=<%=rs("iLogID")%>&Action=Lock">����</a>��<a href="index.asp?id=<%=rs("iLogID")%>&Action=UnLock">����</a>��<a href="index.asp?id=<%=rs("iLogID")%>&Action=del">ɾ��</a> 
              </div></td>
          </tr>
          <%
			rs.movenext
		next
		%>
          <tr> 
            <td colspan="6" class="hback"> 
              <%
			response.Write "<p>"&  fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
		  end if
		  %>
              �� 
              <input name="Action" type="hidden" id="Action">
			  <input type="button" name="Submit" value="ɾ��"  onClick="document.myform.Action.value='del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.myform.submit();return true;}return false;}"> 
              <input type="button" name="Submit2" value="��������"  onClick="document.myform.Action.value='UnLock';{if(confirm('ȷ��������')){this.document.myform.submit();return true;}return false;}"> 
              <input name="Submit3" type="button"  onClick="document.myform.Action.value='Lock';{if(confirm('ȷ�������𣿣�\n�����󽫲�����ʾ')){this.document.myform.submit();return true;}return false;}" value="��������"> 
            </td>
          </tr>
        </form>
      </table>
       </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
</table>
</body>
</html>
<%
Set Fs_User = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = myform.elements[i];  
    if (e.name != 'chkall')  
       e.checked = myform.chkall.checked;  
    }  
	}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





