<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
Dim strpage,int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
dim obj_nothing_Rs,i,str_action,str_id,str_Url_down,rs_d_obj
str_action = NoSqlHack(Request("Action"))
str_id = CintStr(Request.QueryString("id"))
if str_action = "Down" then
	set rs_d_obj = User_Conn.execute("select * From FS_ME_getThing where UserNumber='"&NoSqlHack(Fs_User.UserNumber)&"' and UserDel=0 and id="&CintStr(str_id))
	if rs_d_obj.eof then
		strShowErr = "<li>�Ҳ�����¼</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
		rs_d_obj.close:set rs_d_obj = nothing
	else
		if rs_d_obj("useNum")>=rs_d_obj("MaxNum") then
			strShowErr = "<li>���Ѿ�������"&rs_d_obj("MaxNum")&"��,����������!</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
			rs_d_obj.close:set rs_d_obj = nothing
		end if
		if rs_d_obj("isLock")=1 then
			strShowErr = "<li>�˼�¼�Ѿ�������</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
			rs_d_obj.close:set rs_d_obj = nothing
		else
			'�������ݿ�
			DIM up_d_rs
			set up_d_rs = Server.CreateObject(G_FS_RS)
			up_d_rs.open "select * From FS_ME_getThing where UserNumber='"&Fs_User.UserNumber&"' and id="&CintStr(str_id),User_conn,1,3
			up_d_rs("isUse")=1
			up_d_rs("useNum")=up_d_rs("useNum")+1
			up_d_rs("UpdateTime")=now
			up_d_rs("IP")=NoSqlHack(Request.ServerVariables("Remote_Addr"))
			up_d_rs.update
			up_d_rs.close:set up_d_rs=nothing
			Response.Redirect rs_d_obj("URL_1")
			rs_d_obj.close:set rs_d_obj = nothing
			response.end
		end if
	end if
elseif str_action = "Del" then
	User_Conn.execute("Update FS_ME_getThing set UserDel=1 where UserNumber='"&Fs_User.UserNumber&"' and Id="&CintStr(str_Id))
	strShowErr = "<li>ɾ���ɹ�</li>"
	Response.Redirect("lib/success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Get_Thing.asp")
	Response.end
end if
strpage=request("page")
if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
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
Set obj_nothing_Rs = server.CreateObject(G_FS_RS)
SQL = "Select *  from FS_ME_getThing where UserNumber='"&Fs_User.UserNumber&"' and islock=0 and UserDel=0 Order by id desc"
obj_nothing_Rs.Open SQL,User_Conn,1,3
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-��ȡ��Ʒ</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
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
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="hback"><strong>λ�ã�</strong><a href="../">��վ��ҳ</a> &gt;&gt; 
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ��ȡ��Ʒ </td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback">
          <td colspan="7" class="xingmu">��ȡ��Ʒ</td>
        </tr>
        <tr class="hback"> 
          <td width="25%" class="hback"><div align="left"><strong>����</strong></div></td>
          <td width="13%" class="hback"><div align="left"><strong>�汾</strong></div></td>
          <td width="13%" class="hback"><div align="center"><strong>�ͺ�</strong></div></td>
          <td width="11%" class="hback"><div align="center"><strong>������</strong></div></td>
          <td width="10%" class="hback"><div align="center"><strong>�������</strong></div></td>
          <td width="18%" class="hback"><div align="center"><strong>�������ʱ��</strong></div></td>
          <td width="10%" class="hback"><div align="center"><strong>����</strong></div></td>
        </tr>
		<%
		if obj_nothing_Rs.eof then
		   obj_nothing_Rs.close
		   set obj_nothing_Rs=nothing
		   Response.Write"<tr  class=""hback""><td colspan=""6""  class=""hback"" height=""40"">û�м�¼��</td></tr>"
		else
			obj_nothing_Rs.PageSize=int_RPP
			cPageNo=NoSqlHack(Request.QueryString("Page"))
			If cPageNo="" Then cPageNo = 1
			If not isnumeric(cPageNo) Then cPageNo = 1
			cPageNo = Clng(cPageNo)
			If cPageNo<=0 Then cPageNo=1
			If cPageNo>obj_nothing_Rs.PageCount Then cPageNo=obj_nothing_Rs.PageCount 
			obj_nothing_Rs.AbsolutePage=cPageNo
			for i=1 to obj_nothing_Rs.pagesize
				if obj_nothing_Rs.eof Then exit For 
		%>
        <tr class="hback"> 
          <td class="hback"><a href="#" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(down_<% = obj_nothing_Rs("id")%>);" onMouseOver="this.className='bg'" onMouseOut="this.className='bg1'" language=javascript><% = obj_nothing_Rs("ProductID")%></a></td>
          <td class="hback"><% = obj_nothing_Rs("Version")%></td>
          <td class="hback"><div align="center">
            <% = obj_nothing_Rs("PType")%>
          </div></td>
          <td class="hback"><div align="center">
            <% = obj_nothing_Rs("useNum")%>��</div></td>
          <td class="hback"><div align="center">
            <% = obj_nothing_Rs("MaxNum")%>��</div></td>
          <td class="hback"><div align="center">
            <% = obj_nothing_Rs("UpdateTime")%>
          </div></td>
          <td class="hback"><div align="center"><a href="get_Thing.asp?Action=Down&Id=<%=obj_nothing_Rs("id")%>">����</a>|<a href="get_Thing.asp?Action=Del&Id=<%=obj_nothing_Rs("id")%>" onClick="{if(confirm('ȷ��Ҫɾ����?\nɾ���󽫲��ָܻ�!!')){return true;}return false;}">ɾ��</a></div></td>
        </tr>
         <tr class="hback" id="down_<% = obj_nothing_Rs("id")%>" style="display:none;">
           <td colspan="7" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="5" class="hback_1">
             <tr>
               <td width="31%">�������IP��
                 <% = obj_nothing_Rs("IP")%></td>
               <td width="69%">������ڣ�
                 <% = obj_nothing_Rs("AddTime")%></td>
             </tr>
             <tr>
               <td colspan="2">������
                <% = obj_nothing_Rs("Content")%></td>
              </tr>
           </table>             </td>
         </tr>
		 <%
				obj_nothing_Rs.movenext
			Next
		 %>
		<tr class="hback"> 
          <td colspan="7" class="hback">
			<%
					response.Write "<p>"&  fPageCount(obj_nothing_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
				   obj_nothing_Rs.close
				   set obj_nothing_Rs=nothing
			end if
			%>	
		 </td>
        </tr>
      </table></td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
</body>
</html>
<%
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





