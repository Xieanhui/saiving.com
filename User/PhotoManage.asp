<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
if request("Action")="del" then
	if Request("id")="" then
		strShowErr = "<li>����Ĳ�����</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Delete from FS_ME_Photo where ID in ("&FormatIntArr(Request("id"))&")")
		strShowErr = "<li>ɾ���ɹ���</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../PhotoManage.asp")
		Response.end
	end if
end if
if request("Action")="delall" then
	if Request("chkall")="" then
		strShowErr = "<li>����Ĳ�����</li>"
		Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		User_Conn.execute("Delete from FS_ME_PhotoClass where UserNumber='"&Fs_User.UserNumber&"'")
		User_Conn.execute("Delete from FS_ME_Photo")
		strShowErr = "<li>ɾ���ɹ���</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../PhotoManage.asp")
		Response.end
	end if
	chkall
end if
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,rs,i
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
'if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-������</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="lib/css/lightbox.css" type="text/css" media="screen" />
<script type="text/javascript" src="../FS_INC/prototype.js"></script>
<script type="text/javascript" src="lib/js/scriptaculous.js?load=effects"></script>
<script type="text/javascript" src="lib/js/lightbox.js"></script>
<head>
<body onLoad="initLightbox()">
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; <a href="PhotoManage.asp">������</a> 
            &gt;&gt;</td>
        </tr>
      </table> 
		  
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr>
          <td class="hback"><a href="PhotoManage.asp">�����ҳ</a>��<a href="Photo_add.asp">����ͼƬ</a>��<a href="PhotoManage.asp?isRec=1">���Ƽ���ͼƬ</a>��<a href="Photo_filt.asp">�õ�Ƭ����</a>��<a href="Photo_Class.asp">������</a></td>
        </tr>
        <tr> 
          <td class="hback"> 
            <%
		  response.Write("	<table width=""98%"" align=center cellpadding=""2"" cellspacing=""1""><tr>")
		  dim t_k,rec_str
		  t_k=0
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select id,title,UserNumber From FS_ME_PhotoClass where UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,3
		  do while not rs.eof 
		  	Response.Write("	<td width=""24%"" valign=bottom><img src=""images/folderopened.gif""></img><a href=PhotoManage.asp?classid="&rs("id")&">"&rs("title")&"</a></td>")
		  rs.movenext
		  t_k = t_k+1
		  if t_k mod 4 =0 then
		  	Response.Write("	</tr>")
		  end if
		  loop
		  response.Write("	</table>")
		  rs.close:set rs=nothing
		  %>
          </td>
        </tr>
      </table> 
      
        <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <form name="myform" method="post" action="">
		  <%
		  dim o_class
		  if NosqlHack(Request.QueryString("Classid"))<>"" then
		  	if not isnumeric(NosqlHack(Request.QueryString("Classid"))) then
				strShowErr = "<li>����Ĳ�����</li>"
				Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			end if
		  	o_class=" and classid="&CintStr(Request("Classid"))&""
		  else
		  	o_class =""
		  end if
		  if NoSqlHack(Request.QueryString("isRec"))="1" then
		  	 rec_str = " and isRec=1"
		  else
		  	 rec_str = ""
		  end if
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select * from FS_ME_Photo where UserNumber='"& Fs_User.UserNumber&"' "&o_class&rec_str&" order by id desc",User_Conn,1,1
			if rs.eof then
			   rs.close
			   set rs=nothing
			   Response.Write"<tr  class=""hback""><td colspan=""1""  class=""hback"" height=""40"">û�м�¼��</td></tr>"
			else
				rs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("Page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo>rs.PageCount Then cPageNo=rs.PageCount 
				If cPageNo<=0 Then cPageNo=1
				rs.AbsolutePage=cPageNo
				for i=1 to rs.pagesize
					if rs.eof Then exit For
			%>
          <tr class="hback"> 
            <td width="21%" rowspan="5" class="hback_1">
			<table border="0" align="center" cellpadding="2" cellspacing="1" class="table">
                <tr> 
                  <td class="hback"><%if isnull(trim(rs("PicSavePath"))) then%><img src="Images/nopic_supply.gif" width="90"  id="pic_p_1"><%else%><a href="<%=rs("PicSavePath")%>" rel="lightbox" title="<%=rs("title")%>"><img src="<%=rs("PicSavePath")%>" width="90" border="0" id="pic_p_1"></a><%end if%>
                  </td>
                </tr>
              </table></td>
            <td width="12%" class="hback"><div align="center"><strong>��Ƭ���ƣ�</strong></div></td>
            <td width="40%" class="hback"><%if rs("isRec")=1 then response.Write"<span class=""tx"">[��ͼƬ�Ѿ����Ƽ�]</span>"%><font style="font-size:14px"><span class="hback_1"><strong><%=rs("title")%></strong>;</span></font></td>
            <td width="10%"><div align="center">���������</div></td>
            <td width="17%"><%=rs("hits")%></td>
          </tr>
          <tr class="hback"> 
            <td><div align="center">�������ڣ�</div></td>
            <td><%=rs("Addtime")%></td>
            <td><div align="center">ͼƬ��С��</div></td>
            <td><%=rs("PicSize")%>byte</td>
          </tr>
          <tr class="hback"> 
            <td><div align="center">��Ƭ��ַ��</div></td>
            <td colspan="3"><%=rs("PicSavePath")%></td>
          </tr>
          <tr class="hback"> 
            <td><div align="center">��Ƭ������</div></td>
            <td colspan="3"><%=rs("Content")%></td>
          </tr>
          <tr class="hback"> 
            <td><div align="center">��Ƭ���ࣺ</div></td>
            <td><%
			dim c_rs
			if rs("ClassID")=0 then
				response.Write("û����")
			else
				set c_rs=User_Conn.execute("select ID,title From Fs_me_photoclass where id="&rs("classid"))
				Response.Write "<a href=PhotoManage.asp?ClassiD="&c_rs("ID")&">"&c_rs("title")&"</a>"
				c_rs.close:set c_rs=nothing
			end if
			%></td>
            <td colspan="2"><div align="center"><a href="Photo_Edit.asp?Id=<%=rs("id")%>">�޸�</a>��<a href="PhotoManage.asp?id=<%=rs("id")%>&Action=del" onClick="{if(confirm('ȷ��ͨ��ɾ����')){return true;}return false;}">ɾ��</a> 
                <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>">
              </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="5" height="3" class="xingmu"></td>
          </tr>
          <%
			rs.movenext
		next
		%>
          <tr class="hback"> 
            <td colspan="5"> 
              <%
			response.Write "<p>"&  fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
		  end if
		  %>
              ��
<input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="���ѡ�����л��߳�������ѡ��">
              ȫѡ��
              <input name="Action" type="hidden" id="Action"> <input type="button" name="Submit1" value="ɾ��"  onClick="document.myform.Action.value='delall';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.myform.submit();return true;}return false;}"></td>
          </tr></form>
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





