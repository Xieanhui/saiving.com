<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
	User_Conn.execute("Select UserSystemName From FS_ME_SysPara")
	if request("Action")="del" then
		if Request("id")="" then
			strShowErr = "<li>����Ĳ�����</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		else
			User_Conn.execute("Delete from FS_ME_Favorite where FavoID in ("&FormatIntArr(Request("id"))&")")
			strShowErr = "<li>ɾ���ɹ���</li>"
			Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Favorite.asp")
			Response.end
		end If
	Elseif Request("Action")="sort" Then
		if Request("id")="" Or Request("classID")="" then
			strShowErr = "<li>����Ĳ�����</li>"
			Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		Else
			User_Conn.execute("Update FS_ME_Favorite set FavoClassID="&CintStr(Request("ClassID"))&" where FavoID in ("&FormatIntArr(Request("id"))&")")
			strShowErr = "<li>ת�Ƴɹ�</li>"	
		    Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Favorite.asp")
			Response.end
		end If
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
	strpage=CintStr(request("page"))
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
<form name="myForm" method="post" action="">
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
            &gt;&gt; <a href="Favorite.asp">�ղؼй���</a> &gt;&gt; </td>
        </tr>
        <tr class="hback"> 
          <td  valign="top"><a href="Favorite.asp">ȫ��</a>��<a href="Favorite.asp?Type=0">����</a>��<%if IsExist_SubSys("DS") Then%><a href="Favorite.asp?Type=1">����</a>��<%end if%><a href="Favorite.asp?Type=2">��ҵ��Ա</a>��<%if IsExist_SubSys("SD") Then%><a href="Favorite.asp?Type=3">������Ϣ</a>��<%end if%><%if IsExist_SubSys("MS") Then%><a href="Favorite.asp?Type=4">��Ʒ</a>��<%end if%><%if IsExist_SubSys("HS") Then%><a href="Favorite.asp?Type=5">������Ϣ</a>��<%end if%><%if IsExist_SubSys("AP") Then%><a href="Favorite.asp?Type=6">��Ƹ</a>��<%end if%><a href="Favorite.asp?Type=7">��־</a>��<a href="Favorite_Class.asp">�ղؼ�(����)����</a></td>
        </tr>
        <tr class="hback">
          <td>
          <%
		  response.Write("	<table width=""98%"" align=center cellpadding=""2"" cellspacing=""1""><tr>")
		  dim t_k
		  t_k=0
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select ClassID,ClassCName,UserNumber From FS_ME_FavoriteClass where UserNumber='"&Fs_User.UserNumber&"'",User_Conn,1,1
		  do while not rs.eof 
		  	Response.Write("<td width=""24%"" valign=bottom><input type=""radio"" name=""classID"" value="""&rs("ClassID")&"""/><a href=""Favorite.asp?classid="&rs("ClassID")&""">"&rs("ClassCName")&"</a></td>")
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
       <%
	   
		  	dim o_type,o_class
		  	select case NoSqlHack(Request.QueryString("type"))
				case "1"
					o_type = " and FavoriteType=1"
				case "2"
					o_type = " and FavoriteType=2"
				case "3"
					o_type = " and FavoriteType=3"
				case "4"
					o_type = " and FavoriteType=4"
				case "5"
					o_type = " and FavoriteType=5"
				case "6"
					o_type = " and FavoriteType=6"
				case "7"
					o_type = " and FavoriteType=7"
				case "0"
					o_type = " and FavoriteType=0"
				case else
					o_type = ""
			end select
			if Request.QueryString("ClassId")<>"" then
				o_class=" and FavoClassID="&CintStr(Request.QueryString("ClassId"))&""
			else
				o_class=""
			end if
			set rs = Server.CreateObject(G_FS_RS)
			rs.open "select * from FS_ME_Favorite where UserNumber='"& Fs_User.UserNumber&"' "& o_type & o_class &" order by FavoID desc",User_Conn,1,1
			if rs.eof then
			   Response.Write"<tr  class=""hback""><td colspan=""6""  class=""hback"" height=""40"">û�м�¼��</td></tr>"
			else
			%>
          <tr> 
            <td width="4%" class="hback_1"><div align="center"> 
                <input name="chkall" type="checkbox" id="chkall" onClick="CheckAll(this.form);" value="checkbox" title="���ѡ�����л��߳�������ѡ��">
              </div></td>
            <td class="hback_1"><div align="left"><strong>��Ϣ����</strong></div></td>
            <td width="7%" class="hback_1"><div align="center"><strong>����</strong></div></td>
            <td width="17%" class="hback_1"><div align="center"><strong>����</strong></div></td>
            <td width="15%" class="hback_1"><div align="center"><strong>����</strong></div></td>
            <td width="12%" class="hback_1"><div align="center"><strong>����</strong></div></td>
          </tr>
          <%
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
          <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="center"> 
                <input name="id" type="checkbox" id="id" value="<%=rs("FavoID")%>">
              </div></td>
            <td class="hback"> 
              <%
			dim f_rs
			select case rs("FavoriteType")
				case 0
					set f_rs=Conn.execute("select id,NewsTitle From FS_NS_News where id="&CintStr(rs("FID")))
					if f_rs.eof then:Response.Write("<span class=tx>��Ϣ�Ѿ�������Աɾ��</span>"):else:response.Write"<a href=Public_info.asp?type=NS&Url="&rs("FID")&" target=_blank>"&f_rs("NewsTitle")&"</a>":end if
					f_rs.close:set f_rs=nothing
				case 1
					set f_rs=Conn.execute("select id,Name From FS_DS_List where id="&CintStr(rs("FID")))
					if f_rs.eof then:Response.Write("<span class=tx>��Ϣ�Ѿ�������Աɾ��</span>"):else:response.Write"<a href=Public_info.asp?type=DS&Url="&rs("FID")&" target=_blank>"&f_rs("Name")&"</a>":end if
					f_rs.close:set f_rs=nothing
				case 2
					response.Write"<a href=Public_info.asp?type=AP_1&Url="&rs("FID")&" target=_blank>�鿴</a>"
				case 3
					response.Write"<a href=../Supply/Supply.asp&id="&rs("FID")&" target=_blank>�鿴</a>"
				case 4
					response.Write"<a href=Public_info.asp?type=MS&Url="&rs("FID")&" target=_blank>�鿴</a>"
				case 5
					response.Write"<a href=../House/house.asp?ID="&rs("FID")&" target=_blank>�鿴</a>"
				case 6
					response.Write"<a href=Public_info.asp?type=AP_2&Url="&rs("FID")&" target=_blank>�鿴</a>"
				case 7
					response.Write"<a href=../Blog/Blog.asp?id="&rs("FID")&" target=_blank>�鿴</a>"
				case else
			end select
			%>
            </td>
            <td class="hback"><div align="center"> 
                <%
			select case rs("FavoriteType")
				case 0
					response.Write"<a href=Favorite.asp?type=0>����</a>"
				case 1
					response.Write"<a href=Favorite.asp?type=1>����</a>"
				case 2
					response.Write"<a href=Favorite.asp?type=2>��ҵ</a>"
				case 3
					response.Write"<a href=Favorite.asp?type=3>����</a>"
				case 4
					response.Write"<a href=Favorite.asp?type=4>��Ʒ</a>"
				case 5
					response.Write"<a href=Favorite.asp?type=5>����</a>"
				case 6
					response.Write"<a href=Favorite.asp?type=6>��Ƹ</a>"
				case 7
					response.Write"<a href=Favorite.asp?type=6>��־</a>"
				case else
					response.Write"<a href=Favorite.asp>-</a>"
			end select
			
			%>
              </div></td>
            <td class="hback"><div align="center"><%=rs("AddTime")%></div></td>
            <td class="hback"> <div align="center">
			<%
			if rs("FavoClassID")=0 then
				response.Write"<a href=Favorite.asp?ClassID=0>δ����</a>"
			else
				dim crs
				set crs=user_Conn.execute("select ClassID,ClassCName,UserNumber From FS_ME_FavoriteClass where ClassID="&rs("FavoClassID"))
				Response.Write "<a href=""Favorite.asp?ClassID="&crs("ClassID")&""">"&crs("ClassCName")&"</a>"
			end if
			%> 
                </div></td>
            <td class="hback"><div align="center">
			<a href="Favorite.asp?Action=del&id=<%=rs("FavoID")%>" onClick="{if(confirm('ȷ��Ҫɾ����?')){return true;}return false;}">ɾ��</a>
			| <a href="#" onclick="sort('<%=rs("FavoID")%>')">ת��</>
			</div></td>
          </tr>
          <%
			  rs.movenext
		  next
		  %>
          <tr  class="hback"> 
            <td height="31" colspan="6">
			<div align="right"> 
            <%
			response.Write "<p>"&  fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
			%>
                <input name="Action" type="hidden" id="Action">
                <input type="button" name="Submit" value="ɾ��"  onClick="document.myForm.Action.value='del';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.myForm.submit();return true;}return false;}">
				 <input type="button" name="Submit" value="ת��"  onClick="document.myForm.Action.value='sort';this.document.myForm.submit();">
              </div></td>
          </tr>
          <% 
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
function sort(id)
{
	var elements=document.all("classID")
	var classid="";
	if(elements.length=='undefined')
	{
		classid=elements.value;
	}
	else
	{
		for(var i=0;i<elements.length;i++)
		{
			if (elements[i].checked)
			{
				classid=elements[i].value;
			}
		}
	}
	if (classid=="")
	{
		alert("��ѡ��Ŀ����࣡");
		return false;
	}
	location.href="Favorite.asp?Action=sort&id="+id+"&classiD="+classid;
}
</script>
<%
Set Fs_User = Nothing
rs.close
set rs=nothing
set conn=nothing
set User_Conn=nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





