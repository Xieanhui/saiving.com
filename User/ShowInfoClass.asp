<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
		Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i_ii
		Dim SQL_list,rs_l,str_title,str_id,str_addtime,str_lockTF,type_s,novalue
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
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-ר������</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ר��</td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td colspan="2" class="xingmu">ר������</td>
        </tr>
        <tr class="hback"> 
          <%
		  dim rs,classid,usernumber,s_type
		  classid = CintStr(Request.QueryString("ClassID"))
		  usernumber = NoSqlHack(Request.QueryString("UserNumber"))
		  if classid="" or not isnumeric(classid) or usernumber="" then
				strShowErr = "<li>����Ĳ���</li>"
				Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		  end if
		  set rs = Server.CreateObject(G_FS_RS)
		  rs.open "select ClassID,ClassCName,ClassEName,ParentID,UserNumber,ClassTypes,AddTime,ClassContent From FS_ME_InfoClass where Classid="& classid &" and UserNumber='"& usernumber &"'",User_Conn,1,3
		  if rs.eof then
				strShowErr = "<li>�Ҳ�����¼</li>"
				Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
		  end if
		  %>
          <td colspan="2" class="hback"> </td>
        </tr>
        <tr class="hback"> 
          <td width="20%" class="hback_1"><div align="right">ר������</div></td>
          <td width="80%" class="hback"><%=rs("ClassCName")%></td>
        </tr>
        <tr class="hback"> 
          <td class="hback_1"><div align="right">ר����������</div></td>
          <td class="hback"><%=rs("addtime")%></td>
        </tr>
        <tr class="hback"> 
          <td class="hback_1"><div align="right">ר����������</div></td>
          <td class="hback">
		  <%
				select case rs("ClassTypes")
					case 0
						s_type = "������ѧ��"
					case 1
						s_type = "����"
					case 2
						s_type = "��Ʒ"
					case 3
						s_type = "����"
					case 4
						s_type = "����"
					case 5
						s_type = "��ְ"
					case 6
						s_type = "��Ƹ"
					case 7
						s_type = "������־"
				end select
				Response.Write s_type
		 %>
		 </td>
        </tr>
        <tr class="hback">
          <td height="38" class="hback_1"> 
            <div align="right">ר������</div></td>
          <td class="hback"><%=rs("ClassContent")%></td>
        </tr>
      </table>
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr> 
          <td colspan="4" class="hback_1">ר�����б�<span class="tx">(���Ҫ�޸ģ�������Ӧ��ϵͳ��ȥ�޸�)</span></td>
        </tr>
        <tr class="hback_1"> 
          <td width="38%"><div align="center"><strong>����</strong></div></td>
          <td width="35%"><div align="center"><strong>��������</strong></div></td>
          <td width="16%"><div align="center"><strong>����</strong></div></td>
          <td width="11%"><div align="center"><strong>״̬</strong></div></td>
        </tr>
        <%
			strpage=request("page")
			if len(strpage)=0 Or strpage<1 or trim(strpage)=""Then:strpage="1":end if
			set rs_l= Server.CreateObject(G_FS_RS)
			select case rs("ClassTypes")
				case 0
					SQL_list = "select ContID,ContTitle,AddTime,AuditTF From FS_ME_InfoContribution where UserNumber='"& usernumber &"' and ClassID="&rs("ClassID")&""
					rs_l.open SQL_list,User_Conn,1,3
					if rs_l.eof then
						call no_s()
					else
						call no_s_1()
					end if
				case 1
					 
				case 2
					 
				case 3
					 
				case 4
					SQL_list = "select ID,PubTitle,AddTime,isPass,MyClassID From FS_SD_News where UserNumber='"& usernumber &"' and MyClassID="&rs("ClassID")&""
					rs_l.open SQL_list,Conn,1,3
					if rs_l.eof then
						call no_s()
					else
						call no_s_1()
					end if
				case 5
					 
				case 6
				
				case 7
					 
			end select
			sub no_s()
				   rs_l.close
				   set rs_l=nothing
				   Response.Write"<tr  class=""hback""><td colspan=""4""  class=""hback"" height=""40"">û�м�¼��</td></tr>"
			end sub
			sub no_s_1()
						rs_l.PageSize=int_RPP
						cPageNo=NoSqlHack(Request.QueryString("Page"))
						If cPageNo="" Then cPageNo = 1
						If not isnumeric(cPageNo) Then cPageNo = 1
						cPageNo = Clng(cPageNo)
						If cPageNo<=0 Then cPageNo=1
						If cPageNo>rs_l.PageCount Then cPageNo=rs_l.PageCount 
						rs_l.AbsolutePage=cPageNo
						for i_ii=1 to rs_l.pagesize
							if rs_l.eof Then exit For 
						  select case rs("ClassTypes")
								case 0
									str_title=rs_l("ContTitle")
									str_id = rs_l("ContID")
									str_addtime =rs_l("AddTime")
									str_lockTF=rs_l("AuditTF")
									type_s = "Ͷ��"
								case 1
									 
								case 2
									 
								case 3
									 
								case 4
									str_title=rs_l("PubTitle")
									str_id = rs_l("ID")
									str_addtime =rs_l("AddTime")
									str_lockTF=rs_l("isPass")
									type_s = "����"
								case 5
									 
								case 6
								
								case 7
									 
							end select
	%>
        <tr class="hback"> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="7%"><img src="Images/folderopened.gif" width="16" height="16"></td>
                <td width="93%" valign="bottom"> <% = str_title %> </td>
              </tr>
            </table></td>
          <td><div align="center"> 
              <% = str_addtime %>
            </div></td>
          <td><div align="center"> 
              <% = type_s %>
            </div></td>
          <td><div align="center"> 
              <% 
			  select case rs("ClassTypes")
					case 0
						if str_lockTF=1 then:response.Write("�����"):else:response.Write("<span class=""tx"">δ���</span>"):end if
					case 1
						 
					case 2
						 
					case 3
						 
					case 4
						if str_lockTF=1 then:response.Write("�����"):else:response.Write("<span class=""tx"">δ���</span>"):end if
					case 5
						 
					case 6
					
					case 7
						 
				end select
			    %>
            </div></td>
        </tr>
        <%
			rs_l.movenext
		next
		%>
        <tr class="hback"> 
          <td colspan="4"><div align="right"> 
              <%response.Write "<p>"&  fPageCount(rs_l,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo) %>
            </div></td>
        </tr>
        <%
		rs_l.close:set rs_l=nothing
		end sub
		%>
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
rs.close:set rs=nothing
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





