<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
dim ShowChar,strAction,ShowChar_1,str_m_type
str_m_type = CintStr(Request.QueryString("M_type"))
if isnull(str_m_type) or not isnumeric(str_m_type) or trim(str_m_type)="" then
	strShowErr = "<li>�������</li>"
	Response.Redirect("lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
end if
ShowChar = "�ҵ�����"
ShowChar_1 = "������"
strAction = ""
If Request.Form("Action") = "Del" then
	Dim DelID,Str_Tmp,Str_Tmp1
	DelID = request.Form("BookID")
	if DelID = "" then 
		strShowErr = "<li>�����ѡ��һ����ɾ��</li>"
		Call ReturnError(strShowErr,"")
	End if
	User_Conn.execute("Delete From FS_ME_Book where BookId in ("&FormatIntArr(DelID)&") and M_ReadUserNumber ='"& Fs_User.UserNumber&"'")
	strShowErr = "<li>ɾ�����Գɹ�</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
End if
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage,i
Dim tmp_re,tmp_er
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
strpage=NoSqlHack(request("page"))
If Trim(strpage)="" Then
	strpage="1"
ElseIf len(strpage)=0 Or strpage<"1" Then
	strpage="1"
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-����</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ���ԣ�<% = ShowChar %>
            &lt;&lt; 
            <%
			select case Request.QueryString("M_type")
					case "0"
						Response.Write("��Ա����")
					case "1"
						Response.Write("��������")
					case "2"
						Response.Write("��������")
					case "3"
						Response.Write("��ְ��Ƹ����")
					case "4"
						Response.Write("��������")
					case "5"
						Response.Write("��������")
			end select
			%>
            &gt;&gt; </td>
          </tr>
        </table>
        
      <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr class="hback"> 
            <td colspan="12" class="hback"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="hback"> 
                  <td width="27%" class="hback"> <%
				Dim RsBookObj,RsBookSQL
				Dim strSQLs
				strpage=request("page")
				if len(strpage)=0 Or strpage<1 or trim(strpage)=""  Then strpage="1"
				Set RsBookObj = Server.CreateObject(G_FS_RS)
				RsBookSQL = "Select BookID,M_Title,M_FromUserNumber,M_type,M_ReadUserNumber,M_Content,M_FromDate,M_ReadTF,LenContent From FS_ME_Book  where M_ReadUserNumber='"&Fs_User.UserNumber&"' and M_Type="& str_m_type &" Order by BookID desc"
				RsBookObj.Open RsBookSQL,User_Conn,1,3
				%>
                    �ռ�ռ��: 
                    <%
				     Dim UnTotle,FS_Book_1
					 Set FS_Book_1 = new Cls_Message
					UnTotle=FS_Book_1.LenbContent(Fs_User.UserNumber)/(1024*200)*100
					Set FS_Book_1 = Nothing 
					If IsNull(UnTotle) then UnTotle=0
					Response.Write Formatnumber(UnTotle,2,-1)&"%"
					%>
                    (��200K)</div></td>
                  <td width="73%" class="hback"> <table width="100%" height="17" border="0" cellpadding="0" cellspacing="1" class="table">
                      <tr> 
                        <td class="hback_1"><img src="images/space_pic_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.gif" width="<% = Formatnumber((UnTotle),2,-1)%>%" height="17"></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr class="hback">
          <tr class="hback"> 
            <td width="5%" height="22" class="xingmu"><div align="left"><strong>�Ѷ�</strong></div></td>
            <td width="15%" class="xingmu"><strong> 
              <% = ShowChar_1 %>
              </strong></td>
            <td width="36%" height="22" class="xingmu"><div align="left"><strong>����</strong></div></td>
            <td width="20%" height="22" class="xingmu"><div align="left"><strong>����</strong></div></td>
            <td width="11%" height="22" class="xingmu"><strong>����</strong></td>
            <td width="7%" height="22" class="xingmu"><div align="center">�鿴</div></td>
            <td width="6%" height="22" class="xingmu"><div align="center"><strong>����</strong></div></td>
          </tr>
          <%
			if RsBookObj.eof then
			   RsBookObj.close
			   set RsBookObj=nothing
			   Response.Write"<tr  class=""hback""><td colspan=""7""  class=""hback"" height=""40"">û�����ԡ�</td></tr>"
			else
				RsBookObj.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("Page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo<=0 Then cPageNo=1
				If cPageNo>RsBookObj.PageCount Then cPageNo=RsBookObj.PageCount 
				RsBookObj.AbsolutePage=cPageNo
				for i=1 to RsBookObj.pagesize
				  if RsBookObj.eof Then exit For 
					Dim Returvaluestr_R,Returvaluestr_F,strbstat,strben,strcss,strReadTF
					if RsBookObj("M_ReadTF") =0 then 
						strbstat = "<b>"
						strben = "</b>"
						strcss = "hback"
						strReadTF = "<font color=red><b>��</b></font>"
					Else
						strbstat = ""
						strben = ""
						strcss = "hback"
						strReadTF = "<font color=#999999><b>��</b></font>"
					End if
					Returvaluestr_R = Fs_User.GetFriendName(RsBookObj("M_ReadUserNumber"))
					if Trim(RsBookObj("M_FromUserNumber")) <> "0" then
						Returvaluestr_F = "<a href=ShowUser.asp?UserNumber="& RsBookObj("M_FromUserNumber") &" target=""_blank"">"&Fs_User.GetFriendName(RsBookObj("M_FromUserNumber"))&"</a>"
					Else
						Returvaluestr_F = "�û�������"
					End if
		%>
          <tr class="hback"> 
            <td height="31" class="<% = strcss %>"> 
              <div align="center">
                <% = strReadTF%>
              </div></td>
            <td class="<% = strcss %>"> <% =   Returvaluestr_F %> </td>
            <td class="<% = strcss %>"><a href="Book_Read.asp?BookId=<%=RsBookObj("BookId")%>&M_type=<%=RsBookObj("M_type")%>"><% = strbstat & RsBookObj("M_title") & strben %></a></td>
            <td class="<% = strcss %>"><% =  RsBookObj("M_FromDate")  %></td>
            <td class="<% = strcss %>"><% =  RsBookObj("LenContent")  %>
              Byte</td>
            <td class="<% = strcss %>"> <div align="center"> 
                <%
				Response.Write "<a href=""Book_Read.asp?BookID="& RsBookObj("BookID") &"&M_Type="&RsBookObj("M_type")&""">�ظ� </a>"
				%>
              </div></td>
            <td class="<% = strcss %>"><input name="BookID" type="checkbox" id="BookID" value="<% = RsBookObj("BookID")%>"></td>
          </tr>
          <%
			  RsBookObj.MoveNext
		  Next
		  %>
          <tr class="hback"> 
            <td colspan="12"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr  class="hback"> 
                  <td colspan="2"> <%
					response.Write "<p>"&  fPageCount(RsBookObj,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
					%> <div align="right"> </div></td>
                </tr>
                <tr  class="hback"> 
                  <td width="64%"><div align="right">��ʡÿһ�ֿռ䣬�뼰ʱɾ��������Ϣ 
                      <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
                      ѡ���������� 
                      <input name="Action" type="hidden" id="Action" value="Del">
                      <input name="strAction" type="hidden" id="strAction" value="<% = strAction%>">
                      �� </div></td>
                  <td width="18%"><input type="submit" name="Submit" value="ɾ��ѡ�е�����" onClick="{if(confirm('ȷ���������ѡ��ļ�¼��?')){this.document.form1.submit();return true;}return false;}"> 
                  </td>
                </tr>
              </table></td>
          </tr>
          <%end if%>
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
<%
Set Fs_User = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = form.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form.chkall.checked;  
    }  
  }
</script>


<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





