<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="../FS_Inc/Md5.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-�����Ա��(��Ա����,���)</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; �����Ա��(��Ա����,���)</td>
        </tr>
      </table> 
      
        
      
      <form name="form1" method="post" action="Pay.asp" onSubmit="return checkinput();">
        <table width="98%" height="119" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr> 
            <td class="hback" height="30"> 
              <%
			Dim RsBuyObj
			Set RsBuyObj = server.CreateObject(G_FS_RS)
			RsBuyObj.open "select  GroupName, GroupID From FS_ME_Group  where GroupMoney>0 Order by GroupID desc",User_Conn,1,3
			%> <select name="GroupID" id="GroupID">
                <%Do while Not RsBuyObj.eof %>
                <option value="<%= RsBuyObj("GroupID")%>"><%= RsBuyObj("GroupName")%></option>
                <%
					RsBuyObj.movenext
				Loop
			RsBuyObj.Close
			set RsBuyObj = nothing
				%>
              </select>
            </td>
          </tr>
          <tr> 
            <td height="30" class="hback"> 
              <input type="submit" name="Submit" value="ȷ�Ϲ���,ת��֧��ҳ��">
            </td>
          </tr>
          <tr> 
            <td height="30" class="hback"  id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(mid100100100);"  language=javascript>���鿴ÿ����Ա���Ȩ��</td>
          </tr>
        </table>
        <table width="98%" height="145" border="0" align="center" cellpadding="5" cellspacing="1" class="table"  id="mid100100100" style="display:none;">
          <tr> 
            <td class="hback">&nbsp;</td>
          </tr>
        </table>
      </form> </td>
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
function checkinput()
{
	if(document.form1.GroupID.value=='')
	{
		alert('û�л�Ա��');
		return false;
	}
}

</script>

<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





