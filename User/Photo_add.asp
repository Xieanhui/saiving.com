<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<!--#include file="../FS_Inc/Func_Page.asp" -->
<%
Dim str_CurrPath
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-������</title>
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
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; <a href="PhotoManage.asp">������</a> 
            &gt;&gt;�������</td>
        </tr>
      </table> 
		  
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr> 
          <td class="hback"><a href="PhotoManage.asp">�����ҳ</a>��<a href="Photo_add.asp">����ͼƬ</a>��<a href="PhotoManage.asp?isRec=1">���Ƽ���ͼƬ</a>��<a href="Photo_filt.asp">�õ�Ƭ����</a>��<a href="Photo_Class.asp">������</a></td>
        </tr>
      </table>
      
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="s_form" method="post" action="Photo_save.asp" onSubmit="return checkinput();">
          <tr> 
            <td colspan="2" class="xingmu">�������</td>
          </tr>
          <tr> 
            <td width="18%" class="hback"> <div align="right">��Ƭ���⣺</div></td>
            <td width="82%" class="hback"><input name="title" type="text" id="title" size="45"></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">ͼƬ��</div></td>
            <td class="hback"><table width="81%" border="0" cellspacing="1" cellpadding="5">
                <tr> 
                  <td width="33%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <td class="hback"><img src="Images/nopic_supply.gif" width="90" height="90" id="pic_p_1"></td>
                        </tr>
                      </table>
                      <input name="pic_1" type="hidden" id="pic_1" size="40" >
                    </div></td>
                  <td width="34%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <td class="hback"><img src="Images/nopic_supply.gif" width="90" height="90" id="pic_p_2"></td>
                        </tr>
                      </table>
                      <input name="pic_2" type="hidden" id="pic_2" size="40">
                    </div></td>
                  <td width="33%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <td class="hback"><img src="Images/nopic_supply.gif" width="90" height="90" id="pic_p_3"></td>
                        </tr>
                      </table>
                      <input name="pic_3" type="hidden" id="pic_3" size="40">
                    </div></td>
                </tr>
                <tr> 
                  <td><div align="center"><img  src="Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_1);" style="cursor:hand;"> 
                      ��<img src="Images/del_supply.gif" width="44" height="22" onClick="dels_1();" style="cursor:hand;"> 
                    </div></td>
                  <td><div align="center"><img  src="Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_2);" style="cursor:hand;"> 
                      ��<img src="Images/del_supply.gif" width="44" height="22" onClick="dels_2();" style="cursor:hand;"> 
                    </div></td>
                  <td><div align="center"><img  src="Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_3);" style="cursor:hand;"> 
                      ��<img src="Images/del_supply.gif" width="44" height="22" onClick="dels_3();" style="cursor:hand;"> 
                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">���</div></td>
            <td class="hback"><select name="Classid">
                <option value="0">ѡ��������</option>
                <%
				dim rs
				set rs=User_Conn.execute("select id,title From FS_ME_PhotoClass where UserNumber='"&session("FS_UserNumber")&"' order by id desc")
				do while not rs.eof
						response.Write"		<option value="""&rs("id")&""">"&rs("title")&"</option>"&chr(13)
					rs.movenext
				loop
				rs.close:set rs=nothing
				%>
              </select></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right">ͼƬ˵����</div></td>
            <td class="hback"><textarea name="content" rows="8" id="content" style="width:80%"></textarea></td>
          </tr>
          <tr> 
            <td class="hback"><div align="right"></div></td>
            <td class="hback"><input type="submit" name="Submit" value="����ͼƬ�����">
              <input name="Action" type="hidden" id="Action" value="add"></td>
          </tr>
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
new Form.Element.Observer($('pic_1'),1,pics_1);
	function pics_1()
		{
			if ($('pic_1').value=='')
			{
				$('pic_p_1').src='Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_1').src=$('pic_1').value
			}
		} 
new Form.Element.Observer($('pic_2'),1,pics_2);
	function pics_2()
		{
			if($('pic_2').value=='')
			{
			$('pic_p_2').src='Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_2').src=$('pic_2').value
			}
		} 
new Form.Element.Observer($('pic_3'),1,pics_3);
	function pics_3()
		{
			if($('pic_3').value=='')
			{
			$('pic_p_3').src='Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_3').src=$('pic_3').value
			}
		}
function dels_1()
	{
		document.s_form.pic_1.value=''
	}
function dels_2()
	{
		document.s_form.pic_2.value=''
	}
function dels_3()
	{
		document.s_form.pic_3.value=''
	}
function checkinput()
{
	if(document.s_form.title.value=='')
	{
		alert('��д������');
		document.s_form.title.focus();
		return false;
	}
	if(document.s_form.pic_1.value=='')
	{
		alert('��д��������һ��ͼƬ��ַ');
		//document.s_form.pic_1.focus();
		return false;
	}
	if(document.s_form.content.value=='')
	{
		alert('��дͼƬ����');
		document.s_form.content.focus();
		return false;
	}
}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





