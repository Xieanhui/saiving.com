<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-��Ƹ����</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
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
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr class="hback"> 
            
          <td class="hback"><strong>λ�ã�</strong><a href="../../">��վ��ҳ</a> &gt;&gt; 
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; ��Ƹ</td>
          </tr>
        </table>
<%Dim AP_Rs,isValidate,GroupLevel,BeginDate,EndDate
isValidate = -1
set AP_Rs = User_Conn.execute("select C_Name from FS_ME_CorpUser where UserNumber='"&session("FS_UserNumber")&"'")
if not AP_Rs.eof then
	session("C_Name") = AP_Rs(0)
else
	session("C_Name") = ""
end if
AP_Rs.close
if session("C_Name")="" then 
	response.Write("�㲻�ǹ�˾��û��Ȩ���������ҳ�档������Ҫ������ע�ᡣ")	
else	
%>        
      <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
<%

set AP_Rs = Conn.execute("select Audited,GroupLevel,BeginDate,EndDate,LeftCount from FS_AP_UserList where UserNumber='"&session("FS_UserNumber")&"'")
if not AP_Rs.eof then isValidate = AP_Rs(0) : GroupLevel = AP_Rs(1) : BeginDate = AP_Rs(2) : EndDate = AP_Rs(3)
if isnull(isValidate) then isValidate=0
if isValidate=-1 then 
%>
    <tr  class="hback"> 
      <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job0.style.display==''){document.getElementById('img0').src='../Images/+.gif';job0.style.display='none'}else{job0.style.display='';document.getElementById('img0').src='../Images/-.gif'}"><img id="img0" src="../Images/+.gif">ע�ṫ˾��չ��Ϣ���Ա������Ƹ��</td> 
    </tr>
    <tr  class="hback" id="job0" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img1" src="../Images/-.gif"> <a href="AP_Register.asp">ע�ṫ˾��չ��Ϣ���Ա������Ƹ��</a> 
           </td> 	
    </tr>
<%elseif isValidate=0 then%>	

    <tr  class="hback"> 
      <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job0.style.display==''){document.getElementById('img0').src='../Images/+.gif';job0.style.display='none'}else{job0.style.display='';document.getElementById('img0').src='../Images/-.gif'}"><img id="img0" src="../Images/+.gif">�޸Ĺ�˾��չ��Ϣ���Ա������Ƹ��</td> 
    </tr>
    <tr  class="hback" id="job0" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img1" src="../Images/-.gif"> <a href="AP_Register.asp">�޸Ĺ�˾��չ��Ϣ���Ա������Ƹ��</a> <span class="tx">״̬:�����...</span>
           </td> 	
    </tr>

<%else%>
    <tr  class="hback"> 
      <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job0.style.display==''){document.getElementById('img0').src='../Images/+.gif';job0.style.display='none'}else{job0.style.display='';document.getElementById('img0').src='../Images/-.gif'}"><img id="img0" src="../Images/+.gif">�޸Ĺ�˾��չ��Ϣ���Ա������Ƹ��</td> 
    </tr>
    <tr  class="hback" id="job0" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img1" src="../Images/-.gif"> <a href="AP_Register.asp">�޸Ĺ�˾������Ϣ</a> <span class="tx">״̬:
		  <%
		  select case GroupLevel
		  case 1
		  	response.Write("��ͨ�û�(ʣ��"&LeftCount&"��)")
		  case 2
		  	response.Write("�����û�(��ͨʱ��:"&BeginDate&" ����ʱ��:"&EndDate&")")
		  case else
		  	response.Write("VIP�û�(������)")
		  end select
		  %></span>
           </td> 	
    </tr>

    <tr  class="hback"> 
      <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job1.style.display==''){document.getElementById('img1').src='../Images/+.gif';job1.style.display='none'}else{job1.style.display='';document.getElementById('img1').src='../Images/-.gif'}"><img id="img1" src="../Images/+.gif">��Ƹ��Ϣ����</td> 
    </tr>
    <tr  class="hback" id="job1" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img1" src="../Images/-.gif"> <a href="AP_Job_Public_List.asp">�ѷ�������Ϣ</a> 
            | <a href="AP_Job_Public_AddUpdate.asp?Act=Add">�����Ƹ��Ϣ</a> | <a href="AP_Job_Public_AddUpdate.asp?Act=Search">��ѯ��Ƹ��Ϣ</a></td> 	
    </tr>
    <tr  class="hback"> 
      <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job2.style.display==''){document.getElementById('img2').src='../Images/+.gif';job2.style.display='none'}else{job2.style.display='';document.getElementById('img2').src='../Images/-.gif'}"><img id="img2" src="../Images/+.gif">��ֵ���ѹ���</td> 
    </tr>
    <tr  class="hback" id="job2" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img2" src="../Images/-.gif"> <a href="AP_Payment_List.asp">��ֵ��¼��ѯ</a> 
            | <a href="AP_Consume_List.asp">���Ѽ�¼��ѯ</a></td> 	
    </tr>
    <tr  class="hback"> 
          <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job3.style.display==''){document.getElementById('img2').src='../Images/+.gif';job3.style.display='none'}else{job3.style.display='';document.getElementById('img2').src='../Images/-.gif'}"><img id="img2" src="../Images/+.gif">�˲Ų�ѯ</td> 
    </tr>
    <tr  class="hback" id="job3" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img2" src="../Images/-.gif"> <a href="AP_Person_Search.asp">�˲Ų�ѯ</a> 
            </td> 	
    </tr>
<%end if%>
	</table>
 
<%end if%>
       </td>
    </tr>
    <tr class="back"> 
      <td height="20" colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
<%
Set AP_Rs  = nothing
Set Fs_User = Nothing
%>

<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





