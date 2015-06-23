<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-招聘管理</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
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
            
          <td class="hback"><strong>位置：</strong><a href="../../">网站首页</a> &gt;&gt; 
            <a href="../main.asp">会员首页</a> &gt;&gt; 招聘</td>
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
	response.Write("你不是公司，没有权利管理这个页面。若有需要请重新注册。")	
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
      <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job0.style.display==''){document.getElementById('img0').src='../Images/+.gif';job0.style.display='none'}else{job0.style.display='';document.getElementById('img0').src='../Images/-.gif'}"><img id="img0" src="../Images/+.gif">注册公司扩展信息（以便管理招聘）</td> 
    </tr>
    <tr  class="hback" id="job0" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img1" src="../Images/-.gif"> <a href="AP_Register.asp">注册公司扩展信息（以便管理招聘）</a> 
           </td> 	
    </tr>
<%elseif isValidate=0 then%>	

    <tr  class="hback"> 
      <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job0.style.display==''){document.getElementById('img0').src='../Images/+.gif';job0.style.display='none'}else{job0.style.display='';document.getElementById('img0').src='../Images/-.gif'}"><img id="img0" src="../Images/+.gif">修改公司扩展信息（以便管理招聘）</td> 
    </tr>
    <tr  class="hback" id="job0" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img1" src="../Images/-.gif"> <a href="AP_Register.asp">修改公司扩展信息（以便管理招聘）</a> <span class="tx">状态:审核中...</span>
           </td> 	
    </tr>

<%else%>
    <tr  class="hback"> 
      <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job0.style.display==''){document.getElementById('img0').src='../Images/+.gif';job0.style.display='none'}else{job0.style.display='';document.getElementById('img0').src='../Images/-.gif'}"><img id="img0" src="../Images/+.gif">修改公司扩展信息（以便管理招聘）</td> 
    </tr>
    <tr  class="hback" id="job0" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img1" src="../Images/-.gif"> <a href="AP_Register.asp">修改公司部分信息</a> <span class="tx">状态:
		  <%
		  select case GroupLevel
		  case 1
		  	response.Write("普通用户(剩余"&LeftCount&"点)")
		  case 2
		  	response.Write("包月用户(开通时间:"&BeginDate&" 到期时间:"&EndDate&")")
		  case else
		  	response.Write("VIP用户(无限制)")
		  end select
		  %></span>
           </td> 	
    </tr>

    <tr  class="hback"> 
      <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job1.style.display==''){document.getElementById('img1').src='../Images/+.gif';job1.style.display='none'}else{job1.style.display='';document.getElementById('img1').src='../Images/-.gif'}"><img id="img1" src="../Images/+.gif">招聘信息管理</td> 
    </tr>
    <tr  class="hback" id="job1" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img1" src="../Images/-.gif"> <a href="AP_Job_Public_List.asp">已发布的信息</a> 
            | <a href="AP_Job_Public_AddUpdate.asp?Act=Add">添加招聘信息</a> | <a href="AP_Job_Public_AddUpdate.asp?Act=Search">查询招聘信息</a></td> 	
    </tr>
    <tr  class="hback"> 
      <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job2.style.display==''){document.getElementById('img2').src='../Images/+.gif';job2.style.display='none'}else{job2.style.display='';document.getElementById('img2').src='../Images/-.gif'}"><img id="img2" src="../Images/+.gif">充值消费管理</td> 
    </tr>
    <tr  class="hback" id="job2" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img2" src="../Images/-.gif"> <a href="AP_Payment_List.asp">充值记录查询</a> 
            | <a href="AP_Consume_List.asp">消费记录查询</a></td> 	
    </tr>
    <tr  class="hback"> 
          <td align="left" height="25" valign="top" style="cursor:hand" class="xingmu" onClick="if(job3.style.display==''){document.getElementById('img2').src='../Images/+.gif';job3.style.display='none'}else{job3.style.display='';document.getElementById('img2').src='../Images/-.gif'}"><img id="img2" src="../Images/+.gif">人才查询</td> 
    </tr>
    <tr  class="hback" id="job3" style="display:"> 
	      <td height="23"> &nbsp;&nbsp;<img id="img2" src="../Images/-.gif"> <a href="AP_Person_Search.asp">人才查询</a> 
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

<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





