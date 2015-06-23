<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->

<%
Dim Conn,User_Conn,UserSysParaRs
Dim AllowReg,AllowChinese,NeedAudit,Corp_NeedAudit,Login_Filed,OnlyMemberLogin,ID_Rule,IDRule_Array,ID_Elem,ID_Postfix,Needword,UserName_Length,UserName_Length_Max,UserName_Length_Max_Min,UserName_Length_Array,Forbid_UserName,Pwd_Length,Pwd_Length_Max,Pwd_Length_Min,Pwd_Length_Array,Pwd_Contain_Word,ResigterNeedFull,isSendMail,Email_Aduit,Reg_Help
Dim CheckCodeStyle,LoginSytle,ReturnUrl,ErrorPwdTimes,ShowNumberPerPage,UpfileType,UpfileSize,MessageSize,RssFeed,LimitClass,LimitClass_Array,MemberFile,CertDir,LimitReviewChar
Dim MoneyName,MoneyUnit,RegPointmoney,RegPointmoney_Array,LoginPointmoney,LoginPointmoney_Array,PointChange,PointChange_TF,PointChange_Array,isPrompt,isPrompt_TF,isPrompt_Array,LenLoginTime
MF_Default_Conn
MF_User_Conn
MF_Session_TF
Set UserSysParaRs=server.CreateObject(G_FS_RS)
UserSysParaRs.open "select RegisterTF,AllowChineseName,RegisterCheck,isCheckCorp,OnlyMemberLogin,UserNumberRule,LenUserName,LimitUserName,LenPassword,isSendMail,isValidate,RegisterNotice,VerCodeStyle, Login_Style,ReturnUrl,LoginLockNum,MemberList,UpfileType,UpfileSize,MessageSize,RssFeed,limitClass,MemberFile,CertDir,LimitReviewChar,MoneyName,RegPointmoney,LoginPointmoney,PointChange,isPrompt,LenLoginTime From FS_ME_SysPara",User_Conn,1,1
'*****************************************************************
if not UserSysParaRs.eof then
	'baseParam
	AllowReg=UserSysParaRs("RegisterTF")
	AllowChinese=UserSysParaRs("AllowChineseName")
	NeedAudit=UserSysParaRs("RegisterCheck")
	Corp_NeedAudit=UserSysParaRs("isCheckCorp")
	OnlyMemberLogin=UserSysParaRs("OnlyMemberLogin")
	ID_Rule=UserSysParaRs("UserNumberRule")
	UserName_Length=UserSysParaRs("LenUserName")
	Forbid_UserName=UserSysParaRs("LimitUserName")
	Pwd_Length=UserSysParaRs("LenPassword")
	isSendMail=UserSysParaRs("isSendMail")
	Email_Aduit=UserSysParaRs("isValidate")
	Reg_Help=UserSysParaRs("RegisterNotice")
	'OtherParam
	CheckCodeStyle=UserSysParaRs("VerCodeStyle")
	LoginSytle=UserSysParaRs("Login_Style")
	ReturnUrl=UserSysParaRs("ReturnUrl")
	ErrorPwdTimes=UserSysParaRs("LoginLockNum")
	ShowNumberPerPage=UserSysParaRs("MemberList")
	UpfileType=UserSysParaRs("UpfileType")
	UpfileSize=UserSysParaRs("UpfileSize")
	MessageSize=UserSysParaRs("MessageSize")
	RssFeed=UserSysParaRs("RssFeed")
	LimitClass=UserSysParaRs("limitClass")
	MemberFile=UserSysParaRs("MemberFile")
	CertDir=UserSysParaRs("CertDir")
	LimitReviewChar=UserSysParaRs("LimitReviewChar")
	'About Money
	MoneyName=UserSysParaRs("MoneyName")
	RegPointmoney=UserSysParaRs("RegPointmoney")
	LoginPointmoney=UserSysParaRs("LoginPointmoney")
	PointChange=UserSysParaRs("PointChange")
	isPrompt=UserSysParaRs("isPrompt")
	LenLoginTime=UserSysParaRs("LenLoginTime")
End if
'******************************************************************
if len(UserName_Length)>0 then
	UserName_Length_Array=split(UserName_Length,",")
End if
if (Pwd_Length)>0 then
	Pwd_Length_Array=split(Pwd_Length,",")
End if
if len(ID_Rule)>0 then
	IDRule_Array=split(ID_Rule,",")
End if
if Ubound(IDRule_Array)>=1 then
	ID_Elem=IDRule_Array(1)
end if
if Ubound(IDRule_Array)>=2 then
	ID_Postfix=IDRule_Array(2)
end if
if len(LimitClass)>0 then
	LimitClass_Array=split(LimitClass,",")
End if
if len(RegPointmoney)>0 then 
	RegPointmoney_Array=split(RegPointmoney,",")
end if
if len(LoginPointmoney)>0 then 
	LoginPointmoney_Array=split(LoginPointmoney,",")
end if
if len(PointChange)>0 then 
	PointChange_Array=split(PointChange,",")
end if
if ubound(PointChange_Array)>=0 then 
	PointChange_TF=PointChange_Array(0)
end if
if len(isPrompt)>0 then 
	isPrompt_Array=split(isPrompt,",")
end if
if Ubound(isPrompt_Array)>=0 then
	isPrompt_TF=isPrompt_Array(0)
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script>
var selected="Lab_Base";
var flag=true;//是否可以提交表单
function showParamPanel(param)
{
	switch(param)
	{
		case 1:
		document.getElementById("Layer1").style.display="block";
		document.getElementById("Layer2").style.display="none";	
		document.getElementById("Layer3").style.display="none";	
		document.getElementById("Lab_Base").className ="";
		if(selected!="Lab_Base")
		document.getElementById(selected).className ="xingmu";
		selected="Lab_Base";
		break;
		case 2:
		document.getElementById("Layer1").style.display="none";
		document.getElementById("Layer2").style.display="block";
		document.getElementById("Layer3").style.display="none";	
		document.getElementById("Lab_Other").className="";
		if(selected!="Lab_Other")
		document.getElementById(selected).className ="xingmu";
		selected="Lab_Other";
		break;
		case 3:
		document.getElementById("Layer1").style.display="none";
		document.getElementById("Layer2").style.display="none";	
		document.getElementById("Layer3").style.display="block";	
		document.getElementById("Lab_Money").className="";
		if(selected!="Lab_Money")
		document.getElementById(selected).className ="xingmu";
		selected="Lab_Money";
		break;
	}
}
function ReplaceDot(Obj)
{
	var oldValue=Obj.value;
	while(oldValue.indexOf("，")!=-1)
	{
		Obj.value=oldValue.replace("，",",");
		oldValue=Obj.value;
	}
}

function CheckContentLen(Obj,FS_Alert,Len)
{
	if(Obj.value.length>Len)
	{
		document.getElementById("FS_Alert").innerHTML="<font color='F43631'>长度请不要超过"+Len+"</font>";
		flag=false;
	}

}
function isChinese(Obj,FS_Alert)
{ 
	var Number = "0123456789.,abcdefghijklmnopqrstuvwxyz-\/ABCDEFGHIJKLMNOPQRSTUVWXYZ`~!@#$%^&*()_";
	for (i = 0; i < Obj.value.length;i++)
	{   
		var c = Obj.value.charAt(i);
		if (Number.indexOf(c) == -1) 
		{
			document.getElementById(FS_Alert).innerHTML="<font color='F43631'>请不要使用中文字符</font>";
			Obj.focus()
			flag=false;
			return;
		}
	}
	document.getElementById(FS_Alert).innerHTML="";
	flag=true
}
function MySubmit(Obj)
{
	if(flag)
	document.getElementById(Obj).submit();
}
</script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return false;">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr class="hback"> 
      <td width="140" align="right" class="xingmu" colspan="3"><div align="left">会员系统参数设置</div></td>
    </tr>
	<tr class="hback"> 
	<td width="33%"  id="Lab_Base"><div align="center"><a href="#" onClick="showParamPanel(1)">基本参数设置</a></div></td>
	<td width="33%" height="19" class="xingmu" id="Lab_Other"> <div align="center"><a href="#" onClick="showParamPanel(2)">其他参数设置</a></div></td>
	<td width="33%" height="19" class="xingmu" id="Lab_Money"> <div align="center"><a href="#" onClick="showParamPanel(3)">积分金币设置</a></div></td>
	</tr>
    <tr class="hback">
      <td align="right"  colspan="3">
        <div id="Layer1" style="position:relative; z-index:1; left: 0px; top: 0px;"> 
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <form action="ParamAction.asp?Act=BaseParam" method="post" name="BaseParam" id="BaseParam">
            <tr class="hback"> 
              <td align="right">允许注册：</td>
              <td> <input type="radio" name="AllowReg" value="1" <%if AllowReg=1 then Response.Write("checked") end if%>>
                是&nbsp; <input type="radio" name="AllowReg" value="0" <%if AllowReg=0 then Response.Write("checked") end if %>>
                否 </td>
            </tr>
            <tr class="hback"> 
              <td align="right"> 允许中文名称注册：</td>
              <td width="596"> <input type="radio" name="AllowChinese" value="1" <% if AllowChinese=1 then Response.Write("checked") end if%>>
                是&nbsp; <input type="radio" name="AllowChinese" value="0" <% if AllowChinese=0 then Response.Write("checked") end if%>>
                否 </td>
            </tr>
            <tr class="hback"> 
              <td align="right"> 个人注册需要审核：</td>
              <td><input type="radio" name="NeedAudit" value="1" <% if NeedAudit=1 then Response.Write("checked") end if%> >
                是&nbsp; <input type="radio" name="NeedAudit" value="0" <% if NeedAudit=0 then Response.Write("checked")  end if%>>
                否</td>
            </tr>
            <tr class="hback"> 
              <td align="right">企业注册需要审核：</td>
              <td><input type="radio" name="Corp_NeedAudit" value="1" <% if Corp_NeedAudit=1 then Response.Write("checked") end if%>>
                是&nbsp; <input type="radio" name="Corp_NeedAudit" value="0" <% if Corp_NeedAudit=0 then Response.Write("checked") end if%>>
                否</td>
            </tr>
            <tr class="hback"> 
              <td align="right">只允许一个人登录：</td>
              <td><input type="radio" name="OnlyMemberLogin" value="1" <% if OnlyMemberLogin=1 then Response.Write("checked")%>/>
                是&nbsp; <input type="radio" name="OnlyMemberLogin" value="0" <% if OnlyMemberLogin=0 then Response.Write("checked")%>>
                否</td>
            </tr>
            <tr class="hback"> 
              <td align="right">用户编号规则_前缀：</td>
              <td> <input name="ID_Prefix" type="text" id="ID_Prefix2" size="50" value="<% if Ubound(IDRule_Array)>0 then Response.Write(IDRule_Array(0))%>"/> 
              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">用户编号规则_组成元素：</td>
              <td> <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="y" <%if instr(ID_Elem,"y")>0 then Response.Write("checked") end if%>/>
                年 
                <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="m" <%if instr(ID_Elem,"m")>0 then Response.Write("checked") end if%>/>
                月 
                <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="d" <%if instr(ID_Elem,"d")>0 then Response.Write("checked") end if%>/>
                日 
                <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="h" <%if instr(ID_Elem,"h")>0 then Response.Write("checked") end if%>/>
                时 
                <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="i" <%if instr(ID_Elem,"i")>0 then Response.Write("checked") end if%>/>
                分 
                <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="s" <%if instr(ID_Elem,"s")>0 then Response.Write("checked") end if%>/>
                秒 </td>
            </tr>
            <tr class="hback"> 
              <td align="right">用户编号规则_ID后缀：</td>
              <td><p> 
                  <label> 
                  <input type="radio" name="ID_Postfix" value="2" <% if ID_Postfix=2 then Response.Write("Checked") end if%>/>
                  2位随机数</label>
                  <label> 
                  <input type="radio" name="ID_Postfix" value="3" <% if ID_Postfix=3 then Response.Write("Checked") end if%>>
                  3位随机数</label>
                  <label> 
                  <input type="radio" name="ID_Postfix" value="4" <% if ID_Postfix=4 then Response.Write("Checked") end if%>>
                  4位随机数</label>
                  <label> 
                  <input type="radio" name="ID_Postfix" value="5" <% if ID_Postfix=5 then Response.Write("Checked") end if%>>
                  5位随机数</label>
                  <input name="needWord" type="checkbox" id="needWord2" value="w" <% if instr(ID_Rule,"w")>0 then Response.Write("checked")%> />
                  字母数字组合</p></td>
            </tr>
            <tr class="hback"> 
              <td align="right">用户编号规则_分割符：</td>
              <td><input name="ID_Devide" type="text" id="ID_Devide2" size="50" value="<% if Ubound(IDRule_Array)>=3 then Response.Write(IDRule_Array(3)) end if%>"/></td>
            </tr>
            <tr class="hback"> 
              <td align="right">用户名长度限制：</td>
              <td> 最小长度： 
                <input name="UserName_Length_Min" type="text" id="UserName_Length_Min" size="14" value="<%if ubound(UserName_Length_Array)>=0 then Response.Write(UserName_Length_Array(0)) end if%>" onBlur="isNumber(this,'UserName_Length_alert','长度应该为正整数',true)"/>
                最大长度： 
                <input name="UserName_Length_Max" type="text" id="UserName_Length_Max" size="14" value="<%if ubound(UserName_Length_Array)>=1 then Response.Write(UserName_Length_Array(1)) end if%>" onBlur="isNumber(this,'UserName_Length_alert','长度应该为正整数',true)"/> 
                <span id="UserName_Length_alert">&nbsp;</span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">禁止注册的用户名：</td>
              <td><textarea name="Forbid_UserName" cols="50" id="textarea2" onKeyUp="ReplaceDot(this)"><%=Forbid_UserName%></textarea></td>
            </tr>
            <tr class="hback"> 
              <td align="right">密码长度限制：</td>
              <td> 最小长度： 
                <input name="Pwd_Length_Max" type="text" id="Pwd_Length_Max" size="14"  value="<%if ubound(Pwd_Length_Array)>=0 then Response.Write(Pwd_Length_Array(0)) end if%>" onBlur="isNumber(this,'Pwd_Length_alert','长度应该为正整数',true)"/>
                最大长度： 
                <input name="Pwd_Length_Min" type="text" id="Pwd_Length_Min" size="14"  value="<%if ubound(Pwd_Length_Array)>=1 then Response.Write(Pwd_Length_Array(1)) end if%>" onBlur="isNumber(this,'Pwd_Length_alert','长度应该为正整数',true)"/> 
                <span id="Pwd_Length_alert">&nbsp;</span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">发送电子邮件信息：</td>
              <td><input type="radio" name="isSendMail" value="1" <%if isSendMail=1 then Response.Write("checked") end if%>>
                是&nbsp; <input type="radio" name="isSendMail" value="0" <%if isSendMail=0 then Response.Write("checked") end if%>/>
                否 </td>
            </tr>
            <tr class="hback"> 
              <td align="right">电子邮件验证：</td>
              <td><input type="radio" name="Email_Aduit" value="1" <%if Email_Aduit=1 then Response.Write("checked") end if%>/>
                是&nbsp; <input type="radio" name="Email_Aduit" value="0" <%if Email_Aduit=0 then Response.Write("checked") end if%>>
                否 </td>
            </tr>
            <tr class="hback"> 
              <td align="right">会员注册须知：<br>
                （支持html语法）<br> <span id="Reg_Help_alert">&nbsp;</span></td>
              <td><textarea name="Reg_Help" cols="80" rows="10" id="Reg_Help" onBlur="CheckContentLen(this,'Reg_Help_alert',2000)"><%=Reg_Help%></textarea></td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td><input type="Button" name="BaseSubmit"  onClick="MySubmit('BaseParam')" value=" 保存 " /> 
                <input type="reset" name="reset" value=" 重置 " /></td>
            </tr>
          </form>
        </table>
      </div>	  
        
      <div id="Layer2" style="position:relative;z-index:1; left: 0px; top: 0px; display:none"> 
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <form action="ParamAction.asp?Act=OtherParam" method="post" name="OtherParam" id="OtherParam">
            <tr class="hback"> 
              <td align="right">登录风格： </td>
              <td width="596"><select name="LoginStyle" id="LoginStyle">
                  <option value="1" <%if LoginSytle=1 then Response.Write("selected") end if%>>默认风格</option>
                  <option value="2" <%if LoginSytle=2 then Response.Write("selected") end if%>>银色风格</option>
                  <option value="3" <%if LoginSytle=3 then Response.Write("selected") end if%>>蓝色海洋</option>
                  <option value="4" <%if LoginSytle=4 then Response.Write("selected") end if%>>浪漫咖啡</option>
                  <option value="5" <%if LoginSytle=5 then Response.Write("selected") end if%>>青青河草</option>
                </select></td>
            </tr>
            <tr class="hback"> 
              <td align="right">登录成功返回页面：</td>
              <td><p> 
                  <label> </label>
                  <label> 
                  <input type="radio" name="ReturnUrl" value="0" <%if ReturnUrl=0 then Response.Write("checked") end if%>>
                  会员中心 </label>
                  <label> 
                  <input type="radio" name="ReturnUrl" value="1" <%if ReturnUrl=1 then Response.Write("checked") end if%> />
                  首页 </label>
                  <label> 
                  <input type="radio" name="ReturnUrl" value="2" <%if ReturnUrl=2 then Response.Write("checked") end if%>>
                  之前页</label>
                  <br>
                </p></td>
            </tr>
            <tr class="hback"> 
              <td align="right">防登录猜测：</td>
              <td>密码错误 
                <input name="ErrorPwdTimes" type="text" id="EErrorPwdTimes" onBlur="isNumber(this,'ErrorPwdTimes_alert','次数应该为正整数',true)" value="<%=ErrorPwdTimes%>" size="10" />
                次后，锁定该用户功能 （0为不限制） <span id="ErrorPwdTimes_alert">&nbsp;</span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">会员列表每页显示数：</td>
              <td><input name="ShowNumberPerPage" type="text" id="PShowNumberPerPage" size="50" value="<%=ShowNumberPerPage%>" onBlur="isNumber(this,'ShowNumberPerPage_alert','页数应该为正整数',true)"/> 
                <span id="ShowNumberPerPage_alert">&nbsp;</span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">上传文件类型：</td>
              <td><input name="UpfileType" type="text" id="UpfileType" size="50" onKeyUp="ReplaceDot(this)" value="<%=UpfileType%>"/>
                （用 ，隔开）</td>
            </tr>
            <tr class="hback"> 
              <td align="right">上传大小限制：</td>
              <td><input name="UpfileSize" type="text" id="UpfileSize" size="50" onBlur="isNumber(this,'UpfileSize_alert','文件大小应该为正整数',true)" value="<%=UpfileSize%>" />
                k <span id="UpfileSize_alert">&nbsp;</span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">站内短信容量:</td>
              <td><input name="MessageSize" type="text" id="MessageSize" size="50" onBlur="isNumber(this,'MessageSize_alert','短信容量应该为正整数',ture)" value="<%=MessageSize%>"/>
                byte <span id="MessageSize_alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">开通RSS订阅：</td>
              <td><label> 
                <input type="radio" name="RssFeed" value="1" <%if RssFeed=1 then Response.Write("checked") end if%>>
                是 </label> <label> 
                <input type="radio" name="RssFeed" value="0" <%if RssFeed=0 then Response.Write("checked") end if%>>
                否</label></td>
            </tr>
            <tr class="hback"> 
              <td align="right">信息分类控制：</td>
              <td>分类级数量上限：
                <input name="LimitClass" type="text" id="LimitClass" size="14"  value="<%if ubound(LimitClass_Array)>=0 then Response.Write(LimitClass_Array(0)) end if%>"> 
                &nbsp;&nbsp;分类数量上限 
                <input name="LimitClass2" type="text" id="LimitClass2" size="14"  value="<%if ubound(LimitClass_Array)>=1 then Response.Write(LimitClass_Array(1)) end if%>"></td>
            </tr>
            <tr class="hback"> 
              <td align="right">用户目录：</td>
              <td><input name="MemberFile" type="text" id="MemberFile"  size="50" value="<%=MemberFile%>"/>
                （默认：userfiles）</td>
            </tr>
            <tr class="hback"> 
              <td align="right">认证文件目录：</td>
              <td><input name="CertDir" type="text" id="CertDir"  value="<%=CertDir%>" size="50"/>
                （默认：certfiles）</td>
            </tr>
            <tr class="hback"> 
              <td align="right">过滤评论关键字：</td>
              <td><textarea name="LimitReviewChar" cols="50" onKeyUp="ReplaceDot(this)" id="LimitReviewChar"><%=LimitReviewChar%></textarea></td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td><input type="Button" name="Submit" onClick="MySubmit('OtherParam')" value=" 保存 " /> 
                <input type="reset" name="Submit2" value=" 重置 " /></td>
            </tr>
          </form>
        </table>
      </div>
	  <div id="Layer3" style="position:relative;z-index:1; left: 0px; top: 0px; display:none"> 
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <form action="ParamAction.asp?Act=MoneyParam" method="post" name="MoneyParam" id="MoneyParam">
            <tr class="hback"> 
              <td align="right">金币名称：</td>
              <td width="596"> <input name="MoneyName" type="text" id="MoneyName" size="50" value="<%=MoneyName%>" /> 
              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">注册获取积分和金币：</td>
              <td>积分： 
                <input name="Reg_Point" type="text" id="Reg_Point" size="15"  value="<%if ubound(RegPointmoney_Array)>=0 then Response.Write(RegPointmoney_Array(0)) end if%>" onBlur="isNumber(this,'Reg_PointMoney_Alert','积分应该为正整数',true)"/> 
                &nbsp; 金币： 
                <input name="Reg_Money" type="text" id="Reg_Money" size="15"  value="<%if ubound(RegPointmoney_Array)>=1 then Response.Write(RegPointmoney_Array(1)) end if%>" onBlur="isNumber(this,'Reg_PointMoney_Alert','金钱应该为正整数,true')"/> 
                <span id="Reg_PointMoney_Alert">&nbsp;</span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">登录获取积分和金币：</td>
              <td>积分： 
                <input name="Login_Point" type="text" id="Login_Point" size="15"  value="<%if ubound(LoginPointmoney_Array)>=0 then Response.Write(LoginPointmoney_Array(0)) end if%>" onBlur="isNumber(this,'Login_PointMoney_Alert','积分应该为正整数',true)"/> 
                &nbsp;&nbsp;金币： 
                <input name="Login_money" type="text" id="Login_money" size="15"  value="<%if ubound(LoginPointmoney_Array)>=1 then Response.Write(LoginPointmoney_Array(1)) end if%>"onBlur="isNumber(this,'Login_PointMoney_Alert','金钱应该为正整数',true)" /> 
                <span id="Login_PointMoney_Alert">&nbsp;</span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">积分金币兑换比例：</td>
              <td> <label> 
                <input type="radio" name="PointChange" value="0" <%if PointChange_TF=0 then Response.Write("checked") end if%>/>
                不启用</label> <label> 
                <input type="radio" name="PointChange" value="1" <%if PointChange_TF=1 then Response.Write("checked") end if%>>
                积分换金币</label> <label> 
                <input type="radio" name="PointChange" value="2" <%if PointChange_TF=2 then Response.Write("checked") end if%>>
                金币换积分</label> <label> </label> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td>1金币 = 
                <input name="Money_to_Point" type="text" id="Money_to_Point" size="15" value="<%if Ubound(PointChange_Array)>=1 then Response.Write(PointChange_Array(1)) end if%>" onBlur="isNumber(this,'PointChange_Alert','积分应该为正数,true')"/>
                积分 | 1积分 = 
                <input name="Point_to_Money" type="text" id="Point_to_Money" size="15" value="<%if Ubound(PointChange_Array)>=2 then Response.Write(PointChange_Array(2)) end if%>" onBlur="isNumber(this,'PointChange_Alert','金钱应该为正数',true)"/>
                金币 <span id="PointChange_Alert"></span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">金币提示条件：</td>
              <td> <label> 
                <input type="radio" name="isPrompt" value="1" <%if isPrompt_TF=1 then Response.Write("checked") end if%> />
                开启</label> <label> 
                <input type="radio" name="isPrompt" value="0" <%if isPrompt_TF=0 then Response.Write("checked") end if%> />
                关闭</label>
                | 金币小于 
                <input name="PromptCondition" type="text" id="PromptCondition" size="15" value="<%if ubound(isPrompt_Array)>=1 then Response.Write(isPrompt_Array(1)) end if%>" onBlur="isNumber(this,'Prompt_Alert','金钱应该为正数'，false)"/>
                时提示 <span id="Prompt_Alert"></span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">防作弊设置：</td>
              <td>登录 
                <input name="LenLoginTime" type="text" id="LenLoginTime" size="10" value="<%=LenLoginTime%>" onBlur="isNumber(this,'LenLoginTime_Alert','时间应该为正整数',true)"/>
                分后，再次登录积分增加 <span id="LenLoginTime_Alert"></span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td><input type="Button" name="MoneySubmit" onClick="MySubmit('MoneyParam')"value=" 保存 " /> 
                <input type="reset" name="Submit2" value=" 重置 " /></td>
            </tr>
          </form>
        </table>
      </div>	  </td>
    </tr>
</table>
</body>
<%
	UserSysParaRs.close
	Set UserSysParaRs=nothing
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
%>
</html>






