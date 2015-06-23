<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
Dim Conn,User_Conn,UserSysParaRs
Dim AllowReg,AllowChinese,NeedAudit,Corp_NeedAudit,Login_Filed,OnlyMemberLogin,ID_Rule,IDRule_Array,ID_Elem,ID_Postfix,Needword,UserName_Length,UserName_Length_Max,UserName_Length_Max_Min,UserName_Length_Array,Forbid_UserName,Pwd_Length,Pwd_Length_Max,Pwd_Length_Min,Pwd_Length_Array,Pwd_Contain_Word,ResigterNeedFull,isSendMail,Email_Aduit,Reg_Help
Dim CheckCodeStyle,LoginSytle,ReturnUrl,ErrorPwdTimes,ShowNumberPerPage,UpfileType,UpfileSize,MessageSize,RssFeed,LimitClass,LimitClass_Array,CertDir,LimitReviewChar,isPassCard,isYellowCheck
Dim DefaultGroupID,MoneyName,MoneyUnit,RegPointmoney,RegPointmoney_Array,LoginPointmoney,LoginPointmoney_Array,PointChange,PointChange_TF,PointChange_Array,isPrompt,isPrompt_TF,isPrompt_Array,LenLoginTime,LenLoginTimeArray,ReviewTF,contrPoint,contrMoney,contrAuditPoint,contrAuditMoney,UserSystemName
MF_Default_Conn
MF_User_Conn
MF_Session_TF
Set UserSysParaRs=server.CreateObject(G_FS_RS)
UserSysParaRs.open "select top 1 RegisterTF,AllowChineseName,RegisterCheck,isCheckCorp,OnlyMemberLogin,UserNumberRule,LenUserName,LimitUserName,LenPassword,isSendMail,isValidate,RegisterNotice,VerCodeStyle, Login_Style,ReturnUrl,LoginLockNum,MemberList,UpfileType,UpfileSize,MessageSize,RssFeed,limitClass,CertDir,LimitReviewChar,MoneyName,RegPointmoney,LoginPointmoney,PointChange,isPrompt,LenLoginTime,DefaultGroupID,isYellowCheck,isPassCard,ReviewTF,contrPoint,contrMoney,contrAuditPoint,contrAuditMoney,UserSystemName From FS_ME_SysPara",User_Conn,1,3
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
	CertDir=UserSysParaRs("CertDir")
	LimitReviewChar=UserSysParaRs("LimitReviewChar")
	ReviewTF=UserSysParaRs("ReviewTF")
	isYellowCheck=UserSysParaRs("isYellowCheck")
	isPassCard=UserSysParaRs("isPassCard")
	DefaultGroupID=UserSysParaRs("DefaultGroupID")
	'About Money
	MoneyName=UserSysParaRs("MoneyName")
	RegPointmoney=UserSysParaRs("RegPointmoney")
	LoginPointmoney=UserSysParaRs("LoginPointmoney")
	PointChange=UserSysParaRs("PointChange")
	isPrompt=UserSysParaRs("isPrompt")
	LenLoginTime=UserSysParaRs("LenLoginTime")
	contrPoint=UserSysParaRs("contrPoint")
	contrMoney=UserSysParaRs("contrMoney")
	contrAuditPoint=UserSysParaRs("contrAuditPoint")
	contrAuditMoney=UserSysParaRs("contrAuditMoney")
	UserSystemName=UserSysParaRs("UserSystemName")
End if
'******************************************************************
if len(UserName_Length)>0 then
	UserName_Length_Array=split(UserName_Length,",")
End if
if len(Pwd_Length)>0 then
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
if len(LenLoginTime)>0 then
	LenLoginTimeArray=split(LenLoginTime,",")
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
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr class="hback"> 
      <td width="140" align="right" class="xingmu" colspan="3"><div align="left">会员系统参数设置</div></td>
    </tr>
	<tr class="hback"> 
	<td width="33%"  id="Lab_Base"><div align="center"><strong><a href="#" onClick="showParamPanel(1)">基本参数设置</a></strong></div></td>
	<td width="33%" height="19" class="hback_1" id="Lab_Other"> <div align="center"><strong><a href="#" onClick="showParamPanel(2)">其他参数设置</a></strong></div></td>
	<td width="33%" height="19" class="hback_1" id="Lab_Money"> <div align="center"><strong><a href="#" onClick="showParamPanel(3)">积分金币设置</a></strong></div></td>
	</tr>
    <tr class="hback">
      <td align="right"  colspan="3">
        <div id="Layer1" style="position:relative; z-index:1; left: 0px; top: 0px;"> 
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <form action="ParamAction.asp?Act=BaseParam" method="post" name="BaseParam" id="BaseParam">
            <tr class="hback"> 
              <td width="210" align="right"><div align="right">会员注册默认会员组：</div></td>
              <td><label> 
                <select name="DefaultGroupID" id="DefaultGroupID">
                  <%
				dim rsGroup
				set rsGroup = User_Conn.execute("select GroupID,GroupName from FS_ME_Group where GroupType=1 Order by GroupID asc")
				do while not rsGroup.eof
					if clng(DefaultGroupID) = rsGroup("GroupID") Then
						Response.Write ("<option value="""&rsGroup("GroupID")&""" selected>"&rsGroup("GroupName")&"</option>")
					else
						Response.Write ("<option value="""&rsGroup("GroupID")&""">"&rsGroup("GroupName")&"</option>")
					end if
					rsGroup.Movenext
				loop
				rsGroup.close:set rsGroup=nothing
				%>
                </select>
                </label></td>
            </tr>
            <tr class="hback"> 
              <td align="right">会员系统名称：</td>
              <td><input name="UserSystemName" type="text" id="UserSystemName" size="30" maxlength="30" value="<%=UserSystemName%>"></td>
            </tr>
            <tr class="hback">
              <td align="right">允许注册：</td>
              <td><input type="radio" name="AllowReg" value="1" <%if AllowReg=1 then Response.Write("checked") end if%>>
是&nbsp;
<input type="radio" name="AllowReg" value="0" <%if AllowReg=0 then Response.Write("checked") end if %>>
否 </td>
            </tr>
            <tr class="hback"> 
              <td align="right"> 允许中文名称注册：</td>
              <td width="743"> <input type="radio" name="AllowChinese" value="1" <% if AllowChinese=1 then Response.Write("checked") end if%>>
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
              <td> <input name="ID_Prefix" type="text" id="ID_Prefix2" size="50" value="<% if Ubound(IDRule_Array)>0 then Response.Write(IDRule_Array(0))%>"/>              </td>
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
              <td><input name="ID_Devide" type="text" id="ID_Devide2" size="50" value="<% if Ubound(IDRule_Array)>=3 then Response.Write(IDRule_Array(3))%>"/></td>
            </tr>
            <tr class="hback"> 
              <td align="right">用户名长度限制：</td>
              <td> 最小长度： 
                <input name="UserName_Length_Min" type="text" id="UserName_Length_Min" size="14" value="<% = UserName_Length_Array(0)%>"/>
                最大长度： 
                <input name="UserName_Length_Max" type="text" id="UserName_Length_Max" size="14" value="<% = UserName_Length_Array(1)%>" /> 
                <span id="UserName_Length_alert1"></span><span id="UserName_Length_alert2"></span>              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">禁止注册的用户名：</td>
              <td><textarea name="Forbid_UserName" cols="50" id="textarea2" onKeyUp="ReplaceDot('Forbid_UserName')"><%=Forbid_UserName%></textarea></td>
            </tr>
            <tr class="hback"> 
              <td align="right">密码长度限制：</td>
              <td> 最小长度： 
                <input name="Pwd_Length_Min" type="text" id="Pwd_Length_Min" size="14"  value="<% = Pwd_Length_Array(0)%>"/>
                最大长度： 
                <input name="Pwd_Length_Max" type="text" id="Pwd_Length_Max" size="14"  value="<% = Pwd_Length_Array(1)%>"/> 
                <span id="Pwd_Length_alert1"></span><span id="Pwd_Length_alert2"></span ></td>
            </tr>
            <tr class="hback"  style="display:none"> 
              <td align="right">发送电子邮件信息：</td>
              <td><input type="radio" name="isSendMail" value="1">
                是&nbsp; <input type="radio" name="isSendMail" value="0" checked/>
                否 </td>
            </tr>
            <tr class="hback" style="display:none"> 
              <td align="right">电子邮件验证：</td>
              <td><input type="radio" name="Email_Aduit" value="1"/>
                是&nbsp; <input type="radio" name="Email_Aduit" value="0" checked>
                否 </td>
            </tr>
            <tr class="hback"> 
              <td align="right">会员注册须知：<br>
                （支持html语法）<br> <span id="Reg_Alert"></span></td>
              <td><textarea name="Reg_Help" cols="80" rows="10" id="Reg_Help"><%=Reg_Help%></textarea></td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td><input type="Button" name="BaseSubmitButton"  onClick="BaseSubmit()" value=" 保存 " /> 
                <input type="reset" name="reset" value=" 重置 " /></td>
            </tr>
          </form>
        </table>
      </div>	  
        
      <div id="Layer2" style="position:relative;z-index:1; left: 0px; top: 0px; display:none"> 
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <form action="ParamAction.asp?Act=OtherParam" method="post" name="OtherParam" id="OtherParam">
            <tr class="hback"> 
              <td align="right" width="190">登录风格： </td>
              <td><select name="LoginStyle" id="LoginStyle">
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
            <tr class="hback" style="display:none"> 
              <td align="right">黄页是否审核：</td>
              <td> <input type="radio" name="isYellowCheck" value="1" <%if isYellowCheck=1 then Response.Write("checked")%>>
                是 
                <input type="radio" name="isYellowCheck" value="0" <%if isYellowCheck=0 then Response.Write("checked")%>>
                否</td>
            </tr>
            <tr class="hback"> 
              <td align="right">名片收藏需要允许：</td>
              <td><input type="radio" name="isPassCard" value="1" <%if isPassCard=1 then Response.Write("checked")%>>
                是 
                <input type="radio" name="isPassCard" value="0" <%if isPassCard=0 then Response.Write("checked")%>>
                否</td>
            </tr>
            <tr class="hback"> 
              <td align="right">防登录猜测：</td>
              <td>密码错误 
                <input name="ErrorPwdTimes" type="text" id="EErrorPwdTimes" value="<%=ErrorPwdTimes%>" size="10" />
                次后，锁定该用户功能 （0为不限制） <span id="ErrorPwdTimes_alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">会员列表每页显示数：</td>
              <td><input name="ShowNumberPerPage" type="text" id="PShowNumberPerPage" size="50" value="<%=ShowNumberPerPage%>" /> 
                <span id="ShowNumberPerPage_alert"></span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">上传文件类型：</td>
              <td><input name="UpfileType" type="text" id="UpfileType" size="50" onKeyUp="ReplaceDot('UpfileType')" value="<%=UpfileType%>"/>
                （用 ，隔开）</td>
            </tr>
            <tr class="hback"> 
              <td align="right">上传大小限制：</td>
              <td><input name="UpfileSize" type="text" id="UpfileSize" size="50"  value="<%=UpfileSize%>" />
                k <span id="UpfileSize_alert">&nbsp;</span>单个上传文件大小，对所有用户 </td>
            </tr>
            <tr class="hback"> 
              <td align="right">站内短信容量:</td>
              <td><input name="MessageSize" type="text" id="MessageSize" size="50" value="<%=MessageSize%>"/>
                kbyte <span id="MessageSize_alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">开通RSS订阅：</td>
              <td><label> 
                <input type="radio" name="RssFeed" value="1" <%if RssFeed=1 then Response.Write("checked") end if%>>
                是 </label> <label> 
                <input type="radio" name="RssFeed" value="0" <%if RssFeed=0 then Response.Write("checked") end if%>>
                否</label></td>
            </tr>
            <tr class="hback" style="display:none"> 
              <td align="right">信息分类(专区,文集)：</td>
              <td>最多分类： 
                <input name="LimitClass" type="text" id="LimitClass" size="14"  value="<%if ubound(LimitClass_Array)>=0 then Response.Write(LimitClass_Array(0)) end if%>"> 
                &nbsp;&nbsp;分类的数量 
                <input name="LimitClass2" type="text" id="LimitClass2" size="14"  value="<%if ubound(LimitClass_Array)>=1 then Response.Write(LimitClass_Array(1)) end if%>"> 
                <span id="LimitClass_Alert1"></span>&nbsp;<span id="LimitClass_Alert2"></span>如果设置了会员默认会员组权限，此项设置不起作用</td>
            </tr>
            <tr class="hback" style="display:none;"> 
              <td align="right">用户目录：</td>
              <td><input name="MemberFile" type="text" id="MemberFile"  size="50" value=""/>
                （默认：userfiles）</td>
            </tr>
            <tr class="hback" style="display:none;"> 
              <td align="right">认证文件目录：</td>
              <td><input name="CertDir" type="text" id="CertDir"  value="<%=CertDir%>" size="50"/>
                （默认：certfiles）</td>
            </tr>
            <tr class="hback"> 
              <td align="right">过滤评论关键字：</td>
              <td><textarea name="LimitReviewChar" cols="50" onKeyUp="ReplaceDot('LimitReviewChar')" id="LimitReviewChar"><%=LimitReviewChar%></textarea></td>
            </tr>
            <tr class="hback"> 
              <td align="right">评论是否需要审核：</td>
              <td><input name="ReviewTF" type="radio" value="1" <%if ReviewTF=1 then response.Write"checked"%>/>
                是 
                <input name="ReviewTF" type="radio" value="0" <%if ReviewTF=0 then response.Write"checked"%> />
                否</td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td><input type="Button" name="OtherSubmitButtont" onClick="OtherSubmit()" value=" 保存 " /> 
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
                <span id="moneyName_alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">注册获取积分和金币：</td>
              <td>积分： 
                <input name="Reg_Point" type="text" id="Reg_Point" size="15"  value="<%if ubound(RegPointmoney_Array)>=0 then Response.Write(RegPointmoney_Array(0)) end if%>" /> 
                &nbsp; 金币： 
                <input name="Reg_Money" type="text" id="Reg_Money" size="15"  value="<%if ubound(RegPointmoney_Array)>=1 then Response.Write(RegPointmoney_Array(1)) end if%>" /> 
                <span id="Reg_Point_Alert"></span>&nbsp; <span id="Reg_Money_Alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">积分金币兑换比例：</td>
              <td> 
			  <input type="radio" name="PointChange" value="0" <%if PointChange_TF=0 then Response.Write("checked") end if%> onClick="DisAbledPorM('dd','dd');"/>
                不启用 
                <input type="radio" name="PointChange" value="1" <%if PointChange_TF=1 then Response.Write("checked") end if%> onClick="DisAbledPorM('dd','d');">
                积分换金币 
                <input type="radio" name="PointChange" value="2" <%if PointChange_TF=2 then Response.Write("checked") end if%> onClick="DisAbledPorM('d','dd');">
                金币换积分 
                <input type="radio" name="PointChange" value="3" <%if PointChange_TF=3 then Response.Write("checked") end if%> onClick="DisAbledPorM('d','d');">
                互换</td>
            </tr>
            <%
				Dim DisAbledPoint,DisAbledMoney,PointAutoV,MoneyAutoV
				IF PointChange_TF=0 Then
					DisAbledPoint = " disabled"
					DisAbledMoney = " disabled"
					PointAutoV = "300"
					MoneyAutoV = "0.0001"
				ElseIf PointChange_TF=1 Then
					DisAbledPoint = " disabled"
					DisAbledMoney = ""
					PointAutoV = "300"
					if Ubound(PointChange_Array) >= 2 then 
						MoneyAutoV = PointChange_Array(2) 
					Else 
						MoneyAutoV =  "0.0001" 
					End if
				ElseIF PointChange_TF=2 Then
					DisAbledPoint = ""		
					DisAbledMoney = " disabled"
					if Ubound(PointChange_Array) >= 1 then 
						PointAutoV = PointChange_Array(1) 
					Else 
						PointAutoV =  "300" 
					End if
					MoneyAutoV = "0.0001"
				ElseIF PointChange_TF=3 Then
					DisAbledPoint = ""
					DisAbledMoney = ""
					if Ubound(PointChange_Array) >= 1 then 
						PointAutoV = PointChange_Array(1) 
					Else 
						PointAutoV =  "300" 
					End if
					if Ubound(PointChange_Array) >= 2 then 
						MoneyAutoV = PointChange_Array(2) 
					Else 
						MoneyAutoV =  "0.0001" 
					End if
				End If		
			%>
			<tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td>1金币 = 
                <input name="Money_to_Point" type="text"<% = DisAbledPoint %> id="Money_to_Point" size="15" value="<% = PointAutoV %>" />
                积分 | 1积分 = 
                <input name="Point_to_Money" type="text"<% = DisAbledMoney %> id="Point_to_Money" size="15" value="<% = MoneyAutoV %>" />
                金币 <span id="Money_Point_Alert"></span>&nbsp; <span id="Point_Money_Alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">投稿奖励：</td>
              <td>积分： 
                <input name="txt_contrPoint" type="text" id="txt_contrPoint" value="<%=contrPoint%>">
                |&nbsp;金币： 
                <input name="txt_contrMoney" type="text" id="txt_contrMoney" value="<%=contrMoney%>"> 
                <span id="contrPoint_Alert"></span>&nbsp; <span id="contrMoney_Alert"></span> 
              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">投稿审核奖励：</td>
              <td>积分： 
                <input name="txt_contrAuditPoint" type="text" id="txt_contrAuditPoint" value="<%=contrAuditPoint%>">
                |&nbsp;金币： 
                <input name="txt_contrAuditMoney" type="text" id="txt_contrAuditMoney" value="<%=contrAuditMoney%>"> 
                <span id="contrAuditPoint_Alert"></span>&nbsp; <span id="contrAuditMoney_Alert"> 
              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">金币提示条件：</td>
              <td> <label> 
                <input type="radio" name="isPrompt" value="1" <%if isPrompt_TF=1 then Response.Write("checked") end if%> />
                开启</label> <label> 
                <input type="radio" name="isPrompt" value="0" <%if isPrompt_TF=0 then Response.Write("checked") end if%> />
                关闭</label>
                | 金币小于 
                <input name="PromptCondition" type="text" id="PromptCondition" size="15" value="<%if ubound(isPrompt_Array)>=1 then Response.Write(isPrompt_Array(1)) end if%>"/>
                时提示 <span id="Prompt_Alert"></span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">防作弊设置：</td>
              <td>登录 
                <input name="LenLoginTime" type="text" id="LenLoginTime" size="10" value="<%if Ubound(LenLoginTimeArray)>=0 then Response.Write(LenLoginTimeArray(0))%>" />
                分后，再次登录积分增加 
                <input name="Login_Point" type="text" id="Login_Point" size="15"  value="<%if ubound(LoginPointmoney_Array)>=0 then Response.Write(LoginPointmoney_Array(0)) end if%>" /> 
                <span id="LenLoginTime_Alert"></span>&nbsp;<span id="Login_Point_Alert"></span> 
              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td>登录 
                <input name="LenLoginTime2" type="text" id="LenLoginTime2" size="10" value="<%if Ubound(LenLoginTimeArray)>=1 then Response.Write(LenLoginTimeArray(1))%>" />
                天后，再次登录金币增加 
                <input name="Login_money" type="text" id="Login_money" size="15"  value="<%if ubound(LoginPointmoney_Array)>=1 then Response.Write(LoginPointmoney_Array(1)) end if%>" /> 
                <span id="LenLoginTime2_Alert"></span>&nbsp;<span id="Login_money_Alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td><input type="Button" name="MoneySubmitButton" onClick="MoneySubmit()"value=" 保存 " /> 
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
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<script>
var selected="Lab_Base";
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
		document.getElementById(selected).className ="hback_1";
		selected="Lab_Base";
		break;
		case 2:
		document.getElementById("Layer1").style.display="none";
		document.getElementById("Layer2").style.display="block";
		document.getElementById("Layer3").style.display="none";	
		document.getElementById("Lab_Other").className="";
		if(selected!="Lab_Other")
		document.getElementById(selected).className ="hback_1";
		selected="Lab_Other";
		break;
		case 3:
		document.getElementById("Layer1").style.display="none";
		document.getElementById("Layer2").style.display="none";	
		document.getElementById("Layer3").style.display="block";	
		document.getElementById("Lab_Money").className="";
		if(selected!="Lab_Money")
		document.getElementById(selected).className ="hback_1";
		selected="Lab_Money";
		break;
	}
}
function BaseSubmit()
{
	var flag1=isNumber('UserName_Length_Min','UserName_Length_alert1','长度应该为正整数',true)
	var flag2=isNumber('UserName_Length_Max','UserName_Length_alert2','长度应该为正整数',true)
	var flag3=isNumber('Pwd_Length_Min','Pwd_Length_alert1','长度应该为正整数',true)
	var flag4=isNumber('Pwd_Length_Max','Pwd_Length_alert2','长度应该为正整数',true)
	var flag5=CheckContentLen('Reg_Help','Reg_alert',2000)
	if(flag1&&flag2&&flag3&&flag4&&flag5)
		document.getElementById("BaseParam").submit();
		
}
function OtherSubmit()
{
	var flag1=isNumber('ErrorPwdTimes','ErrorPwdTimes_alert','次数应该为正整数',true)
	var flag2=isNumber('ShowNumberPerPage','ShowNumberPerPage_alert','页数应该为正整数',true)
	var flag3=isNumber('UpfileSize','UpfileSize_alert','文件大小应该为正整数',true)
	var flag4=isNumber('MessageSize','MessageSize_alert','短信容量应该为正整数',true)
	var flag5=isNumber('LimitClass','LimitClass_alert1','最多分类应该为正整数',true)
	var flag6=isNumber('LimitClass2','LimitClass_alert2','分类的数量应该为正整数',true)
	if(flag1&&flag2&&flag3&&flag4&&flag5&&flag6)
		document.getElementById("OtherParam").submit();
		
}
function MoneySubmit()
{
	var flag0=isEmpty("MoneyName","MoneyName_alert");
	var flag1=isNumber('Reg_Point','Reg_Point_Alert','积分应该为正整数',true)
	var flag2=isNumber('Reg_Money','Reg_Money_Alert','金钱应该为正整数',true)
	var flag3=isNumber('Login_Point','Login_Point_Alert','积分应该为正整数',true)
	var flag4=isNumber('Login_money','Login_Money_Alert','金钱应该为正整数',true)
	var flag5=isNumber('Money_to_Point','Money_Point_Alert','积分应该为正数',false)
	var flag6=isNumber('Point_to_Money','Point_Money_Alert','金钱应该为正数',false)
	var flag7=isNumber('PromptCondition','Prompt_Alert','金钱应该为正整数',true)
	var flag8=isNumber('LenLoginTime','LenLoginTime_Alert','时间应该为正整数',true)
	var flag9=isNumber('LenLoginTime2','LenLoginTime2_Alert','时间应该为正整数',true)
	var flag10=isNumber('txt_contrPoint','contrPoint_Alert','积分应该为正整数',true)
	var flag11=isNumber('txt_contrMoney','contrMoney_Alert','金钱应该为正整数',true)
	var flag12=isNumber('txt_contrAuditPoint','contrAuditPoint_Alert','积分应该为正整数',true)
	var flag13=isNumber('txt_contrAuditMoney','contrAuditMoney_Alert','金钱应该为正整数',true)
	if(flag0&&flag1&&flag2&&flag3&&flag4&&flag5&&flag6&&flag7&&flag8&&flag9&&flag10&&flag11&&flag12&&flag13)
		document.getElementById("MoneyParam").submit();
		
}


function DisAbledPorM(StrID,IDStr)
{
	if (StrID == 'dd')
	{
		document.getElementById('Money_to_Point').disabled = true;
		document.getElementById('Money_to_Point').value = '300';
	}
	else if (StrID == 'd')
	{
		document.getElementById('Money_to_Point').disabled = false;
	}
	if (IDStr == 'dd')
	{
		document.getElementById('Point_to_Money').disabled = true;
		document.getElementById('Point_to_Money').value = '0.0001';
	}
	else if (IDStr == 'd')
	{
		document.getElementById('Point_to_Money').disabled = false;
	}
}
</script>
</html>






