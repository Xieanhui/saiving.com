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
      <td width="140" align="right" class="xingmu" colspan="3"><div align="left">��Աϵͳ��������</div></td>
    </tr>
	<tr class="hback"> 
	<td width="33%"  id="Lab_Base"><div align="center"><strong><a href="#" onClick="showParamPanel(1)">������������</a></strong></div></td>
	<td width="33%" height="19" class="hback_1" id="Lab_Other"> <div align="center"><strong><a href="#" onClick="showParamPanel(2)">������������</a></strong></div></td>
	<td width="33%" height="19" class="hback_1" id="Lab_Money"> <div align="center"><strong><a href="#" onClick="showParamPanel(3)">���ֽ������</a></strong></div></td>
	</tr>
    <tr class="hback">
      <td align="right"  colspan="3">
        <div id="Layer1" style="position:relative; z-index:1; left: 0px; top: 0px;"> 
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <form action="ParamAction.asp?Act=BaseParam" method="post" name="BaseParam" id="BaseParam">
            <tr class="hback"> 
              <td width="210" align="right"><div align="right">��Աע��Ĭ�ϻ�Ա�飺</div></td>
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
              <td align="right">��Աϵͳ���ƣ�</td>
              <td><input name="UserSystemName" type="text" id="UserSystemName" size="30" maxlength="30" value="<%=UserSystemName%>"></td>
            </tr>
            <tr class="hback">
              <td align="right">����ע�᣺</td>
              <td><input type="radio" name="AllowReg" value="1" <%if AllowReg=1 then Response.Write("checked") end if%>>
��&nbsp;
<input type="radio" name="AllowReg" value="0" <%if AllowReg=0 then Response.Write("checked") end if %>>
�� </td>
            </tr>
            <tr class="hback"> 
              <td align="right"> ������������ע�᣺</td>
              <td width="743"> <input type="radio" name="AllowChinese" value="1" <% if AllowChinese=1 then Response.Write("checked") end if%>>
                ��&nbsp; <input type="radio" name="AllowChinese" value="0" <% if AllowChinese=0 then Response.Write("checked") end if%>>
                �� </td>
            </tr>
            <tr class="hback"> 
              <td align="right"> ����ע����Ҫ��ˣ�</td>
              <td><input type="radio" name="NeedAudit" value="1" <% if NeedAudit=1 then Response.Write("checked") end if%> >
                ��&nbsp; <input type="radio" name="NeedAudit" value="0" <% if NeedAudit=0 then Response.Write("checked")  end if%>>
                ��</td>
            </tr>
            <tr class="hback"> 
              <td align="right">��ҵע����Ҫ��ˣ�</td>
              <td><input type="radio" name="Corp_NeedAudit" value="1" <% if Corp_NeedAudit=1 then Response.Write("checked") end if%>>
                ��&nbsp; <input type="radio" name="Corp_NeedAudit" value="0" <% if Corp_NeedAudit=0 then Response.Write("checked") end if%>>
                ��</td>
            </tr>
            <tr class="hback"> 
              <td align="right">ֻ����һ���˵�¼��</td>
              <td><input type="radio" name="OnlyMemberLogin" value="1" <% if OnlyMemberLogin=1 then Response.Write("checked")%>/>
                ��&nbsp; <input type="radio" name="OnlyMemberLogin" value="0" <% if OnlyMemberLogin=0 then Response.Write("checked")%>>
                ��</td>
            </tr>
            <tr class="hback"> 
              <td align="right">�û���Ź���_ǰ׺��</td>
              <td> <input name="ID_Prefix" type="text" id="ID_Prefix2" size="50" value="<% if Ubound(IDRule_Array)>0 then Response.Write(IDRule_Array(0))%>"/>              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">�û���Ź���_���Ԫ�أ�</td>
              <td> <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="y" <%if instr(ID_Elem,"y")>0 then Response.Write("checked") end if%>/>
                �� 
                <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="m" <%if instr(ID_Elem,"m")>0 then Response.Write("checked") end if%>/>
                �� 
                <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="d" <%if instr(ID_Elem,"d")>0 then Response.Write("checked") end if%>/>
                �� 
                <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="h" <%if instr(ID_Elem,"h")>0 then Response.Write("checked") end if%>/>
                ʱ 
                <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="i" <%if instr(ID_Elem,"i")>0 then Response.Write("checked") end if%>/>
                �� 
                <input name="ID_Elem" type="checkbox" id="ID_Elem2" value="s" <%if instr(ID_Elem,"s")>0 then Response.Write("checked") end if%>/>
                �� </td>
            </tr>
            <tr class="hback"> 
              <td align="right">�û���Ź���_ID��׺��</td>
              <td><p> 
                  <label> 
                  <input type="radio" name="ID_Postfix" value="2" <% if ID_Postfix=2 then Response.Write("Checked") end if%>/>
                  2λ�����</label>
                  <label> 
                  <input type="radio" name="ID_Postfix" value="3" <% if ID_Postfix=3 then Response.Write("Checked") end if%>>
                  3λ�����</label>
                  <label> 
                  <input type="radio" name="ID_Postfix" value="4" <% if ID_Postfix=4 then Response.Write("Checked") end if%>>
                  4λ�����</label>
                  <label> 
                  <input type="radio" name="ID_Postfix" value="5" <% if ID_Postfix=5 then Response.Write("Checked") end if%>>
                  5λ�����</label>
                  <input name="needWord" type="checkbox" id="needWord2" value="w" <% if instr(ID_Rule,"w")>0 then Response.Write("checked")%> />
                  ��ĸ�������</p></td>
            </tr>
            <tr class="hback"> 
              <td align="right">�û���Ź���_�ָ����</td>
              <td><input name="ID_Devide" type="text" id="ID_Devide2" size="50" value="<% if Ubound(IDRule_Array)>=3 then Response.Write(IDRule_Array(3))%>"/></td>
            </tr>
            <tr class="hback"> 
              <td align="right">�û����������ƣ�</td>
              <td> ��С���ȣ� 
                <input name="UserName_Length_Min" type="text" id="UserName_Length_Min" size="14" value="<% = UserName_Length_Array(0)%>"/>
                ��󳤶ȣ� 
                <input name="UserName_Length_Max" type="text" id="UserName_Length_Max" size="14" value="<% = UserName_Length_Array(1)%>" /> 
                <span id="UserName_Length_alert1"></span><span id="UserName_Length_alert2"></span>              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">��ֹע����û�����</td>
              <td><textarea name="Forbid_UserName" cols="50" id="textarea2" onKeyUp="ReplaceDot('Forbid_UserName')"><%=Forbid_UserName%></textarea></td>
            </tr>
            <tr class="hback"> 
              <td align="right">���볤�����ƣ�</td>
              <td> ��С���ȣ� 
                <input name="Pwd_Length_Min" type="text" id="Pwd_Length_Min" size="14"  value="<% = Pwd_Length_Array(0)%>"/>
                ��󳤶ȣ� 
                <input name="Pwd_Length_Max" type="text" id="Pwd_Length_Max" size="14"  value="<% = Pwd_Length_Array(1)%>"/> 
                <span id="Pwd_Length_alert1"></span><span id="Pwd_Length_alert2"></span ></td>
            </tr>
            <tr class="hback"  style="display:none"> 
              <td align="right">���͵����ʼ���Ϣ��</td>
              <td><input type="radio" name="isSendMail" value="1">
                ��&nbsp; <input type="radio" name="isSendMail" value="0" checked/>
                �� </td>
            </tr>
            <tr class="hback" style="display:none"> 
              <td align="right">�����ʼ���֤��</td>
              <td><input type="radio" name="Email_Aduit" value="1"/>
                ��&nbsp; <input type="radio" name="Email_Aduit" value="0" checked>
                �� </td>
            </tr>
            <tr class="hback"> 
              <td align="right">��Աע����֪��<br>
                ��֧��html�﷨��<br> <span id="Reg_Alert"></span></td>
              <td><textarea name="Reg_Help" cols="80" rows="10" id="Reg_Help"><%=Reg_Help%></textarea></td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td><input type="Button" name="BaseSubmitButton"  onClick="BaseSubmit()" value=" ���� " /> 
                <input type="reset" name="reset" value=" ���� " /></td>
            </tr>
          </form>
        </table>
      </div>	  
        
      <div id="Layer2" style="position:relative;z-index:1; left: 0px; top: 0px; display:none"> 
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <form action="ParamAction.asp?Act=OtherParam" method="post" name="OtherParam" id="OtherParam">
            <tr class="hback"> 
              <td align="right" width="190">��¼��� </td>
              <td><select name="LoginStyle" id="LoginStyle">
                  <option value="1" <%if LoginSytle=1 then Response.Write("selected") end if%>>Ĭ�Ϸ��</option>
                  <option value="2" <%if LoginSytle=2 then Response.Write("selected") end if%>>��ɫ���</option>
                  <option value="3" <%if LoginSytle=3 then Response.Write("selected") end if%>>��ɫ����</option>
                  <option value="4" <%if LoginSytle=4 then Response.Write("selected") end if%>>��������</option>
                  <option value="5" <%if LoginSytle=5 then Response.Write("selected") end if%>>����Ӳ�</option>
                </select></td>
            </tr>
            <tr class="hback"> 
              <td align="right">��¼�ɹ�����ҳ�棺</td>
              <td><p> 
                  <label> </label>
                  <label> 
                  <input type="radio" name="ReturnUrl" value="0" <%if ReturnUrl=0 then Response.Write("checked") end if%>>
                  ��Ա���� </label>
                  <label> 
                  <input type="radio" name="ReturnUrl" value="1" <%if ReturnUrl=1 then Response.Write("checked") end if%> />
                  ��ҳ </label>
                  <label> 
                  <input type="radio" name="ReturnUrl" value="2" <%if ReturnUrl=2 then Response.Write("checked") end if%>>
                  ֮ǰҳ</label>
                  <br>
                </p></td>
            </tr>
            <tr class="hback" style="display:none"> 
              <td align="right">��ҳ�Ƿ���ˣ�</td>
              <td> <input type="radio" name="isYellowCheck" value="1" <%if isYellowCheck=1 then Response.Write("checked")%>>
                �� 
                <input type="radio" name="isYellowCheck" value="0" <%if isYellowCheck=0 then Response.Write("checked")%>>
                ��</td>
            </tr>
            <tr class="hback"> 
              <td align="right">��Ƭ�ղ���Ҫ����</td>
              <td><input type="radio" name="isPassCard" value="1" <%if isPassCard=1 then Response.Write("checked")%>>
                �� 
                <input type="radio" name="isPassCard" value="0" <%if isPassCard=0 then Response.Write("checked")%>>
                ��</td>
            </tr>
            <tr class="hback"> 
              <td align="right">����¼�²⣺</td>
              <td>������� 
                <input name="ErrorPwdTimes" type="text" id="EErrorPwdTimes" value="<%=ErrorPwdTimes%>" size="10" />
                �κ��������û����� ��0Ϊ�����ƣ� <span id="ErrorPwdTimes_alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">��Ա�б�ÿҳ��ʾ����</td>
              <td><input name="ShowNumberPerPage" type="text" id="PShowNumberPerPage" size="50" value="<%=ShowNumberPerPage%>" /> 
                <span id="ShowNumberPerPage_alert"></span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">�ϴ��ļ����ͣ�</td>
              <td><input name="UpfileType" type="text" id="UpfileType" size="50" onKeyUp="ReplaceDot('UpfileType')" value="<%=UpfileType%>"/>
                ���� ��������</td>
            </tr>
            <tr class="hback"> 
              <td align="right">�ϴ���С���ƣ�</td>
              <td><input name="UpfileSize" type="text" id="UpfileSize" size="50"  value="<%=UpfileSize%>" />
                k <span id="UpfileSize_alert">&nbsp;</span>�����ϴ��ļ���С���������û� </td>
            </tr>
            <tr class="hback"> 
              <td align="right">վ�ڶ�������:</td>
              <td><input name="MessageSize" type="text" id="MessageSize" size="50" value="<%=MessageSize%>"/>
                kbyte <span id="MessageSize_alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">��ͨRSS���ģ�</td>
              <td><label> 
                <input type="radio" name="RssFeed" value="1" <%if RssFeed=1 then Response.Write("checked") end if%>>
                �� </label> <label> 
                <input type="radio" name="RssFeed" value="0" <%if RssFeed=0 then Response.Write("checked") end if%>>
                ��</label></td>
            </tr>
            <tr class="hback" style="display:none"> 
              <td align="right">��Ϣ����(ר��,�ļ�)��</td>
              <td>�����ࣺ 
                <input name="LimitClass" type="text" id="LimitClass" size="14"  value="<%if ubound(LimitClass_Array)>=0 then Response.Write(LimitClass_Array(0)) end if%>"> 
                &nbsp;&nbsp;��������� 
                <input name="LimitClass2" type="text" id="LimitClass2" size="14"  value="<%if ubound(LimitClass_Array)>=1 then Response.Write(LimitClass_Array(1)) end if%>"> 
                <span id="LimitClass_Alert1"></span>&nbsp;<span id="LimitClass_Alert2"></span>��������˻�ԱĬ�ϻ�Ա��Ȩ�ޣ��������ò�������</td>
            </tr>
            <tr class="hback" style="display:none;"> 
              <td align="right">�û�Ŀ¼��</td>
              <td><input name="MemberFile" type="text" id="MemberFile"  size="50" value=""/>
                ��Ĭ�ϣ�userfiles��</td>
            </tr>
            <tr class="hback" style="display:none;"> 
              <td align="right">��֤�ļ�Ŀ¼��</td>
              <td><input name="CertDir" type="text" id="CertDir"  value="<%=CertDir%>" size="50"/>
                ��Ĭ�ϣ�certfiles��</td>
            </tr>
            <tr class="hback"> 
              <td align="right">�������۹ؼ��֣�</td>
              <td><textarea name="LimitReviewChar" cols="50" onKeyUp="ReplaceDot('LimitReviewChar')" id="LimitReviewChar"><%=LimitReviewChar%></textarea></td>
            </tr>
            <tr class="hback"> 
              <td align="right">�����Ƿ���Ҫ��ˣ�</td>
              <td><input name="ReviewTF" type="radio" value="1" <%if ReviewTF=1 then response.Write"checked"%>/>
                �� 
                <input name="ReviewTF" type="radio" value="0" <%if ReviewTF=0 then response.Write"checked"%> />
                ��</td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td><input type="Button" name="OtherSubmitButtont" onClick="OtherSubmit()" value=" ���� " /> 
                <input type="reset" name="Submit2" value=" ���� " /></td>
            </tr>
          </form>
        </table>
      </div>
	  <div id="Layer3" style="position:relative;z-index:1; left: 0px; top: 0px; display:none"> 
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
          <form action="ParamAction.asp?Act=MoneyParam" method="post" name="MoneyParam" id="MoneyParam">
            <tr class="hback"> 
              <td align="right">������ƣ�</td>
              <td width="596"> <input name="MoneyName" type="text" id="MoneyName" size="50" value="<%=MoneyName%>" /> 
                <span id="moneyName_alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">ע���ȡ���ֺͽ�ң�</td>
              <td>���֣� 
                <input name="Reg_Point" type="text" id="Reg_Point" size="15"  value="<%if ubound(RegPointmoney_Array)>=0 then Response.Write(RegPointmoney_Array(0)) end if%>" /> 
                &nbsp; ��ң� 
                <input name="Reg_Money" type="text" id="Reg_Money" size="15"  value="<%if ubound(RegPointmoney_Array)>=1 then Response.Write(RegPointmoney_Array(1)) end if%>" /> 
                <span id="Reg_Point_Alert"></span>&nbsp; <span id="Reg_Money_Alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">���ֽ�Ҷһ�������</td>
              <td> 
			  <input type="radio" name="PointChange" value="0" <%if PointChange_TF=0 then Response.Write("checked") end if%> onClick="DisAbledPorM('dd','dd');"/>
                ������ 
                <input type="radio" name="PointChange" value="1" <%if PointChange_TF=1 then Response.Write("checked") end if%> onClick="DisAbledPorM('dd','d');">
                ���ֻ���� 
                <input type="radio" name="PointChange" value="2" <%if PointChange_TF=2 then Response.Write("checked") end if%> onClick="DisAbledPorM('d','dd');">
                ��һ����� 
                <input type="radio" name="PointChange" value="3" <%if PointChange_TF=3 then Response.Write("checked") end if%> onClick="DisAbledPorM('d','d');">
                ����</td>
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
              <td>1��� = 
                <input name="Money_to_Point" type="text"<% = DisAbledPoint %> id="Money_to_Point" size="15" value="<% = PointAutoV %>" />
                ���� | 1���� = 
                <input name="Point_to_Money" type="text"<% = DisAbledMoney %> id="Point_to_Money" size="15" value="<% = MoneyAutoV %>" />
                ��� <span id="Money_Point_Alert"></span>&nbsp; <span id="Point_Money_Alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">Ͷ�影����</td>
              <td>���֣� 
                <input name="txt_contrPoint" type="text" id="txt_contrPoint" value="<%=contrPoint%>">
                |&nbsp;��ң� 
                <input name="txt_contrMoney" type="text" id="txt_contrMoney" value="<%=contrMoney%>"> 
                <span id="contrPoint_Alert"></span>&nbsp; <span id="contrMoney_Alert"></span> 
              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">Ͷ����˽�����</td>
              <td>���֣� 
                <input name="txt_contrAuditPoint" type="text" id="txt_contrAuditPoint" value="<%=contrAuditPoint%>">
                |&nbsp;��ң� 
                <input name="txt_contrAuditMoney" type="text" id="txt_contrAuditMoney" value="<%=contrAuditMoney%>"> 
                <span id="contrAuditPoint_Alert"></span>&nbsp; <span id="contrAuditMoney_Alert"> 
              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">�����ʾ������</td>
              <td> <label> 
                <input type="radio" name="isPrompt" value="1" <%if isPrompt_TF=1 then Response.Write("checked") end if%> />
                ����</label> <label> 
                <input type="radio" name="isPrompt" value="0" <%if isPrompt_TF=0 then Response.Write("checked") end if%> />
                �ر�</label>
                | ���С�� 
                <input name="PromptCondition" type="text" id="PromptCondition" size="15" value="<%if ubound(isPrompt_Array)>=1 then Response.Write(isPrompt_Array(1)) end if%>"/>
                ʱ��ʾ <span id="Prompt_Alert"></span> </td>
            </tr>
            <tr class="hback"> 
              <td align="right">���������ã�</td>
              <td>��¼ 
                <input name="LenLoginTime" type="text" id="LenLoginTime" size="10" value="<%if Ubound(LenLoginTimeArray)>=0 then Response.Write(LenLoginTimeArray(0))%>" />
                �ֺ��ٴε�¼�������� 
                <input name="Login_Point" type="text" id="Login_Point" size="15"  value="<%if ubound(LoginPointmoney_Array)>=0 then Response.Write(LoginPointmoney_Array(0)) end if%>" /> 
                <span id="LenLoginTime_Alert"></span>&nbsp;<span id="Login_Point_Alert"></span> 
              </td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td>��¼ 
                <input name="LenLoginTime2" type="text" id="LenLoginTime2" size="10" value="<%if Ubound(LenLoginTimeArray)>=1 then Response.Write(LenLoginTimeArray(1))%>" />
                ����ٴε�¼������� 
                <input name="Login_money" type="text" id="Login_money" size="15"  value="<%if ubound(LoginPointmoney_Array)>=1 then Response.Write(LoginPointmoney_Array(1)) end if%>" /> 
                <span id="LenLoginTime2_Alert"></span>&nbsp;<span id="Login_money_Alert"></span></td>
            </tr>
            <tr class="hback"> 
              <td align="right">&nbsp;</td>
              <td><input type="Button" name="MoneySubmitButton" onClick="MoneySubmit()"value=" ���� " /> 
                <input type="reset" name="Submit2" value=" ���� " /></td>
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
	var flag1=isNumber('UserName_Length_Min','UserName_Length_alert1','����Ӧ��Ϊ������',true)
	var flag2=isNumber('UserName_Length_Max','UserName_Length_alert2','����Ӧ��Ϊ������',true)
	var flag3=isNumber('Pwd_Length_Min','Pwd_Length_alert1','����Ӧ��Ϊ������',true)
	var flag4=isNumber('Pwd_Length_Max','Pwd_Length_alert2','����Ӧ��Ϊ������',true)
	var flag5=CheckContentLen('Reg_Help','Reg_alert',2000)
	if(flag1&&flag2&&flag3&&flag4&&flag5)
		document.getElementById("BaseParam").submit();
		
}
function OtherSubmit()
{
	var flag1=isNumber('ErrorPwdTimes','ErrorPwdTimes_alert','����Ӧ��Ϊ������',true)
	var flag2=isNumber('ShowNumberPerPage','ShowNumberPerPage_alert','ҳ��Ӧ��Ϊ������',true)
	var flag3=isNumber('UpfileSize','UpfileSize_alert','�ļ���СӦ��Ϊ������',true)
	var flag4=isNumber('MessageSize','MessageSize_alert','��������Ӧ��Ϊ������',true)
	var flag5=isNumber('LimitClass','LimitClass_alert1','������Ӧ��Ϊ������',true)
	var flag6=isNumber('LimitClass2','LimitClass_alert2','���������Ӧ��Ϊ������',true)
	if(flag1&&flag2&&flag3&&flag4&&flag5&&flag6)
		document.getElementById("OtherParam").submit();
		
}
function MoneySubmit()
{
	var flag0=isEmpty("MoneyName","MoneyName_alert");
	var flag1=isNumber('Reg_Point','Reg_Point_Alert','����Ӧ��Ϊ������',true)
	var flag2=isNumber('Reg_Money','Reg_Money_Alert','��ǮӦ��Ϊ������',true)
	var flag3=isNumber('Login_Point','Login_Point_Alert','����Ӧ��Ϊ������',true)
	var flag4=isNumber('Login_money','Login_Money_Alert','��ǮӦ��Ϊ������',true)
	var flag5=isNumber('Money_to_Point','Money_Point_Alert','����Ӧ��Ϊ����',false)
	var flag6=isNumber('Point_to_Money','Point_Money_Alert','��ǮӦ��Ϊ����',false)
	var flag7=isNumber('PromptCondition','Prompt_Alert','��ǮӦ��Ϊ������',true)
	var flag8=isNumber('LenLoginTime','LenLoginTime_Alert','ʱ��Ӧ��Ϊ������',true)
	var flag9=isNumber('LenLoginTime2','LenLoginTime2_Alert','ʱ��Ӧ��Ϊ������',true)
	var flag10=isNumber('txt_contrPoint','contrPoint_Alert','����Ӧ��Ϊ������',true)
	var flag11=isNumber('txt_contrMoney','contrMoney_Alert','��ǮӦ��Ϊ������',true)
	var flag12=isNumber('txt_contrAuditPoint','contrAuditPoint_Alert','����Ӧ��Ϊ������',true)
	var flag13=isNumber('txt_contrAuditMoney','contrAuditMoney_Alert','��ǮӦ��Ϊ������',true)
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






