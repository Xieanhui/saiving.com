<!--#include file="User_sys.asp" -->
<!--#include file="../../../User/lib/Cls_user.asp" -->
<%
Dim Conn,User_Conn,strShowErr
Dim UpdateUserRs,GetUserDataObj_Rs,GetUserParaObj,UserSql,RegisterNotice,RegisterTF,UserDir,LoginLockNum,p_LoginStyle,p_LenPassworMin,p_LenPassworMax,p_isValidate,p_isPassCard,p_isYellowCheck
Dim p_VerCodeStyle,p_AllowChineseName,p_LenUserName,p_NumLenMax,p_NumLenMin,p_strisPromptTF,p_strisPromptnNum,p_isSendMail,p_LimitUserName,p_PointChange
Dim p_isCheckCorp,p_RegisterCheck,p_MoneyName,p_RegPointmoney,p_NumGetPoint,p_NumGetMoney,p_NumReturnUrl,p_UserNumberRule,p_LoginPointmoneyarr_1,p_LoginPointmoneyarr_2
Dim strLenUserNameArr,strRegPointmoneyArr,p_isPromptarr,p_LenPasswordarr,p_OnlyMemberLogin,p_LenLoginTimearr,p_LoginGetIntegral,p_LoginGetMoney,p_LoginPointmoneyarr
Dim GetConfigObj ,Sql,p_Soft_Version
Dim strCard_month_1,strCard_day_1,strCard_hour_1,strCard_minute_1,strTodaydate,strTodaydate_1
MF_Default_Conn
MF_Session_TF 
MF_User_Conn
%>






