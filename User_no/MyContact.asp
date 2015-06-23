<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
If Request.Form("Action") = "Save" then
		Dim RsSaveIObj
		Set RsSaveIObj = server.CreateObject(G_FS_RS)
		RsSaveIObj.open "select  UserID,isLock,UserName,RealName,GroupID,Integral,LoginNum,RegTime, LastLoginTime,LastLoginIP,UserNumber,FS_Money,ConNumber,UserID,HomePage,BothYear,Tel,MSN,QQ,Corner,Province,City,Address,PostCode,PassQuestion,SelfIntro,isOpen,Certificate,CertificateCode,Vocation,HeadPic,NickName,Mobile,CloseTime,IsCorporation,isMessage,Email,sex,safeCode,UserLoginCode,HeadPicsize,OnlyLogin,UserFavor,IsMarray From FS_ME_Users where UserNumber = '"& Fs_User.UserNumber &"'",User_Conn,1,3
		RsSaveIObj("tel") = NoSqlHack(Replace(Request.Form("tel"),"''",""))
		RsSaveIObj("mobile") = NoSqlHack(Replace(Request.Form("mobile"),"''",""))
		RsSaveIObj("homepage") = NoSqlHack(Replace(Request.Form("homepage"),"''",""))
		RsSaveIObj("Province")  = NoSqlHack(Replace(Request.Form("Province"),"''",""))
		RsSaveIObj("city")  = NoSqlHack(Replace(Request.Form("city"),"''",""))
		RsSaveIObj("address")  =NoSqlHack(Replace(Request.Form("address"),"''",""))
		RsSaveIObj("postcode")  = NoSqlHack(Replace(Request.Form("postcode"),"''",""))
		If Request.Form("qq")<>"" Then
			RsSaveIObj("qq")  = NoSqlHack(Request.Form("qq"))
		End If
		RsSaveIObj("msn")  = NoSqlHack(Request.Form("msn"))
		RsSaveIObj.update
		RsSaveIObj.close
		set RsSaveIObj = nothing
		strShowErr = "<li>联系资料资料修改成功!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../MyContact.asp")
		Response.end
Else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-我的联系资料</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
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
          <td class="hback"><strong>位置：</strong><a href="../">网站首页</a> &gt;&gt; 
            <a href="main.asp">会员首页</a> &gt;&gt; 联系资料</td>
        </tr>
      </table> 
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="form1" method="post" action="">
          <tr class="hback"> 
            <td width="16%" class="hback_1"><div align="center"><strong>电话</strong></div></td>
            <td width="35%" class="hback"><input name="Tel" type="text" id="Tel" value="<% = Fs_User.Tel%>" size="26" maxlength="20"></td>
            <td width="9%" class="hback_1"><div align="center"><strong>移动电话</strong></div></td>
            <td width="40%" class="hback"><input name="Mobile" type="text" id="Mobile" value="<% = Fs_User.Mobile%>" size="26" maxlength="50"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>电子邮件</strong></div></td>
            <td class="hback"><input name="Email" type="text" id="Email" value="<% = Fs_User.Email%>" size="26" maxlength="100" readonly></td>
            <td class="hback_1"><div align="center"><strong>个人主页</strong></div></td>
            <td class="hback"><input name="homepage" type="text" id="homepage" value="<% = Fs_User.homepage%>" size="26" maxlength="200"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>省份</strong></div></td>
            <td class="hback"><select name="Province" size=1 id="Province">
                <option value="">请选择　　</option>
                <option value="北京" <% If Fs_User.Province ="北京" then response.Write("selected")%>>北京</option>
                <option value="上海" <% If Fs_User.Province ="上海" then response.Write("selected")%>>上海</option>
                <option value="四川" <% If Fs_User.Province ="四川" then response.Write("selected")%>>四川</option>
                <option value="天津" <% If Fs_User.Province ="天津" then response.Write("selected")%>>天津</option>
                <option value="重庆" <% If Fs_User.Province ="重庆" then response.Write("selected")%>>重庆</option>
                <option value="安徽" <% If Fs_User.Province ="安徽" then response.Write("selected")%>>安徽</option>
                <option value="甘肃" <% If Fs_User.Province ="甘肃" then response.Write("selected")%>>甘肃</option>
                <option value="广东" <% If Fs_User.Province ="广东" then response.Write("selected")%>>广东</option>
                <option value="广西" <% If Fs_User.Province ="广西" then response.Write("selected")%>>广西</option>
                <option value="贵州" <% If Fs_User.Province ="贵州" then response.Write("selected")%>>贵州</option>
                <option value="福建" <% If Fs_User.Province ="福建" then response.Write("selected")%>>福建</option>
                <option value="海南" <% If Fs_User.Province ="海南" then response.Write("selected")%>>海南</option>
                <option value="河北" <% If Fs_User.Province ="河北" then response.Write("selected")%>>河北</option>
                <option value="河南" <% If Fs_User.Province ="河南" then response.Write("selected")%>>河南</option>
                <option value="黑龙江" <% If Fs_User.Province ="黑龙江" then response.Write("selected")%>>黑龙江</option>
                <option value="湖北" <% If Fs_User.Province ="湖北" then response.Write("selected")%>>湖北</option>
                <option value="湖南" <% If Fs_User.Province ="湖南" then response.Write("selected")%>>湖南</option>
                <option value="吉林" <% If Fs_User.Province ="吉林" then response.Write("selected")%>>吉林</option>
                <option value="江苏" <% If Fs_User.Province ="江苏" then response.Write("selected")%>>江苏</option>
                <option value="江西" <% If Fs_User.Province ="江西" then response.Write("selected")%>>江西</option>
                <option value="辽宁" <% If Fs_User.Province ="辽宁" then response.Write("selected")%>>辽宁</option>
                <option value="内蒙古" <% If Fs_User.Province ="内蒙古" then response.Write("selected")%>>内蒙古</option>
                <option value="宁夏" <% If Fs_User.Province ="宁夏" then response.Write("selected")%>>宁夏</option>
                <option value="青海" <% If Fs_User.Province ="青海" then response.Write("selected")%>>青海</option>
                <option value="山东" <% If Fs_User.Province ="山东" then response.Write("selected")%>>山东</option>
                <option value="山西" <% If Fs_User.Province ="山西" then response.Write("selected")%>>山西</option>
                <option value="陕西" <% If Fs_User.Province ="陕西" then response.Write("selected")%>>陕西</option>
                <option value="西藏" <% If Fs_User.Province ="西藏" then response.Write("selected")%>>西藏</option>
                <option value="新疆" <% If Fs_User.Province ="新疆" then response.Write("selected")%>>新疆</option>
                <option value="云南" <% If Fs_User.Province ="云南" then response.Write("selected")%>>云南</option>
                <option value="浙江" <% If Fs_User.Province ="浙江" then response.Write("selected")%>>浙江</option>
                <option value="港澳台" <% If Fs_User.Province ="港澳台" then response.Write("selected")%>>港澳台</option>
                <option value="海外" <% If Fs_User.Province ="海外" then response.Write("selected")%>>海外</option>
                <option value="其它" <% If Fs_User.Province ="其它" then response.Write("selected")%>>其它</option>
              </select></td>
            <td class="hback_1"><div align="center"><strong>城市</strong></div></td>
            <td class="hback"><input name="City" type="text" id="City" value="<% = Fs_User.City%>" size="26" maxlength="20"> 
            </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>地址</strong></div></td>
            <td class="hback"><input name="address" type="text" id="address" value="<% = Fs_User.address%>" size="26" maxlength="100"> 
            </td>
            <td class="hback_1"><div align="center"><strong>邮政编码</strong></div></td>
            <td class="hback"><input name="postcode" type="text" id="postcode" value="<% = Fs_User.postcode%>" size="26" maxlength="20"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="center"><strong>QQ</strong></div></td>
            <td class="hback"><input name="QQ" type="text" id="QQ" value="<% = Fs_User.QQ%>" size="26" maxlength="20"> 
            </td>
            <td class="hback_1"><div align="center"><strong>MSN</strong></div></td>
            <td class="hback"><input name="MSN" type="text" id="MSN" value="<% = Fs_User.MSN%>" size="26" maxlength="100"> 
            </td>
          </tr>
          <tr class="hback"> 
            <td colspan="4" class="hback"><div align="center"> 
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="submit" name="Submit" value="保存资料"   onClick="{if(confirm('确认保存您所修改的参数吗?')){this.document.form1.submit();return true;}return false;}">
                　 
                <input type="reset" name="Submit3" value="重新填写">
              </div></td>
          </tr>
          <tr class="hback">
            <td colspan="4" class="hback"><div align="center"></div></td>
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
End if
Set Fs_User = Nothing
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





