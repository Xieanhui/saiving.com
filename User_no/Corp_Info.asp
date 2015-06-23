<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%

''得到相关表的值。
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if
	if instr(lcase(This_Fun_Sql)," in ")>0 then 
		do while not This_Fun_Rs.eof
			Get_OtherTable_Value = Get_OtherTable_Value & This_Fun_Rs(0) &"&nbsp;"
			This_Fun_Rs.movenext
		loop
	else			
		if not This_Fun_Rs.eof then 
			Get_OtherTable_Value = This_Fun_Rs(0)
		else
			Get_OtherTable_Value = ""
		end if
	end if	
	set This_Fun_Rs=nothing 
End Function


if Request("YellowPage") = "Open" then
	User_GetParm
	If p_isYellowCheck = 0 then
		User_Conn.execute("Update FS_ME_CorpUser set isYellowPage=1,isYellowPageCheck=1 where CorpID="&CintStr(Replace(Request.QueryString("CorpID"),"''","")))
		strShowErr = "<li>开启黄页成功，等待审核</li>"
	Else
		User_Conn.execute("Update FS_ME_CorpUser set isYellowPage=1,isYellowPageCheck=0 where CorpID="&CintStr(Replace(Request.QueryString("CorpID"),"''","")))
		strShowErr = "<li>开启黄页成功，等待审核</li>"
	End if
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Corp_Info.asp")
	Response.end
Elseif  Request("YellowPage") = "Close" then
	User_Conn.execute("Update FS_ME_CorpUser set isYellowPage=0, isYellowPageCheck=0 where CorpID="&CintStr(Replace(Request.QueryString("CorpID"),"''","")))
	strShowErr = "<li>关闭成功</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Corp_Info.asp")
	Response.end
End if
If Request.Form("Action") = "Save" then
	if trim(Request.Form("C_Name")) ="" then 
		strShowErr = "<li>请输入公司名称</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif   len(trim(Request.Form("C_Name")))>250 then
		strShowErr = "<li>业务范围不能超过250个字符</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif  len(trim(Request.Form("C_Name")))>500 then
		strShowErr = "<li>公司产品不能超过500个字符</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif  len(trim(Request.Form("C_Content")))>1000 then
		strShowErr = "<li>公司简介不能超过1000个字符</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	else
		Dim RsSaveIObj
		Set RsSaveIObj = server.CreateObject(G_FS_RS)
		RsSaveIObj.open "select CorpID,UserNumber,C_Name,C_ShortName,C_Province,C_City,C_Address,C_ConactName,C_Vocation,C_Sex,C_Tel,C_Fax,C_VocationClassID,C_WebSite,C_size,C_Operation,C_Products,C_Content,C_Capital,isLockCorp,C_Templet,C_PostCode,C_property,isYellowPage,isYellowPageCheck,C_Byear From FS_ME_CorpUser where UserNumber = '"& Fs_User.UserNumber &"' and CorpID="&CintStr(Request.Form("ID")),User_Conn,1,3
		RsSaveIObj("C_Name") = NoSqlHack(Replace(Request.Form("C_Name"),"''",""))
		RsSaveIObj("C_ShortName") = NoSqlHack(Replace(Request.Form("C_ShortName"),"''",""))
		RsSaveIObj("C_Province") = NoSqlHack(Replace(Request.Form("C_Province"),"''",""))
		RsSaveIObj("C_City")  = NoSqlHack(Replace(Request.Form("C_City"),"''",""))
		RsSaveIObj("C_Address")  = NoSqlHack(Replace(Request.Form("C_Address"),"''",""))
		RsSaveIObj("C_ConactName")  = NoSqlHack(Replace(Request.Form("C_ConactName"),"''",""))
		RsSaveIObj("C_Vocation")  = NoSqlHack(Replace(Request.Form("C_Vocation"),"''",""))
		RsSaveIObj("C_Sex")  = NoSqlHack(Replace(Request.Form("C_Sex"),"''",""))
		RsSaveIObj("C_Tel")  = NoSqlHack(Replace(Request.Form("C_Tel"),"''",""))
		RsSaveIObj("C_Fax")  = NoSqlHack(Replace(Request.Form("C_Fax"),"''",""))
		RsSaveIObj("C_VocationClassID")  = NoSqlHack(Replace(Request.Form("C_VocationClassID"),"''",""))
		RsSaveIObj("C_Fax")  = NoSqlHack(Replace(Request.Form("C_Fax"),"''",""))
		RsSaveIObj("C_size")  = NoSqlHack(Replace(Request.Form("C_size"),"''",""))
		RsSaveIObj("C_WebSite")  = NoSqlHack(Replace(Request.Form("C_WebSite"),"''",""))
		RsSaveIObj("C_Operation")  = NoSqlHack(NoHtmlHackInput(Replace(Request.Form("C_Operation"),"''","")))
		RsSaveIObj("C_Products")  = NoSqlHack(NoHtmlHackInput(Replace(Request.Form("C_Products"),"''","")))
		RsSaveIObj("C_Content")  = NoSqlHack(NoHtmlHackInput(Replace(Request.Form("C_Content"),"''","")))
		RsSaveIObj("C_WebSite")  = NoSqlHack(Replace(Request.Form("C_WebSite"),"''",""))
		RsSaveIObj("C_Capital")  = NoSqlHack(Replace(Request.Form("C_Capital"),"''",""))
		RsSaveIObj("C_PostCode")  = NoSqlHack(Replace(Request.Form("C_PostCode"),"''",""))
		RsSaveIObj("C_property")  = NoSqlHack(Replace(Request.Form("C_property"),"''",""))
		RsSaveIObj("C_Byear")  = NoSqlHack(Replace(Request.Form("C_Byear"),"''",""))
		RsSaveIObj.update
		RsSaveIObj.close
		set RsSaveIObj = nothing
		strShowErr = "<li>企业资料修改成功!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Corp_info.asp")
		Response.end
	End if
Else
		Dim RsCorpObj
		Set RsCorpObj = server.CreateObject(G_FS_RS)
		RsCorpObj.open "select  CorpID,UserNumber,C_Name,C_ShortName,C_Province,C_City,C_Address,C_ConactName,C_Vocation,C_Sex,C_Tel,C_Fax,C_VocationClassID,C_WebSite,C_size,C_Operation,C_Products,C_Content,C_Capital,isLockCorp,C_Templet,C_PostCode,C_property,isYellowPage,isYellowPageCheck,C_Byear From FS_ME_CorpUser where UserNumber = '"& Fs_User.UserNumber &"'",User_Conn,1,3
		if RsCorpObj.eof then
			strShowErr = "<li>找不到企业数据</li>"
			Call ReturnError(strShowErr,"")
		End if
		if RsCorpObj("isLockCorp") = 1 then
			strShowErr = "<li>您的企业数据还没审核通过</li>"
			Call ReturnError(strShowErr,"")
		End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-企业资料</title>
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
            <a href="main.asp">会员首页</a> &gt;&gt; 企业资料</td>
        </tr>
      </table> 
      
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="UserForm" method="post" action="">
          <tr class="hback" style="display:none"> 
            <td colspan="4" class="xingmu"> <strong>
              <%
			if RsCorpObj("isYellowPage") = 1 then
				if RsCorpObj("isYellowPageCheck") = 0 then
					Response.Write("您已经开通黄页功能，但还没审核,&nbsp;&nbsp;<a href=""Corp_Info.asp?CorpID="&RsCorpObj("CorpID") &"&YellowPage=Close"" class=""top_navi""><b>取消</b></a>")
				Else
					Response.Write("您已经开通黄页功能,&nbsp;&nbsp;<a href=""Corp_Info.asp?CorpID="&RsCorpObj("CorpID") &"&YellowPage=Close"" class=""top_navi""><b>关闭</b></a>")
				End if
			Else
					Response.Write("您还没开通黄页功能,&nbsp;&nbsp;<a href=""Corp_Info.asp?CorpID="&RsCorpObj("CorpID") &"&YellowPage=Open"" class=""top_navi""><b>开通</b></a>")
			End if
			%>
              </strong></td>
          </tr>
          <tr class="hback"> 
            <td width="16%" class="hback_1"><div align="right">公司名称</div></td>
            <td width="35%" class="hback"><input name="C_Name" type="text" id="C_Name" value="<% = RsCorpObj("C_Name")%>" size="26" maxlength="50">
              <input name="ID" type="hidden" id="ID" value="<% = RsCorpObj("CorpID")%>"></td>
            <td width="10%" class="hback_1"><div align="right">公司简称</div></td>
            <td width="39%" class="hback"><input name="C_ShortName" type="text" id="C_ShortName" value="<% = RsCorpObj("C_ShortName")%>" size="26" maxlength="30"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">联系人</div></td>
            <td class="hback"><input name="C_ConactName" type="text" id="C_ConactName" value="<% = RsCorpObj("C_ConactName")%>" size="26" maxlength="20"></td>
            <td class="hback_1"><div align="right">联系人职位</div></td>
            <td class="hback"><input name="C_Vocation" type="text" id="C_Vocation" value="<% = RsCorpObj("C_Vocation")%>" size="26" maxlength="30"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">性别</div></td>
            <td class="hback"> <select name="C_Sex" id="C_Sex">
                <option value="0" <%if RsCorpObj("C_Sex") = 0 then response.Write("selected")%>>男</option>
                <option value="1" <%if RsCorpObj("C_Sex") = 1 then response.Write("selected")%>>女</option>
              </select></td>
            <td class="hback_1"><div align="right">联系电话</div></td>
            <td class="hback"><input name="C_Tel" type="text" id="C_Tel2" value="<% = RsCorpObj("C_Tel")%>" size="26" maxlength="24"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">联系传真</div></td>
            <td class="hback"><input name="C_Fax" type="text" id="C_Fax2" value="<% = RsCorpObj("C_Fax")%>" size="26" maxlength="24"></td>
            <td class="hback_1"><div align="right">省份</div></td>
            <td class="hback"><select name="C_Province" size=1 id="select">
                <option value="">请选择　　</option>
                <option value="北京" <% If RsCorpObj("C_Province") ="北京" then response.Write("selected")%>>北京</option>
                <option value="上海" <% If RsCorpObj("C_Province") ="上海" then response.Write("selected")%>>上海</option>
                <option value="四川" <% If RsCorpObj("C_Province") ="四川" then response.Write("selected")%>>四川</option>
                <option value="天津" <% If RsCorpObj("C_Province") ="天津" then response.Write("selected")%>>天津</option>
                <option value="重庆" <% If RsCorpObj("C_Province") ="重庆" then response.Write("selected")%>>重庆</option>
                <option value="安徽" <% If RsCorpObj("C_Province") ="安徽" then response.Write("selected")%>>安徽</option>
                <option value="甘肃" <% If RsCorpObj("C_Province") ="甘肃" then response.Write("selected")%>>甘肃</option>
                <option value="广东" <% If RsCorpObj("C_Province") ="广东" then response.Write("selected")%>>广东</option>
                <option value="广西" <% If RsCorpObj("C_Province") ="广西" then response.Write("selected")%>>广西</option>
                <option value="贵州" <% If RsCorpObj("C_Province") ="贵州" then response.Write("selected")%>>贵州</option>
                <option value="福建" <% If RsCorpObj("C_Province") ="福建" then response.Write("selected")%>>福建</option>
                <option value="海南" <% If RsCorpObj("C_Province") ="海南" then response.Write("selected")%>>海南</option>
                <option value="河北" <% If RsCorpObj("C_Province") ="河北" then response.Write("selected")%>>河北</option>
                <option value="河南" <% If RsCorpObj("C_Province") ="河南" then response.Write("selected")%>>河南</option>
                <option value="黑龙江" <% If RsCorpObj("C_Province") ="黑龙江" then response.Write("selected")%>>黑龙江</option>
                <option value="湖北" <% If RsCorpObj("C_Province") ="湖北" then response.Write("selected")%>>湖北</option>
                <option value="湖南" <% If RsCorpObj("C_Province") ="湖南" then response.Write("selected")%>>湖南</option>
                <option value="吉林" <% If RsCorpObj("C_Province") ="吉林" then response.Write("selected")%>>吉林</option>
                <option value="江苏" <% If RsCorpObj("C_Province") ="江苏" then response.Write("selected")%>>江苏</option>
                <option value="江西" <% If RsCorpObj("C_Province") ="江西" then response.Write("selected")%>>江西</option>
                <option value="辽宁" <% If RsCorpObj("C_Province") ="辽宁" then response.Write("selected")%>>辽宁</option>
                <option value="内蒙古" <% If RsCorpObj("C_Province") ="内蒙古" then response.Write("selected")%>>内蒙古</option>
                <option value="宁夏" <% If RsCorpObj("C_Province") ="宁夏" then response.Write("selected")%>>宁夏</option>
                <option value="青海" <% If RsCorpObj("C_Province") ="青海" then response.Write("selected")%>>青海</option>
                <option value="山东" <% If RsCorpObj("C_Province") ="山东" then response.Write("selected")%>>山东</option>
                <option value="山西" <% If RsCorpObj("C_Province") ="山西" then response.Write("selected")%>>山西</option>
                <option value="陕西" <% If RsCorpObj("C_Province") ="陕西" then response.Write("selected")%>>陕西</option>
                <option value="西藏" <% If RsCorpObj("C_Province") ="西藏" then response.Write("selected")%>>西藏</option>
                <option value="新疆" <% If RsCorpObj("C_Province") ="新疆" then response.Write("selected")%>>新疆</option>
                <option value="云南" <% If RsCorpObj("C_Province") ="云南" then response.Write("selected")%>>云南</option>
                <option value="浙江" <% If RsCorpObj("C_Province") ="浙江" then response.Write("selected")%>>浙江</option>
                <option value="港澳台" <% If RsCorpObj("C_Province") ="港澳台" then response.Write("selected")%>>港澳台</option>
                <option value="海外" <% If RsCorpObj("C_Province") ="海外" then response.Write("selected")%>>海外</option>
                <option value="其它" <% If RsCorpObj("C_Province") ="其它" then response.Write("selected")%>>其它</option>
              </select> </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">城市</div></td>
            <td class="hback"> <input name="C_City" type="text" id="C_City2" value="<% = RsCorpObj("C_City")%>" size="26" maxlength="20"></td>
            <td class="hback_1"><div align="right">地址</div></td>
            <td class="hback"><input name="C_Address" type="text" id="C_Address2" value="<% = RsCorpObj("C_Address")%>" size="26" maxlength="50">            </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">邮政编码</div></td>
            <td class="hback"><input name="C_PostCode" type="text" id="C_PostCode2" value="<% = RsCorpObj("C_PostCode")%>" size="26" maxlength="15"></td>
            <td class="hback_1"><div align="right">网站</div></td>
            <td class="hback"><input name="C_WebSite" type="text" id="C_WebSite2" value="<% = RsCorpObj("C_WebSite")%>" size="26" maxlength="150">            </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">所在行业</div></td>
            <td class="hback"><input type="hidden" name="C_VocationClassID" id="C_VocationClassID" value="<%=RsCorpObj("C_VocationClassID")%>">
				  <input name="C_VocationClassName" type="text" id="C_VocationClassName" size="20" readonly value="<%=Get_OtherTable_Value("select vClassName from FS_ME_VocationClass where VCID="&RsCorpObj("C_VocationClassID"))%>">
                  <input type="button" name="Submit3" value="选择行业" onClick="SelectClass();">                    </td>
            <td class="hback_1"><div align="right">公司性质</div></td>
            <td class="hback"> <select name="C_property" id="C_property">
                <option value="" <% If Trim(RsCorpObj("C_property"))="" then response.Write("selected")%>>公司性质不限</option>
                <option value="外商独资" <% If RsCorpObj("C_property") ="外商独资" then response.Write("selected")%>>外商独资</option>
                <option value="外企办事处" <% If RsCorpObj("C_property") ="外企办事处" then response.Write("selected")%>>外企办事处</option>
                <option value="中外合营(合资,合作)" <% If RsCorpObj("C_property") ="中外合营(合资,合作)" then response.Write("selected")%>>中外合营(合资,合作)</option>
                <option value="私营,民营企业" <% If RsCorpObj("C_property") ="私营,民营企业" then response.Write("selected")%>>私营,民营企业</option>
                <option value="国有企业" <% If RsCorpObj("C_property") ="国有企业" then response.Write("selected")%>>国有企业</option>
                <option value="国内上市公司" <% If RsCorpObj("C_property") ="国内上市公司" then response.Write("selected")%>>国内上市公司</option>
                <option value="政府机关,非盈利机构"<% If RsCorpObj("C_property") ="政府机关,非盈利机构" then response.Write("selected")%>>政府机关,非盈利机构</option>
                <option value="事业单位" <% If RsCorpObj("C_property") ="事业单位" then response.Write("selected")%>>事业单位</option>
                <option value="股份制企业" <% If RsCorpObj("C_property") ="股份制企业" then response.Write("selected")%>>股份制企业</option>
                <option value="其它" <% If RsCorpObj("C_property") ="其它" then response.Write("selected")%>>其它</option>
              </select></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">公司规模</div></td>
            <td colspan="3" class="hback"><select name="C_size" id="C_size">
                <option value="1-20人" <% If RsCorpObj("C_size") ="1-20人" then response.Write("selected")%>>1-20人</option>
                <option value="21-50人" <% If RsCorpObj("C_size") ="21-50人" then response.Write("selected")%>>21-50人</option>
                <option value="51-100人" <% If RsCorpObj("C_size") ="51-100人" then response.Write("selected")%>>51-100人</option>
                <option value="101-200人" <% If RsCorpObj("C_size") ="101-200人" then response.Write("selected")%>>101-200人</option>
                <option value="201-500人"  <% If RsCorpObj("C_size") ="201-500人" then response.Write("selected")%>>201-500人</option>
                <option value="501-1000人"  <% If RsCorpObj("C_size") ="501-1000人" then response.Write("selected")%>>501-1000人</option>
                <option value="1000人以上"  <% If RsCorpObj("C_size") ="1000人以上" then response.Write("selected")%>>1000人以上</option>
              </select> <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">业务范围</div></td>
            <td colspan="3" class="hback"><textarea name="C_Operation" cols="50" rows="5" id="C_Operation" style="width:80%"><% = RsCorpObj("C_Operation")%></textarea>
              不能超过250个字符 
              <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">产品介绍</div></td>
            <td colspan="3" class="hback"> <textarea name="C_Products" cols="50" rows="5" id="C_Products" style="width:80%"><% = RsCorpObj("C_Products")%></textarea>
              不能超过500个字符 <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">公司简介</div></td>
            <td colspan="3" class="hback"><textarea name="C_Content" cols="50" rows="5" id="C_Content" style="width:80%"><% = RsCorpObj("C_Content")%></textarea>
              不能超过1000个字符</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">注册资本</div></td>
            <td colspan="3" class="hback"> <div align="left"> 
                <select name="C_Capital" id="C_Capital">
                  <option value="10万以下" <% If RsCorpObj("C_Capital") ="10万以下" then response.Write("selected")%>>10万以下</option>
                  <option value="10万-19万"  <% If RsCorpObj("C_Capital") ="10万-19万" then response.Write("selected")%>>10万-19万</option>
                  <option value="20万-49万"  <% If RsCorpObj("C_Capital") ="20万-49万" then response.Write("selected")%>>20万-49万</option>
                  <option value="50万-99万"  <% If RsCorpObj("C_Capital") ="50万-99万" then response.Write("selected")%>>50万-99万</option>
                  <option value="100万-199万"  <% If RsCorpObj("C_Capital") ="100万-199万" then response.Write("selected")%>>100万-199万</option>
                  <option value="200万-499万"  <% If RsCorpObj("C_Capital") ="200万-499万" then response.Write("selected")%>>200万-499万</option>
                  <option value="500万-999万"  <% If RsCorpObj("C_Capital") ="500万-999万" then response.Write("selected")%>>500万-999万</option>
                  <option value="1000万以上"  <% If RsCorpObj("C_Capital") ="1000万以上" then response.Write("selected")%>>1000万以上</option>
                </select>
              </div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">成立时间</div></td>
            <td colspan="3" class="hback"> <input name="C_Byear" type="text" id="C_Byear" value="<% = RsCorpObj("C_Byear")%>" size="26" maxlength="20"> 
              <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="4" class="hback"><div align="center"> 
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="submit" name="Submit" value="保存资料"  onClick="{if(confirm('确认保存吗?')){this.document.UserForm.submit();return true;}return false;}">
                　 
                <input type="reset" name="Submit3" value="重新填写">
                　 </div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="4" class="hback"> <div align="center"> </div></td>
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
	RsCorpObj.close
	set RsCorpObj = nothing
End if
Set Fs_User = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
	function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
	{
		var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
		if (ReturnStr!='') SetObj.value=ReturnStr;
		return ReturnStr;
	}
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectClassFrame.asp',400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		document.all.C_VocationClassID.value=TempArray[0]
		document.all.C_VocationClassName.value=TempArray[1]
	}
}
	
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





