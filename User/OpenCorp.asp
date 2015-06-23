<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
Dim p_C_Name,p_C_ShortName,p_C_Province,p_C_City,p_C_Address,p_C_PostCode,p_C_ConactName,p_C_Tel,p_C_Fax,p_C_VocationClassID,p_C_Website,p_C_size,p_C_Capital,p_C_BankName,p_C_BankUserName
Dim c_C_Name,c_C_ShortName,c_C_Province,c_C_City,c_C_Address,c_C_PostCode,c_C_ConactName,c_C_Tel,c_C_Fax,c_C_VocationClassID,c_C_Website,c_C_size,c_C_Capital,c_C_BankName,c_C_BankUserName
Dim AddCorpDataObj
Dim action
User_GetParm
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


action=request.Form("action")
if action="save" then
	p_C_Name = NoSqlHack(Request.Form("C_Name"))
	p_C_ShortName = NoSqlHack(Request.Form("C_ShortName"))
	p_C_Province = NoSqlHack(Request.Form("C_Province"))
	p_C_City = NoSqlHack(Request.Form("C_City"))
	p_C_Address = NoSqlHack(Request.Form("C_Address"))
	p_C_PostCode = NoSqlHack(Request.Form("C_PostCode"))
	p_C_ConactName = NoSqlHack(Request.Form("C_ConactName"))
	p_C_Tel = NoSqlHack(Request.Form("C_Tel"))
	p_C_Fax = NoSqlHack(Request.Form("C_Fax"))
	p_C_VocationClassID = CintStr(Request.Form("C_VocationClassID"))
	p_C_Website = NoSqlHack(Request.Form("C_Website"))
	p_C_size = NoSqlHack(Request.Form("C_size"))
	p_C_Capital = NoSqlHack(Request.Form("C_Capital"))
	p_C_BankName = NoSqlHack(Request.Form("C_BankName"))
	p_C_BankUserName = NoSqlHack(Request.Form("C_BankUserName"))
	
	If Trim(p_C_Name)="" Or Trim(p_C_Province)="" Or Trim(p_C_City)="" Or Trim(p_C_PostCode)=""  Or Trim(p_C_ConactName)="" Or Trim(p_C_Tel)="" then
		strShowErr = "<li>企业注册需要填写完整资料！</li><li>请填写完整,企业名称,所在省份,城市,邮编,联系人,联系电话!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	Dim AddCorpDataTFObj
	Set AddCorpDataTFObj = server.CreateObject(G_FS_RS)
	AddCorpDataTFObj.open "select  C_Name From FS_ME_CorpUser where C_Name = '"& p_C_Name &"' and usernumber<>'"&session("fS_usernumber")&"'",User_Conn,1,3
	If Not AddCorpDataTFObj.eof then
		strShowErr = "<li>您提交的企业名称已经被注册！</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if	
	Set AddCorpDataObj = server.CreateObject(G_FS_RS)
	AddCorpDataObj.open "select  UserNumber,C_Name,C_ShortName,C_logo,C_Province,C_City,C_Address,C_ConactName,C_Sex,C_Vocation,C_Tel,C_Fax,C_VocationClassID,C_WebSite,C_size,C_Operation,C_Products,C_Content,C_Capital,isLockCorp,C_Templet,C_BankName,C_BankUserName,C_PostCode,isYellowPage,isYellowPageCheck,C_property,C_Byear From FS_ME_CorpUser where usernumber='"&session("FS_Usernumber")&"'",User_Conn,1,3
	if AddCorpDataObj.eof then
		AddCorpDataObj.addNew
	End if
	AddCorpDataObj("UserNumber") = session("FS_UserNumber")
	AddCorpDataObj("C_Name") = p_C_Name
	AddCorpDataObj("C_ShortName") = p_C_ShortName
	AddCorpDataObj("C_Province") = p_C_Province
	AddCorpDataObj("C_City") = p_C_City
	AddCorpDataObj("C_Address") = p_C_Address
	AddCorpDataObj("C_ConactName") = p_C_ConactName
	AddCorpDataObj("C_Tel") = p_C_Tel
	AddCorpDataObj("C_Fax") = p_C_Fax
	AddCorpDataObj("C_VocationClassID") = p_C_VocationClassID
	AddCorpDataObj("C_WebSite") = p_C_WebSite
	AddCorpDataObj("C_size") = p_C_size
	AddCorpDataObj("C_Capital") = p_C_Capital
	AddCorpDataObj("C_BankName") = p_C_BankName
	AddCorpDataObj("C_BankUserName") = p_C_BankUserName
	AddCorpDataObj("C_PostCode") = p_C_PostCode
	AddCorpDataObj("isYellowPage") = 0 
	AddCorpDataObj("isYellowPageCheck") = 0 
	if p_isCheckCorp = 1 then
		AddCorpDataObj("isLockCorp") =1
	Else
		AddCorpDataObj("isLockCorp") =0
	End if
	AddCorpDataObj.update
	AddCorpDataObj.close:set AddCorpDataObj = nothing
	User_Conn.execute("update FS_ME_Users set IsCorporation=1 where usernumber='"&session("FS_usernumber")&"'")
	Set Fs_User = Nothing
	if err.number=0 then
		Response.Redirect("lib/success.asp?ErrCodes=添加成功！&ErrorUrl=../main.asp")
		Response.end
	End if
Else
	Set AddCorpDataObj = server.CreateObject(G_FS_RS)
	AddCorpDataObj.open "select  UserNumber,C_Name,C_ShortName,C_logo,C_Province,C_City,C_Address,C_ConactName,C_Sex,C_Vocation,C_Tel,C_Fax,C_VocationClassID,C_WebSite,C_size,C_Operation,C_Products,C_Content,C_Capital,isLockCorp,C_Templet,C_BankName,C_BankUserName,C_PostCode,isYellowPage,isYellowPageCheck,C_property,C_Byear From FS_ME_CorpUser where usernumber='"&session("FS_Usernumber")&"'",User_Conn,1,3
	if not AddCorpDataObj.eof then
		c_C_Name=AddCorpDataObj("C_Name")
		c_C_ShortName=AddCorpDataObj("C_ShortName")
		c_C_Province=AddCorpDataObj("C_Province")
		c_C_City=AddCorpDataObj("C_City")
		c_C_Address=AddCorpDataObj("C_Address")
		c_C_PostCode=AddCorpDataObj("C_PostCode")
		c_C_ConactName=AddCorpDataObj("C_ConactName")
		c_C_Tel=AddCorpDataObj("C_Tel")
		c_C_Fax=AddCorpDataObj("C_Fax")
		c_C_VocationClassID=AddCorpDataObj("C_VocationClassID")
		c_C_Website=AddCorpDataObj("C_WebSite") 
		c_C_size=AddCorpDataObj("C_size")
		c_C_Capital=AddCorpDataObj("C_Capital")
		c_C_BankName=AddCorpDataObj("C_BankName")
		c_C_BankUserName=AddCorpDataObj("C_BankUserName")
	End if
End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../FS_Inc/prototype.js"></script>
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
            <a href="main.asp">会员首页</a> &gt;&gt;填写企业资料</td>
        </tr>
      </table> 
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <form name="UserForm"  id="UserForm" method="post" action="OpenCorp.asp">
                <tr class="back"> 
                  <td height="16" colspan="3" class="xingmu">填写公司资料： </td>
                </tr>
                <tr class="back"> 
                  <td width="15%" height="16"><div align="right"><span class="tx">*</span>公司名称</div></td>
                  <td width="45%"><input name="C_Name" type="text" id="C_Name" size="30" maxlength="50" value="<%=c_C_Name%>"></td>
                  <td width="40%">请填写您公司的全称</td>
                </tr>
                <tr class="back"> 
                  <td height="-2"><div align="right">公司简称</div></td>
                  <td><input name="C_ShortName" type="text" id="C_ShortName" size="30" maxlength="30" value="<%=c_C_ShortName%>"></td>
                  <td>请填写您公司的简单称呼</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>公司所在省份</div></td>
                  <td> <select name="C_Province" size=1 id="C_Province">
                      <option value="">请选择　　</option>
                      <option value="四川" <%if c_C_Province="四川" then response.Write("selected") %>>四川</option>
                      <option value="北京" <%if c_C_Province="北京" then response.Write("selected") %>>北京</option>
                      <option value="上海" <%if c_C_Province="上海" then response.Write("selected") %>>上海</option>
                      <option value="天津" <%if c_C_Province="天津" then response.Write("selected") %>>天津</option>
                      <option value="重庆" <%if c_C_Province="重庆" then response.Write("selected") %>>重庆</option>
                      <option value="安徽" <%if c_C_Province="安徽" then response.Write("selected") %>>安徽</option>
                      <option value="甘肃" <%if c_C_Province="甘肃" then response.Write("selected") %>>甘肃</option>
                      <option value="广东" <%if c_C_Province="广东" then response.Write("selected") %>>广东</option>
                      <option value="广西" <%if c_C_Province="广西" then response.Write("selected") %>>广西</option>
                      <option value="贵州" <%if c_C_Province="贵州" then response.Write("selected") %>>贵州</option>
                      <option value="福建" <%if c_C_Province="福建" then response.Write("selected") %>>福建</option>
                      <option value="海南" <%if c_C_Province="海南" then response.Write("selected") %>>海南</option>
                      <option value="河北" <%if c_C_Province="河北" then response.Write("selected") %>>河北</option>
                      <option value="河南" <%if c_C_Province="河南" then response.Write("selected") %>>河南</option>
                      <option value="黑龙江" <%if c_C_Province="黑龙江" then response.Write("selected") %>>黑龙江</option>
                      <option value="湖北" <%if c_C_Province="湖北" then response.Write("selected") %>>湖北</option>
                      <option value="湖南" <%if c_C_Province="湖南" then response.Write("selected") %>>湖南</option>
                      <option value="吉林" <%if c_C_Province="吉林" then response.Write("selected") %>>吉林</option>
                      <option value="江苏" <%if c_C_Province="江苏" then response.Write("selected") %>>江苏</option>
                      <option value="江西" <%if c_C_Province="江西" then response.Write("selected") %>>江西</option>
                      <option value="辽宁" <%if c_C_Province="辽宁" then response.Write("selected") %>>辽宁</option>
                      <option value="内蒙古" <%if c_C_Province="内蒙古" then response.Write("selected") %>>内蒙古</option>
                      <option value="宁夏" <%if c_C_Province="宁夏" then response.Write("selected") %>>宁夏</option>
                      <option value="青海" <%if c_C_Province="青海" then response.Write("selected") %>>青海</option>
                      <option value="山东" <%if c_C_Province="山东" then response.Write("selected") %>>山东</option>
                      <option value="山西" <%if c_C_Province="山西" then response.Write("selected") %>>山西</option>
                      <option value="陕西" <%if c_C_Province="陕西" then response.Write("selected") %>>陕西</option>
                      <option value="西藏" <%if c_C_Province="西藏" then response.Write("selected") %>>西藏</option>
                      <option value="新疆" <%if c_C_Province="新疆" then response.Write("selected") %>>新疆</option>
                      <option value="云南" <%if c_C_Province="云南" then response.Write("selected") %>>云南</option>
                      <option value="浙江" <%if c_C_Province="浙江" then response.Write("selected") %>>浙江</option>
                      <option value="港澳台" <%if c_C_Province="港澳台" then response.Write("selected") %>>港澳台</option>
                      <option value="海外" <%if c_C_Province="海外" then response.Write("selected") %>>海外</option>
                      <option value="其它" <%if c_C_Province="其它" then response.Write("selected") %>>其它</option>
                    </select> </td>
                  <td>请填写您公司所在的省份</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>公司所在城市</div></td>
                  <td><input name="C_City" type="text" id="C_City" size="30" maxlength="20"  value="<%=c_C_City%>"></td>
                  <td>请填写您公司所在的城市</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>公司地址</div></td>
                  <td><input name="C_Address" type="text" id="C_Address" size="30" maxlength="100" value="<%=c_C_Address%>"></td>
                  <td>您的公司地址</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>邮政编码</div></td>
                  <td><input name="C_PostCode" type="text" id="C_PostCode" size="30" maxlength="20" value="<%=c_C_PostCode%>"></td>
                  <td>您公司的邮政编码</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>公司联系人</div></td>
                  <td><input name="C_ConactName" type="text" id="C_ConactName" size="30" maxlength="20" value="<%=c_C_ConactName%>"></td>
                  <td>公司联系人</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>公司联系电话</div></td>
                  <td><input name="C_Tel" type="text" id="C_Tel" size="30" maxlength="20" value="<%=c_C_Tel%>" ></td>
                  <td>公司联系电话。有分机请用&quot;-&quot;分开，如：028-85098980-606</td>
                </tr>
                <tr class="back"> 
                  <td height="1"><div align="right">公司传真</div></td>
                  <td><input name="C_Fax" type="text" id="C_Fax" size="30" maxlength="20" value="<%=c_C_Fax%>"></td>
                  <td>公司传真。有分机请用&quot;-&quot;分开，如：028-85098980-606</td>
                </tr>
                <tr class="back"> 
                  <td height="3"><div align="right"><span class="tx">*</span>行业</div></td>
                  <td><input type="hidden" name="C_VocationClassID" id="C_VocationClassID" value="<%=c_C_VocationClassID%>">
				  <input name="C_VocationClassName" type="text" id="C_VocationClassName" size="30" readonly value="<%=Get_OtherTable_Value("select vClassName from FS_ME_VocationClass where VCID="&clng(c_C_VocationClassID))%>"></td>
                  <td><input type="button" name="Submit3" value="选择行业" onClick="SelectClass();">
                    公司所在的行业</td>
                </tr>
                <tr class="back"> 
                  <td height="8"><div align="right">公司网站</div></td>
                  <td><input name="C_Website" type="text" id="C_Website" size="30" maxlength="200" value="<%=c_C_Website%>"></td>
                  <td>公司所在的企业站点</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right">公司规模</div></td>
                  <td><select name="C_size" id="C_size">
                      <option value="1-20人" <%if c_C_size="1-20人" then Response.Write("selected") %>>1-20人</option>
                      <option value="21-50人" <%if c_C_size="21-50人" then Response.Write("selected") %>>21-50人</option>
                      <option value="51-100人" <%if c_C_size="51-100人" then Response.Write("selected") %>>51-100人</option>
                      <option value="101-200人" <%if c_C_size="101-200人" then Response.Write("selected") %>>101-200人</option>
                      <option value="201-500人" <%if c_C_size="201-500人" then Response.Write("selected") %>>201-500人</option>
                      <option value="501-1000人" <%if c_C_size="501-1000人" then Response.Write("selected") %>>501-1000人</option>
                      <option value="1000人以上" <%if c_C_size="1000人以上" then Response.Write("selected") %>>1000人以上</option>
                    </select></td>
                  <td>&nbsp;</td>
                </tr>
                <tr class="back"> 
                  <td height="1"><div align="right">公司注册资本</div></td>
                  <td><select name="C_Capital" id="C_Capital">
                      <option value="10万以下" <%if c_C_Capital="10万以下" then Response.Write("selected") %>>10万以下</option>
                      <option value="10万-19万" <%if c_C_Capital="10万-19万" then Response.Write("selected") %>>10万-19万</option>
                      <option value="20万-49万" <%if c_C_Capital="20万-49万" then Response.Write("selected") %>>20万-49万</option>
                      <option value="50万-99万" <%if c_C_Capital="50万-99万" then Response.Write("selected") %>>50万-99万</option>
                      <option value="100万-199万" <%if c_C_Capital="100万-199万" then Response.Write("selected") %>>100万-199万</option>
                      <option value="200万-499万" <%if c_C_Capital="200万-499万" then Response.Write("selected") %>>200万-499万</option>
                      <option value="500万-999万" <%if c_C_Capital="500万-999万" then Response.Write("selected") %>>500万-999万</option>
                      <option value="1000万以上" <%if c_C_Capital="1000万以上" then Response.Write("selected") %>>1000万以上</option>
                    </select></td>
                  <td>&nbsp;</td>
                </tr>
                <tr class="back"> 
                  <td height="3"><div align="right">开户银行</div></td>
                  <td><input name="C_BankName" type="text" id="C_BankName" size="30" maxlength="50" value="<%=c_C_BankName%>"></td>
                  <td rowspan="2"><p>公司银行帐户，以方便放在您的联系资料中。<br>
                      开户银行例子：中国工商银行成都分行双楠分理处<br>
                      银行帐户名：</p></td>
                </tr>
                <tr class="back"> 
                  <td height="8"><div align="right">银行帐号及帐户名</div></td>
                  <td><textarea name="C_BankUserName" cols="40" rows="6" id="C_BankUserName"><%=c_C_BankUserName%></textarea></td>
                </tr>
                <tr class="back"> 
                  <td height="39" colspan="3"> <div align="center"> 
                      <input name="Action" type="hidden" id="Action" value="save">
					  <input type="submit" name="Submit" value="保存数据,开始注册" style="CURSOR:hand">
                      　 
                      <input type="reset" name="Submit2" value="重置">                      　 
                    </div></td>
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