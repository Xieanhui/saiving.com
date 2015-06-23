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
''�õ���ر��ֵ��
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
		strShowErr = "<li>��ҵע����Ҫ��д�������ϣ�</li><li>����д����,��ҵ����,����ʡ��,����,�ʱ�,��ϵ��,��ϵ�绰!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	Dim AddCorpDataTFObj
	Set AddCorpDataTFObj = server.CreateObject(G_FS_RS)
	AddCorpDataTFObj.open "select  C_Name From FS_ME_CorpUser where C_Name = '"& p_C_Name &"' and usernumber<>'"&session("fS_usernumber")&"'",User_Conn,1,3
	If Not AddCorpDataTFObj.eof then
		strShowErr = "<li>���ύ����ҵ�����Ѿ���ע�ᣡ</li>"
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
		Response.Redirect("lib/success.asp?ErrCodes=��ӳɹ���&ErrorUrl=../main.asp")
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
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
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
          <td class="hback"><strong>λ�ã�</strong><a href="../">��վ��ҳ</a> &gt;&gt; 
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt;��д��ҵ����</td>
        </tr>
      </table> 
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
              <form name="UserForm"  id="UserForm" method="post" action="OpenCorp.asp">
                <tr class="back"> 
                  <td height="16" colspan="3" class="xingmu">��д��˾���ϣ� </td>
                </tr>
                <tr class="back"> 
                  <td width="15%" height="16"><div align="right"><span class="tx">*</span>��˾����</div></td>
                  <td width="45%"><input name="C_Name" type="text" id="C_Name" size="30" maxlength="50" value="<%=c_C_Name%>"></td>
                  <td width="40%">����д����˾��ȫ��</td>
                </tr>
                <tr class="back"> 
                  <td height="-2"><div align="right">��˾���</div></td>
                  <td><input name="C_ShortName" type="text" id="C_ShortName" size="30" maxlength="30" value="<%=c_C_ShortName%>"></td>
                  <td>����д����˾�ļ򵥳ƺ�</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��˾����ʡ��</div></td>
                  <td> <select name="C_Province" size=1 id="C_Province">
                      <option value="">��ѡ�񡡡�</option>
                      <option value="�Ĵ�" <%if c_C_Province="�Ĵ�" then response.Write("selected") %>>�Ĵ�</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="�Ϻ�" <%if c_C_Province="�Ϻ�" then response.Write("selected") %>>�Ϻ�</option>
                      <option value="���" <%if c_C_Province="���" then response.Write("selected") %>>���</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="�㶫" <%if c_C_Province="�㶫" then response.Write("selected") %>>�㶫</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="�ӱ�" <%if c_C_Province="�ӱ�" then response.Write("selected") %>>�ӱ�</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="������" <%if c_C_Province="������" then response.Write("selected") %>>������</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="���ɹ�" <%if c_C_Province="���ɹ�" then response.Write("selected") %>>���ɹ�</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="�ຣ" <%if c_C_Province="�ຣ" then response.Write("selected") %>>�ຣ</option>
                      <option value="ɽ��" <%if c_C_Province="ɽ��" then response.Write("selected") %>>ɽ��</option>
                      <option value="ɽ��" <%if c_C_Province="ɽ��" then response.Write("selected") %>>ɽ��</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="�½�" <%if c_C_Province="�½�" then response.Write("selected") %>>�½�</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="�㽭" <%if c_C_Province="�㽭" then response.Write("selected") %>>�㽭</option>
                      <option value="�۰�̨" <%if c_C_Province="�۰�̨" then response.Write("selected") %>>�۰�̨</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                      <option value="����" <%if c_C_Province="����" then response.Write("selected") %>>����</option>
                    </select> </td>
                  <td>����д����˾���ڵ�ʡ��</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��˾���ڳ���</div></td>
                  <td><input name="C_City" type="text" id="C_City" size="30" maxlength="20"  value="<%=c_C_City%>"></td>
                  <td>����д����˾���ڵĳ���</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��˾��ַ</div></td>
                  <td><input name="C_Address" type="text" id="C_Address" size="30" maxlength="100" value="<%=c_C_Address%>"></td>
                  <td>���Ĺ�˾��ַ</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��������</div></td>
                  <td><input name="C_PostCode" type="text" id="C_PostCode" size="30" maxlength="20" value="<%=c_C_PostCode%>"></td>
                  <td>����˾����������</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��˾��ϵ��</div></td>
                  <td><input name="C_ConactName" type="text" id="C_ConactName" size="30" maxlength="20" value="<%=c_C_ConactName%>"></td>
                  <td>��˾��ϵ��</td>
                </tr>
                <tr class="back"> 
                  <td height="0"><div align="right"><span class="tx">*</span>��˾��ϵ�绰</div></td>
                  <td><input name="C_Tel" type="text" id="C_Tel" size="30" maxlength="20" value="<%=c_C_Tel%>" ></td>
                  <td>��˾��ϵ�绰���зֻ�����&quot;-&quot;�ֿ����磺028-85098980-606</td>
                </tr>
                <tr class="back"> 
                  <td height="1"><div align="right">��˾����</div></td>
                  <td><input name="C_Fax" type="text" id="C_Fax" size="30" maxlength="20" value="<%=c_C_Fax%>"></td>
                  <td>��˾���档�зֻ�����&quot;-&quot;�ֿ����磺028-85098980-606</td>
                </tr>
                <tr class="back"> 
                  <td height="3"><div align="right"><span class="tx">*</span>��ҵ</div></td>
                  <td><input type="hidden" name="C_VocationClassID" id="C_VocationClassID" value="<%=c_C_VocationClassID%>">
				  <input name="C_VocationClassName" type="text" id="C_VocationClassName" size="30" readonly value="<%=Get_OtherTable_Value("select vClassName from FS_ME_VocationClass where VCID="&clng(c_C_VocationClassID))%>"></td>
                  <td><input type="button" name="Submit3" value="ѡ����ҵ" onClick="SelectClass();">
                    ��˾���ڵ���ҵ</td>
                </tr>
                <tr class="back"> 
                  <td height="8"><div align="right">��˾��վ</div></td>
                  <td><input name="C_Website" type="text" id="C_Website" size="30" maxlength="200" value="<%=c_C_Website%>"></td>
                  <td>��˾���ڵ���ҵվ��</td>
                </tr>
                <tr class="back"> 
                  <td height="16"><div align="right">��˾��ģ</div></td>
                  <td><select name="C_size" id="C_size">
                      <option value="1-20��" <%if c_C_size="1-20��" then Response.Write("selected") %>>1-20��</option>
                      <option value="21-50��" <%if c_C_size="21-50��" then Response.Write("selected") %>>21-50��</option>
                      <option value="51-100��" <%if c_C_size="51-100��" then Response.Write("selected") %>>51-100��</option>
                      <option value="101-200��" <%if c_C_size="101-200��" then Response.Write("selected") %>>101-200��</option>
                      <option value="201-500��" <%if c_C_size="201-500��" then Response.Write("selected") %>>201-500��</option>
                      <option value="501-1000��" <%if c_C_size="501-1000��" then Response.Write("selected") %>>501-1000��</option>
                      <option value="1000������" <%if c_C_size="1000������" then Response.Write("selected") %>>1000������</option>
                    </select></td>
                  <td>&nbsp;</td>
                </tr>
                <tr class="back"> 
                  <td height="1"><div align="right">��˾ע���ʱ�</div></td>
                  <td><select name="C_Capital" id="C_Capital">
                      <option value="10������" <%if c_C_Capital="10������" then Response.Write("selected") %>>10������</option>
                      <option value="10��-19��" <%if c_C_Capital="10��-19��" then Response.Write("selected") %>>10��-19��</option>
                      <option value="20��-49��" <%if c_C_Capital="20��-49��" then Response.Write("selected") %>>20��-49��</option>
                      <option value="50��-99��" <%if c_C_Capital="50��-99��" then Response.Write("selected") %>>50��-99��</option>
                      <option value="100��-199��" <%if c_C_Capital="100��-199��" then Response.Write("selected") %>>100��-199��</option>
                      <option value="200��-499��" <%if c_C_Capital="200��-499��" then Response.Write("selected") %>>200��-499��</option>
                      <option value="500��-999��" <%if c_C_Capital="500��-999��" then Response.Write("selected") %>>500��-999��</option>
                      <option value="1000������" <%if c_C_Capital="1000������" then Response.Write("selected") %>>1000������</option>
                    </select></td>
                  <td>&nbsp;</td>
                </tr>
                <tr class="back"> 
                  <td height="3"><div align="right">��������</div></td>
                  <td><input name="C_BankName" type="text" id="C_BankName" size="30" maxlength="50" value="<%=c_C_BankName%>"></td>
                  <td rowspan="2"><p>��˾�����ʻ����Է������������ϵ�����С�<br>
                      �����������ӣ��й��������гɶ�����˫骷���<br>
                      �����ʻ�����</p></td>
                </tr>
                <tr class="back"> 
                  <td height="8"><div align="right">�����ʺż��ʻ���</div></td>
                  <td><textarea name="C_BankUserName" cols="40" rows="6" id="C_BankUserName"><%=c_C_BankUserName%></textarea></td>
                </tr>
                <tr class="back"> 
                  <td height="39" colspan="3"> <div align="center"> 
                      <input name="Action" type="hidden" id="Action" value="save">
					  <input type="submit" name="Submit" value="��������,��ʼע��" style="CURSOR:hand">
                      �� 
                      <input type="reset" name="Submit2" value="����">                      �� 
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

<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->