<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%

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


if Request("YellowPage") = "Open" then
	User_GetParm
	If p_isYellowCheck = 0 then
		User_Conn.execute("Update FS_ME_CorpUser set isYellowPage=1,isYellowPageCheck=1 where CorpID="&CintStr(Replace(Request.QueryString("CorpID"),"''","")))
		strShowErr = "<li>������ҳ�ɹ����ȴ����</li>"
	Else
		User_Conn.execute("Update FS_ME_CorpUser set isYellowPage=1,isYellowPageCheck=0 where CorpID="&CintStr(Replace(Request.QueryString("CorpID"),"''","")))
		strShowErr = "<li>������ҳ�ɹ����ȴ����</li>"
	End if
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Corp_Info.asp")
	Response.end
Elseif  Request("YellowPage") = "Close" then
	User_Conn.execute("Update FS_ME_CorpUser set isYellowPage=0, isYellowPageCheck=0 where CorpID="&CintStr(Replace(Request.QueryString("CorpID"),"''","")))
	strShowErr = "<li>�رճɹ�</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Corp_Info.asp")
	Response.end
End if
If Request.Form("Action") = "Save" then
	if trim(Request.Form("C_Name")) ="" then 
		strShowErr = "<li>�����빫˾����</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif   len(trim(Request.Form("C_Name")))>250 then
		strShowErr = "<li>ҵ��Χ���ܳ���250���ַ�</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif  len(trim(Request.Form("C_Name")))>500 then
		strShowErr = "<li>��˾��Ʒ���ܳ���500���ַ�</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	Elseif  len(trim(Request.Form("C_Content")))>1000 then
		strShowErr = "<li>��˾��鲻�ܳ���1000���ַ�</li>"
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
		strShowErr = "<li>��ҵ�����޸ĳɹ�!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Corp_info.asp")
		Response.end
	End if
Else
		Dim RsCorpObj
		Set RsCorpObj = server.CreateObject(G_FS_RS)
		RsCorpObj.open "select  CorpID,UserNumber,C_Name,C_ShortName,C_Province,C_City,C_Address,C_ConactName,C_Vocation,C_Sex,C_Tel,C_Fax,C_VocationClassID,C_WebSite,C_size,C_Operation,C_Products,C_Content,C_Capital,isLockCorp,C_Templet,C_PostCode,C_property,isYellowPage,isYellowPageCheck,C_Byear From FS_ME_CorpUser where UserNumber = '"& Fs_User.UserNumber &"'",User_Conn,1,3
		if RsCorpObj.eof then
			strShowErr = "<li>�Ҳ�����ҵ����</li>"
			Call ReturnError(strShowErr,"")
		End if
		if RsCorpObj("isLockCorp") = 1 then
			strShowErr = "<li>������ҵ���ݻ�û���ͨ��</li>"
			Call ReturnError(strShowErr,"")
		End if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%>-��ҵ����</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
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
          <td class="hback"><strong>λ�ã�</strong><a href="../">��վ��ҳ</a> &gt;&gt; 
            <a href="main.asp">��Ա��ҳ</a> &gt;&gt; ��ҵ����</td>
        </tr>
      </table> 
      
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="UserForm" method="post" action="">
          <tr class="hback" style="display:none"> 
            <td colspan="4" class="xingmu"> <strong>
              <%
			if RsCorpObj("isYellowPage") = 1 then
				if RsCorpObj("isYellowPageCheck") = 0 then
					Response.Write("���Ѿ���ͨ��ҳ���ܣ�����û���,&nbsp;&nbsp;<a href=""Corp_Info.asp?CorpID="&RsCorpObj("CorpID") &"&YellowPage=Close"" class=""top_navi""><b>ȡ��</b></a>")
				Else
					Response.Write("���Ѿ���ͨ��ҳ����,&nbsp;&nbsp;<a href=""Corp_Info.asp?CorpID="&RsCorpObj("CorpID") &"&YellowPage=Close"" class=""top_navi""><b>�ر�</b></a>")
				End if
			Else
					Response.Write("����û��ͨ��ҳ����,&nbsp;&nbsp;<a href=""Corp_Info.asp?CorpID="&RsCorpObj("CorpID") &"&YellowPage=Open"" class=""top_navi""><b>��ͨ</b></a>")
			End if
			%>
              </strong></td>
          </tr>
          <tr class="hback"> 
            <td width="16%" class="hback_1"><div align="right">��˾����</div></td>
            <td width="35%" class="hback"><input name="C_Name" type="text" id="C_Name" value="<% = RsCorpObj("C_Name")%>" size="26" maxlength="50">
              <input name="ID" type="hidden" id="ID" value="<% = RsCorpObj("CorpID")%>"></td>
            <td width="10%" class="hback_1"><div align="right">��˾���</div></td>
            <td width="39%" class="hback"><input name="C_ShortName" type="text" id="C_ShortName" value="<% = RsCorpObj("C_ShortName")%>" size="26" maxlength="30"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">��ϵ��</div></td>
            <td class="hback"><input name="C_ConactName" type="text" id="C_ConactName" value="<% = RsCorpObj("C_ConactName")%>" size="26" maxlength="20"></td>
            <td class="hback_1"><div align="right">��ϵ��ְλ</div></td>
            <td class="hback"><input name="C_Vocation" type="text" id="C_Vocation" value="<% = RsCorpObj("C_Vocation")%>" size="26" maxlength="30"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">�Ա�</div></td>
            <td class="hback"> <select name="C_Sex" id="C_Sex">
                <option value="0" <%if RsCorpObj("C_Sex") = 0 then response.Write("selected")%>>��</option>
                <option value="1" <%if RsCorpObj("C_Sex") = 1 then response.Write("selected")%>>Ů</option>
              </select></td>
            <td class="hback_1"><div align="right">��ϵ�绰</div></td>
            <td class="hback"><input name="C_Tel" type="text" id="C_Tel2" value="<% = RsCorpObj("C_Tel")%>" size="26" maxlength="24"></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">��ϵ����</div></td>
            <td class="hback"><input name="C_Fax" type="text" id="C_Fax2" value="<% = RsCorpObj("C_Fax")%>" size="26" maxlength="24"></td>
            <td class="hback_1"><div align="right">ʡ��</div></td>
            <td class="hback"><select name="C_Province" size=1 id="select">
                <option value="">��ѡ�񡡡�</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="�Ϻ�" <% If RsCorpObj("C_Province") ="�Ϻ�" then response.Write("selected")%>>�Ϻ�</option>
                <option value="�Ĵ�" <% If RsCorpObj("C_Province") ="�Ĵ�" then response.Write("selected")%>>�Ĵ�</option>
                <option value="���" <% If RsCorpObj("C_Province") ="���" then response.Write("selected")%>>���</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="�㶫" <% If RsCorpObj("C_Province") ="�㶫" then response.Write("selected")%>>�㶫</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="�ӱ�" <% If RsCorpObj("C_Province") ="�ӱ�" then response.Write("selected")%>>�ӱ�</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="������" <% If RsCorpObj("C_Province") ="������" then response.Write("selected")%>>������</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="���ɹ�" <% If RsCorpObj("C_Province") ="���ɹ�" then response.Write("selected")%>>���ɹ�</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="�ຣ" <% If RsCorpObj("C_Province") ="�ຣ" then response.Write("selected")%>>�ຣ</option>
                <option value="ɽ��" <% If RsCorpObj("C_Province") ="ɽ��" then response.Write("selected")%>>ɽ��</option>
                <option value="ɽ��" <% If RsCorpObj("C_Province") ="ɽ��" then response.Write("selected")%>>ɽ��</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="�½�" <% If RsCorpObj("C_Province") ="�½�" then response.Write("selected")%>>�½�</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="�㽭" <% If RsCorpObj("C_Province") ="�㽭" then response.Write("selected")%>>�㽭</option>
                <option value="�۰�̨" <% If RsCorpObj("C_Province") ="�۰�̨" then response.Write("selected")%>>�۰�̨</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
                <option value="����" <% If RsCorpObj("C_Province") ="����" then response.Write("selected")%>>����</option>
              </select> </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">����</div></td>
            <td class="hback"> <input name="C_City" type="text" id="C_City2" value="<% = RsCorpObj("C_City")%>" size="26" maxlength="20"></td>
            <td class="hback_1"><div align="right">��ַ</div></td>
            <td class="hback"><input name="C_Address" type="text" id="C_Address2" value="<% = RsCorpObj("C_Address")%>" size="26" maxlength="50">            </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">��������</div></td>
            <td class="hback"><input name="C_PostCode" type="text" id="C_PostCode2" value="<% = RsCorpObj("C_PostCode")%>" size="26" maxlength="15"></td>
            <td class="hback_1"><div align="right">��վ</div></td>
            <td class="hback"><input name="C_WebSite" type="text" id="C_WebSite2" value="<% = RsCorpObj("C_WebSite")%>" size="26" maxlength="150">            </td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">������ҵ</div></td>
            <td class="hback"><input type="hidden" name="C_VocationClassID" id="C_VocationClassID" value="<%=RsCorpObj("C_VocationClassID")%>">
				  <input name="C_VocationClassName" type="text" id="C_VocationClassName" size="20" readonly value="<%=Get_OtherTable_Value("select vClassName from FS_ME_VocationClass where VCID="&RsCorpObj("C_VocationClassID"))%>">
                  <input type="button" name="Submit3" value="ѡ����ҵ" onClick="SelectClass();">                    </td>
            <td class="hback_1"><div align="right">��˾����</div></td>
            <td class="hback"> <select name="C_property" id="C_property">
                <option value="" <% If Trim(RsCorpObj("C_property"))="" then response.Write("selected")%>>��˾���ʲ���</option>
                <option value="���̶���" <% If RsCorpObj("C_property") ="���̶���" then response.Write("selected")%>>���̶���</option>
                <option value="������´�" <% If RsCorpObj("C_property") ="������´�" then response.Write("selected")%>>������´�</option>
                <option value="�����Ӫ(����,����)" <% If RsCorpObj("C_property") ="�����Ӫ(����,����)" then response.Write("selected")%>>�����Ӫ(����,����)</option>
                <option value="˽Ӫ,��Ӫ��ҵ" <% If RsCorpObj("C_property") ="˽Ӫ,��Ӫ��ҵ" then response.Write("selected")%>>˽Ӫ,��Ӫ��ҵ</option>
                <option value="������ҵ" <% If RsCorpObj("C_property") ="������ҵ" then response.Write("selected")%>>������ҵ</option>
                <option value="�������й�˾" <% If RsCorpObj("C_property") ="�������й�˾" then response.Write("selected")%>>�������й�˾</option>
                <option value="��������,��ӯ������"<% If RsCorpObj("C_property") ="��������,��ӯ������" then response.Write("selected")%>>��������,��ӯ������</option>
                <option value="��ҵ��λ" <% If RsCorpObj("C_property") ="��ҵ��λ" then response.Write("selected")%>>��ҵ��λ</option>
                <option value="�ɷ�����ҵ" <% If RsCorpObj("C_property") ="�ɷ�����ҵ" then response.Write("selected")%>>�ɷ�����ҵ</option>
                <option value="����" <% If RsCorpObj("C_property") ="����" then response.Write("selected")%>>����</option>
              </select></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">��˾��ģ</div></td>
            <td colspan="3" class="hback"><select name="C_size" id="C_size">
                <option value="1-20��" <% If RsCorpObj("C_size") ="1-20��" then response.Write("selected")%>>1-20��</option>
                <option value="21-50��" <% If RsCorpObj("C_size") ="21-50��" then response.Write("selected")%>>21-50��</option>
                <option value="51-100��" <% If RsCorpObj("C_size") ="51-100��" then response.Write("selected")%>>51-100��</option>
                <option value="101-200��" <% If RsCorpObj("C_size") ="101-200��" then response.Write("selected")%>>101-200��</option>
                <option value="201-500��"  <% If RsCorpObj("C_size") ="201-500��" then response.Write("selected")%>>201-500��</option>
                <option value="501-1000��"  <% If RsCorpObj("C_size") ="501-1000��" then response.Write("selected")%>>501-1000��</option>
                <option value="1000������"  <% If RsCorpObj("C_size") ="1000������" then response.Write("selected")%>>1000������</option>
              </select> <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">ҵ��Χ</div></td>
            <td colspan="3" class="hback"><textarea name="C_Operation" cols="50" rows="5" id="C_Operation" style="width:80%"><% = RsCorpObj("C_Operation")%></textarea>
              ���ܳ���250���ַ� 
              <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">��Ʒ����</div></td>
            <td colspan="3" class="hback"> <textarea name="C_Products" cols="50" rows="5" id="C_Products" style="width:80%"><% = RsCorpObj("C_Products")%></textarea>
              ���ܳ���500���ַ� <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">��˾���</div></td>
            <td colspan="3" class="hback"><textarea name="C_Content" cols="50" rows="5" id="C_Content" style="width:80%"><% = RsCorpObj("C_Content")%></textarea>
              ���ܳ���1000���ַ�</td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">ע���ʱ�</div></td>
            <td colspan="3" class="hback"> <div align="left"> 
                <select name="C_Capital" id="C_Capital">
                  <option value="10������" <% If RsCorpObj("C_Capital") ="10������" then response.Write("selected")%>>10������</option>
                  <option value="10��-19��"  <% If RsCorpObj("C_Capital") ="10��-19��" then response.Write("selected")%>>10��-19��</option>
                  <option value="20��-49��"  <% If RsCorpObj("C_Capital") ="20��-49��" then response.Write("selected")%>>20��-49��</option>
                  <option value="50��-99��"  <% If RsCorpObj("C_Capital") ="50��-99��" then response.Write("selected")%>>50��-99��</option>
                  <option value="100��-199��"  <% If RsCorpObj("C_Capital") ="100��-199��" then response.Write("selected")%>>100��-199��</option>
                  <option value="200��-499��"  <% If RsCorpObj("C_Capital") ="200��-499��" then response.Write("selected")%>>200��-499��</option>
                  <option value="500��-999��"  <% If RsCorpObj("C_Capital") ="500��-999��" then response.Write("selected")%>>500��-999��</option>
                  <option value="1000������"  <% If RsCorpObj("C_Capital") ="1000������" then response.Write("selected")%>>1000������</option>
                </select>
              </div></td>
          </tr>
          <tr class="hback"> 
            <td class="hback_1"><div align="right">����ʱ��</div></td>
            <td colspan="3" class="hback"> <input name="C_Byear" type="text" id="C_Byear" value="<% = RsCorpObj("C_Byear")%>" size="26" maxlength="20"> 
              <div align="center"></div></td>
          </tr>
          <tr class="hback"> 
            <td colspan="4" class="hback"><div align="center"> 
                <input name="Action" type="hidden" id="Action" value="Save">
                <input type="submit" name="Submit" value="��������"  onClick="{if(confirm('ȷ�ϱ�����?')){this.document.UserForm.submit();return true;}return false;}">
                �� 
                <input type="reset" name="Submit3" value="������д">
                �� </div></td>
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
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0ϵ��-->





