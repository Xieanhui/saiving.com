<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<% 

Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
'session�ж�
MF_Session_TF
'Call MF_Check_Pop_TF("NS_Class_000001")
'�õ���Ա���б�
Dim Fs_news,NS_SpecialCNameValure,sRootDir,strShowErr ,obj_Special_Rs,str_newsDir,str_CurrPath
set Fs_news = new Cls_News
Fs_News.GetSysParam()
MF_GetUserGroupID  

Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if

'-----------------------------------------
str_newsDir = Fs_news.newsDir & ""
If str_newsDir = "" Or str_newsDir = "/" Then
	str_newsDir = "/"
Else
	str_newsDir = Replace(str_newsDir,"//","/")
End If
'---------------------------------------------
'��������
Dim lng_SpecialID,str_Templet ,str_SpecialCName,str_SpecialEName,str_SpecialSize,str_SpecialContent,str_naviPic
Dim str_SavePath,str_ExtName,bit_isLock,dtm_Addtime,int_sPoint,lng_GroupID,lng_AdminID,Int_SaveType
''+++++++++++++++++++
if Request.QueryString("Action")="add" then
	if not Get_SubPop_TF("","NS026","NS","specail") then Err_Show
	str_SpecialCName = ""
	str_SpecialEName = ""
	str_SpecialSize = "150,120"
	str_SpecialContent = ""
	str_SavePath = Fs_news.GetSysParamDir
	str_Templet = Replace("/"&G_TEMPLETS_DIR&"/NewsClass/Special.htm","//","/")
	str_ExtName = "html"  ''ר����չ��
	bit_isLock = 0
	dtm_Addtime = now
	lng_GroupID = ""
	int_sPoint = ""
	Int_SaveType = 3
Elseif Request.QueryString("Action")="edit" then
	lng_SpecialID = NoSqlHack(Trim(Request.QueryString("SpecialID")))
	if not Get_SubPop_TF(lng_SpecialID,"NS027","NS","specail") then  Err_Show
	if lng_SpecialID="" or not isnumeric(lng_SpecialID) then 
			strShowErr = "<li>��Ҫ��ID�����ṩ�����������֡�</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	end if
	Set obj_Special_Rs = server.CreateObject(G_FS_RS)
	obj_Special_Rs.open "select SpecialID,SpecialCName,SpecialEName,SpecialSize,SpecialContent,SavePath,Templet,ExtName,isLock,Addtime,naviPic,adminName,FileSaveType From FS_NS_Special where SpecialID = "& NoSqlHack(lng_SpecialID) ,Conn,1,3
	if  not obj_Special_Rs.eof then
		str_SpecialCName = obj_Special_Rs("SpecialCName")
		str_SpecialEName = obj_Special_Rs("SpecialEName")
		str_SpecialSize = obj_Special_Rs("SpecialSize")
		str_SpecialContent = obj_Special_Rs("SpecialContent")
		str_SavePath = obj_Special_Rs("SavePath")
		str_Templet = obj_Special_Rs("Templet")
		str_ExtName = obj_Special_Rs("ExtName")
		bit_isLock = obj_Special_Rs("isLock")
		dtm_Addtime = obj_Special_Rs("Addtime")
		str_naviPic=obj_Special_Rs("naviPic")
		lng_AdminID = obj_Special_Rs("adminName")
		Int_SaveType = obj_Special_Rs("FileSaveType")
		obj_Special_Rs.close
		set  obj_Special_Rs = nothing
	Else
		strShowErr = "<li>����Ĳ���</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ר�����___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="js/Public.js"></script>
<script language="JavaScript" src="../../FS_Inc/GetLettersByChinese.js"></script>
</head>
  <body>
<form name="MainForm" method="post" action="Special_Save.asp">
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback"> 
      <td class="xingmu">��Ŀ����</td>
    </tr>
    <tr> 
      <td height="18" class="hback"><a href="Special_Manage.asp">������ҳ</a> | <a href="Special_Add.asp?Action=add">����ר��</a> 
        | <a href="../../help?Lable=NS_Special_Add" target="_blank" style="cursor:help;"><img src="../Images/_help.gif" border="0"></a></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr class="hback"> 
      <td colspan="3" class="xingmu"><%if request.QueryString("Action")="edit" then response.Write("�޸�ר��") else response.Write("���ר��") end if%></td>
    </tr>
    <tr> 
      <td width="18%" class="hback"><div align="right">ר���������ƣ�</div></td>
      <td width="82%" colspan="2" class="hback"><input name="SpecialCName" type="text" <% if Request.QueryString("Action")="add" then %>onBlur="SetClassEName(value,document.MainForm.SpecialEName);"<% end if%> id="SpecialCName" size="35" maxlength="100" value="<% = str_SpecialCName%>"> 
        <span class="tx"> *3-100���ַ�</span></td>
    </tr>
    <tr> 
      <td width="18%" class="hback"><div align="right">ר��Ӣ�����ƣ�</div></td>
      <td width="82%" class="hback"><input name="SpecialEName" type="text" id="SpecialEName" <% if Request.QueryString("Action")="add" then %>onFocus="SetClassEName(value,document.MainForm.SpecialEName);"<% end if %> size="35" maxlength="50" value="<% =str_SpecialEName%>" <%if Request.QueryString("Action")="edit" then response.Write("Readonly")%> onKeyUp="value=value.replace(/[^a-zA-Z0-9_-]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^a-zA-Z0-9_-]/g,''))">   
        <span class="tx"> *<br>
        3-50���ַ�,��������ĸ�����֣��л��ߣ��»���,@,.��һ��ȷ��,�������޸�<br />Ӣ����֮�䲻Ҫ�������������"aaa"��"aaaa"����������û���</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">ר��߶ȺͿ�ȣ�</div></td>
      <td class="hback"><input name="SpecialSize" type="text" id="SpecialSize" size="35" maxlength="150" value="<% = str_SpecialSize %>" onKeyUp="value=value.replace(/[^0-9,0-9]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^0-9,0-9]/g,''))">
        <span class="tx">*</span>��ʽ���߶�,��ȣ�150,120��</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ר��ͼƬ</div></td>
      <td class="hback"><input name="naviPic" type="text" id="naviPic" value="<% = str_naviPic %>" size="35"  maxlength="255">
        <img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>',500,320,window,document.MainForm.naviPic);" style="cursor:hand;"> </td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">ר�⵼��˵����</div></td>
      <td class="hback"><textarea name="SpecialContent" style="width:80%" rows="6" id="SpecialContent"><% = str_SpecialContent %></textarea>      </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">����Ա��</div></td>
      <td class="hback"><SELECT name="AdminName" id="AdminName">
        <%
			Dim obj_AdminList_Rs
			set obj_AdminList_Rs = Conn.Execute("Select Admin_Name,Admin_Real_Name from FS_MF_Admin Where Admin_Parent_Admin='"&Temp_Admin_Name&"' or Admin_Name='"&Temp_Admin_Name&"' order by ID asc")
			If not obj_AdminList_Rs.eof Then
				if lng_AdminID = obj_AdminList_Rs("Admin_Name") then
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """ selected>����Ա�ʺţ�" & obj_AdminList_Rs("Admin_Name") & "������Ա������" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				Else
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """>����Ա�ʺţ�" & obj_AdminList_Rs("Admin_Name") & "������Ա������" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				End if
				obj_AdminList_Rs.MoveNext
			End If
			Do while not obj_AdminList_Rs.eof
				if lng_AdminID = obj_AdminList_Rs("Admin_Name") then
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """ selected>����Ա�ʺţ�" & obj_AdminList_Rs("Admin_Name") & "������Ա������" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				Else
					Response.Write "<OPTION value=""" & obj_AdminList_Rs("Admin_Name") & """>����Ա�ʺţ�" & obj_AdminList_Rs("Admin_Name") & "������Ա������" & obj_AdminList_Rs("Admin_Real_Name") & "</OPTION>"
				End if
				obj_AdminList_Rs.Movenext
			Loop
			obj_AdminList_Rs.Close
			Set obj_AdminList_Rs = Nothing
			%>
      </SELECT></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">ר�Ᵽ��·����</div></td>
      <td class="hback"><input name="SavePath" type="text" id="SavePath"  value="<%=str_SavePath%>" readonly>
        <INPUT type="button"  name="Submit4" value="ѡ��·��" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPathFrame.asp?CurrPath=<%=Replace(sRootDir&"/"&str_newsDir,"//","/")%>',300,250,window,document.MainForm.SavePath);document.MainForm.SavePath.focus();"> 
        <span class="tx"> *</span></td>
    </tr>
	<tr> 
      <td class="hback"><div align="right">ר��ҳ����ģʽ��</div></td>
      <td class="hback">
	  	<select name="SaveType" id="SaveType">
			<option value="0" <% If Int_SaveType = 0 Then Response.Write "selected" %>>/ר��Ӣ����/index.html</option>
			<option value="1" <% If Int_SaveType = 1 Then Response.Write "selected" %>>/ר��Ӣ����/ר��Ӣ����.html</option>
			<option value="2" <% If Int_SaveType = 2 Then Response.Write "selected" %>>/ר��Ӣ����.html</option>
			<option value="3" <% If Int_SaveType = 3 Then Response.Write "selected" %>>/Special_ר��Ӣ����.html</option>
		</select>	
     </td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">ר��ģ���ַ��</div></td>
      <td class="hback"><input name="Templet" type="text" id="Templet" value="<% = str_Templet %>" size="50" maxlength="250" readonly> 
        <input type="button" name="Submit" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../Commpages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.MainForm.Templet);document.MainForm.Templet.focus();"> 
        <span class="tx"> *250���ַ�</span></td>
    </tr>
    <tr> 
      <td class="hback"><div align="right">ר����չ����</div></td>
      <td class="hback">
	    <select name="ExtName" id="ExtName">
          <option value="html" <% if  Trim(str_ExtName) = "html"  then response.Write("selected")%>>.html</option>
          <option value="htm" <% if  Trim(str_ExtName) = "htm"  then response.Write("selected")%>>.htm</option>
          <option value="shtml" <% if  Trim(str_ExtName) = "shtml"  then response.Write("selected")%>>.shtml</option>
          <option value="shtm" <% if  Trim(str_ExtName)= "shtm"  then response.Write("selected")%>>.shtm</option>
          <option value="asp" <% if  Trim(str_ExtName) = "asp"  then response.Write("selected")%>>.asp</option>
        </select> <span class="tx"> *�����Ҫ�Ķ�Ȩ�ޣ���������Ϊ.asp</span></td>
    </tr>
    <tr> 
      <td height="22" class="hback"><div align="right">�Ƿ�������</div></td>
      <td height="22" class="hback"><input name="isLock" type="checkbox" id="isLock" value="1" <% if bit_isLock = 1 then response.Write("checked") %>></td>
    </tr>
    <tr> 
      <td width="18%" height="21" class="hback"><div align="right">������ڣ�</div></td>
      <td width="82%" height="21" class="hback"><input  name="Addtime" type="text" id="Addtime" value="<% = dtm_Addtime %>" readonly>
      <input name="SelectDate" type="button" id="SelectDate" value="ѡ��ʱ��" onClick="OpenWindowAndSetValue('../CommPages/SelectDate.asp',300,130,window,document.all.Addtime);" ></td>
    </tr>
    <tr id="InUrl6" style="display:none"> 
      <td class="hback"><div align="right">�����Ա�飺</div></td>
      <td class="hback"> <input name="BrowPop"  id="BrowPop" type="text" value="<% = lng_GroupID %>" onMouseOver="this.title=this.value;" readonly> 
        <select name="selectPop" id="selectPop" style="overflow:hidden;" onChange="ChooseExeName();">
          <option value="" selected>ѡ���Ա��</option>
          <option value="del" style="color:red;">���</option>
          <% = MF_GetUserGroupID %>
        </select>
        ��Ҫ���� 
        <input name="sPoint" type="text" id="sPoint" size="8" maxlength="5" value="<% = int_sPoint %>"  onChange="ChooseExeName();"  onKeyUp="if(isNaN(value)||event.keyCode==32)execCommand('undo')"  onafterpaste="if(isNaN(value)||event.keyCode==32)execCommand('undo')">
       <span class="tx"> ����0</span> </td>
    </tr>
      <td height="21" class="hback"></td>
      <td height="21" class="hback"><input type="button" name="Submit4222" value="����ר��" onClick="{if(confirm('ȷ�ϱ�������ר����Ϣ��?')){this.document.MainForm.submit();return true;}return false;}"> 
        <input type="reset" name="Submit5222" value="����">
		<input name="SpecialID" type="hidden" id="SpecialID" value="<% = Request.QueryString("SpecialID")%>">
        <input name="Action" type="hidden" id="Action" value="<% = Request.QueryString("Action")%>"></td>
    </tr>
</table>
</form>
</body>

</html>
<%
set Fs_news = nothing
%>
<SCRIPT language="JavaScript">
var DocumentReadyTF=false;
function SetClassEName(Str,Obj)
{
	Obj.value=ConvertToLetters(Str,1);
}
function document.onreadystatechange()
{
	ChooseExeName();
}
function ChooseExeName()
{
  var ObjValue = document.MainForm.selectPop.options[document.MainForm.selectPop.selectedIndex].value;
  if (ObjValue!='')
  {
	if (document.MainForm.BrowPop.value=='')
		document.MainForm.BrowPop.value = ObjValue;
	else if(document.MainForm.BrowPop.value.indexOf(ObjValue)==-1)
		document.MainForm.BrowPop.value = document.MainForm.BrowPop.value+","+ObjValue;
	if (ObjValue=='del')
  		document.MainForm.BrowPop.value ='';
  }
  CheckNumber(document.MainForm.sPoint,"����۵�ֵ");
  if (document.MainForm.sPoint.value>32767||document.MainForm.sPoint.value<-32768)  //||document.MainForm.sPoint.value=='0'
	{
		alert('����۵�ֵ��������Χ��\n���32767���Ҳ���Ϊ0');
		document.MainForm.sPoint.value='';
		document.MainForm.sPoint.focus();
	}
   if (document.MainForm.BrowPop.value!='' || (document.MainForm.sPoint.value!='0' && document.MainForm.sPoint.value!='') ){document.MainForm.ExtName.options[4].selected=true;document.MainForm.ExtName.readonly=true;}
  else {document.MainForm.ExtName.readonly=false;}
}

function CheckExtName(Obj)
{
	if (Obj.value!='')
	{
		for (var i=0;i<document.all.ExtName.length;i++)
		{
			if (document.all.ExtName.options(i).value=='asp') document.all.ExtName.options(i).selected=true;
		}
		document.all.ExtName.readonly=true;
	}
	else
	{
		document.all.ExtName.readonly=false;
	}
}
</SCRIPT>

<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. --> 






