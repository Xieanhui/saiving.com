<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/ns_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
'FileNameRuleArray:�����ļ���̬�ļ����ɹ���,���������ʽ�����������ֿ�
'IndexPageArray:��ҳ�ļ�������չ��,���������ʽ�����������ֿ�,1:��ҳ����2��չ��
'RefreshFileArray:ϵͳ���ٷ����Զ�ˢ����ҳ,ϵͳ���ٷ���ˢ�·���,���������ʽ�������������ֿ�
Dim Conn,Fs_News,FileNameRuleArray,sRootDir,str_CurrPath
MF_Default_Conn
MF_Session_TF
if not MF_Check_Pop_TF("NS_Param") then Err_Show
if not MF_Check_Pop_TF("NS049") then Err_Show
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""

Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if

Set Fs_News=New Cls_news
Fs_News.GetSysParam()
if err.number>0 then
	Response.Redirect("lib/error.asp?ErrCodes=<li>���ȳ��Ի����ݿ�</li>")
	Response.End()
end if
FileNameRuleArray=split(Fs_News.fileNameRule,"$")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {font-weight: bold}
.style2 {color: #FF0000}
-->
</style>
</HEAD>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY>
<script language="JavaScript" src="js/CheckJs.js" type="text/JavaScript"></script>
<script language="JavaScript" src="js/Public.js" type="text/javascript"></script>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="1"></td>
  </tr>
</table> 
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <form action="SetSysParaAction.asp?Act=SetSysPara_Action" method="post" name="SysParaForm" id="SysParaForm">
    <tr> 
      <td align="left" colspan="2" class="xingmu"><a href="#" onMouseOver="this.T_BGCOLOR='#404040';this.T_FONTCOLOR='#FFFFFF';return escape('<div align=\'center\'>FoosunCMS5.0<br> Code by ���ϱ� <BR>Copyright (c) 2006 Foosun Inc</div>')" class="sd"><strong><% = Fs_news.allInfotitle %>ϵͳ��������</strong></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="../../help?Lable=News_Manage" target="_blank" style="cursor:help;'" class="sd">����</a></td>
    </tr>
    <tr class="hback"> 
      <td width="17%" align="right"><% = Fs_news.allInfotitle %>ϵͳվ����⣺</td>
      <td width="83%"><input name="txt_SiteName" type="text" id="txt_SiteName" size="50" value="<%=Fs_News.siteName%>"> 
        <span class="style2">*</span><span id="span_SiteName_Alert"></span></td>
    </tr>
    <tr class="hback"> 
      <td align="right">վ��ؼ��֣� </td>
      <td><textarea name="txt_KeyWords"  style="width:80%" rows="5" id="txt_KeyWords" onKeyUp="ReplaceDot('txt_KeyWords')"><%=Fs_News.keywords%></textarea></td>
    </tr>
    <tr class="hback"> 
      <td align="right"> <% = Fs_news.allInfotitle %>ϵͳǰ̨Ŀ¼��</td>
      <td><input name="txt_NewsDir" type="text" id="txt_NewsDir" value="<%=Fs_News.newsDir%>" size="50"> 
        <span class="style2">*</span><span id="span_NewsDir_Alert"></span>����ڸ�Ŀ¼����ʹ�á�/��</td>
    </tr>
    <tr class="hback"> 
      <td align="right">���ö���������</td>
      <td>
<input name="rad_IsDomain" type="text" id="rad_IsDomain" value="<%=Fs_News.isDomain%>" size="50">
        <br>
        ��ʽ��News.foosun.cn������&quot;http://&quot;��������Ŀ¼�����治��&quot;/&quot;.��������������������ձ���Ϊ��</td>
    </tr>
    <tr class="hback"> 
      <td align="right"> <% = Fs_news.allInfotitle %>\Ŀ¼ɾ�����ã�</td>
      <td align="left"> <input type="radio" name="rad_ReycleTF" value="1" <%if Fs_News.reycleTf=1 then Response.Write("checked")%>>
        ת�����վ 
        <input type="radio" name="rad_ReycleTF" value="0" <%if Fs_News.reycleTf=0 then Response.Write("checked")%>>
        ����ɾ�� </td>
    </tr>
    <tr class="hback"> 
      <td align="right"><% = Fs_news.allInfotitle %>�ļ���ǰ׺��</td>
      <td> <input name="txt_FileNameRule_Element_Prefix" type="text" id="txt_FileNameRule_Element_Prefix" size="50" value="<%=FileNameRuleArray(0)%>"></td>
    </tr>
    <tr class="hback"> 
      <td align="right"><% = Fs_news.allInfotitle %>�ļ���������</td>
      <td> <input name="chk_FileNameRule_Element" type="checkbox" id="chk_FileNameRule_Element" value="Y" <%if instr(FileNameRuleArray(1),"Y")>0 then Response.Write("checked")%>>
        �� 
        <input name="chk_FileNameRule_Element" type="checkbox" id="chk_FileNameRule_Element" value="M" <%if instr(FileNameRuleArray(1),"M")>0 then Response.Write("checked")%>>
        �� 
        <input name="chk_FileNameRule_Element" type="checkbox" id="chk_FileNameRule_Element" value="D" <%if instr(FileNameRuleArray(1),"D")>0 then Response.Write("checked")%>>
        �� 
        <input name="chk_FileNameRule_Element" type="checkbox" id="chk_FileNameRule_Element" value="H" <%if instr(FileNameRuleArray(1),"H")>0 then Response.Write("checked")%>>
        ʱ 
        <input name="chk_FileNameRule_Element" type="checkbox" id="chk_FileNameRule_Element" value="I" <%if instr(FileNameRuleArray(1),"I")>0 then Response.Write("checked")%>>
        �� 
        <input name="chk_FileNameRule_Element" type="checkbox" id="chk_FileNameRule_Element" value="S" <%if instr(FileNameRuleArray(1),"S")>0 then Response.Write("checked")%>>
        �� <br> <input type="radio" name="rad_FileNameRule_Rnd" id="rad_FileNameRule_Rnd" value="2" <%if FileNameRuleArray(2)="2" then Response.Write("checked")%>>
        2λ����� 
        <input type="radio" name="rad_FileNameRule_Rnd" id="rad_FileNameRule_Rnd" value="3" <%if FileNameRuleArray(2)="3" then Response.Write("checked")%>>
        3λ����� 
        <input type="radio" name="rad_FileNameRule_Rnd" id="rad_FileNameRule_Rnd" value="4" <%if FileNameRuleArray(2)="4" then Response.Write("checked")%>>
        4λ����� 
        <input type="radio" name="rad_FileNameRule_Rnd" id="rad_FileNameRule_Rnd" value="5" <%if FileNameRuleArray(2)="5" then Response.Write("checked")%>>
        5λ����� 
        <input name="chk_FileNameRule_UseWord" type="checkbox" id="chk_FileNameRule_UseWord" value="1" <%if ubound(FileNameRuleArray)>=3 then if FileNameRuleArray(3)="1" then Response.Write("checked")%>>
        �Ƿ������ĸ </td>
    </tr>
    <tr class="hback"> 
      <td align="right">�ָ���ţ�</td>
      <td> <input name="txt_FileNameRule_Element_Separator" type="text" id="txt_FileNameRule_Element_Separator" size="50" value="<%=FileNameRuleArray(4)%>"></td>
    </tr>
    <tr class="hback"> 
      <td align="right">�Ƿ�ʹ��<% = Fs_news.allInfotitle %>ID��</td>
      <td> <input type="radio" name="rad_FileNameRule_UseNewsID" value="1" <%if ubound(FileNameRuleArray)>=5 then if FileNameRuleArray(5)="1" then Response.Write("checked")%> onClick="clearAll('rad_FileNameRule_Rnd','chk_FileNameRule_UseWord')">
        �� 
        <input type="radio" name="rad_FileNameRule_UseNewsID" value="0" <%if Ubound(FileNameRuleArray)>=5 then if FileNameRuleArray(5)="0" then Response.Write("checked")%> onClick="checkIt('rad_FileNameRule_Rnd','chk_FileNameRule_UseWord')">
        �� </td>
    </tr>
    <tr class="hback"> 
      <td align="right">�Ƿ�ʹ��NewsID</td>
      <td> <input type="radio" name="rad_FileNameRule_NewsID" value="1" <%if ubound(FileNameRuleArray)>=6 then if FileNameRuleArray(6)="1" then Response.Write("checked")%> onClick="clearAll('rad_FileNameRule_Rnd','chk_FileNameRule_UseWord')">
        �� 
        <input type="radio" name="rad_FileNameRule_NewsID" value="0" <%if Ubound(FileNameRuleArray)>=6 then if FileNameRuleArray(6)="0" then Response.Write("checked")%> onClick="checkIt('rad_FileNameRule_Rnd','chk_FileNameRule_UseWord')">
        �� </td>
    </tr>
    <tr class="hback"> 
      <td align="right">Ŀ¼���ɹ���</td>
      <td> 
	    <input name="rad_FileDirRule" type="radio" value="0" onClick="show_FileDirRule_Detail(this.value)" <%if Fs_News.fileDirRule=0 then Response.Write("checked")%>>
        ����1 
        <input type="radio" name="rad_FileDirRule" value="1" onClick="show_FileDirRule_Detail(this.value)" <%if Fs_News.fileDirRule=1 then Response.Write("checked")%>>
        ����2 
        <input type="radio" name="rad_FileDirRule" value="2" onClick="show_FileDirRule_Detail(this.value)" <%if Fs_News.fileDirRule=2 then Response.Write("checked")%>>
        ����3 
        <input type="radio" name="rad_FileDirRule" value="3" onClick="show_FileDirRule_Detail(this.value)" <%if Fs_News.fileDirRule=3 then Response.Write("checked")%>>
        ����4 
        <input type="radio" name="rad_FileDirRule" value="4" onClick="show_FileDirRule_Detail(this.value)" <%if Fs_News.fileDirRule=4 then Response.Write("checked")%>>
        ����5 
        <input type="radio" name="rad_FileDirRule" value="5" onClick="show_FileDirRule_Detail(this.value)" <%if Fs_News.fileDirRule=5 then Response.Write("checked")%>>
        ����6 
        <input type="radio" name="rad_FileDirRule" value="6" onClick="show_FileDirRule_Detail(this.value)" <%if Fs_News.fileDirRule=6 then Response.Write("checked")%>>
        ����7
		<input type="radio" name="rad_FileDirRule" value="7" onClick="show_FileDirRule_Detail(this.value)" <%if Fs_News.fileDirRule=7 then Response.Write("checked")%>>
		����8
		<br /><span id="span_FileDirRule" style="color:blue"></span></td>
   
    </tr>
    <tr class="hback"> 
      <td align="right">��ҳ���ɹ���</td>
      <td> <input name="rad_ClassSaveType" type="radio" value="0" onClick="show_ClassSaveType_Detail(this.value)" <%if Fs_News.classSaveType=0 then Response.Write("checked")%>>
        ����1 
        <input type="radio" name="rad_ClassSaveType" value="1" onClick="show_ClassSaveType_Detail(this.value)" <%if Fs_News.classSaveType=1 then Response.Write("checked")%>>
        ����2 
        <input type="radio" name="rad_ClassSaveType" value="2" onClick="show_ClassSaveType_Detail(this.value)" <%if Fs_News.classSaveType=2 then Response.Write("checked")%>>
        ����3 &nbsp;&nbsp;<span id="span_ClassSaveType" style="color:blue"></span>      </td>
    </tr>
    <tr class="hback"> 
      <td align="right">��ҳ�ļ�����</td>
      <td><input name="txt_IndexPage_Name" type="text" id="txt_IndexPage_Name" size="50" value="<%=Fs_News.indexPage%>"> 
        <span class="style2">*</span><span id="span_IndexPage_Name_Alert"></span></td>
    </tr>
    <tr class="hback"> 
      <td align="right"><% = Fs_news.allInfotitle %>վ�ļ���չ����</td>
      <td>
        <select name="rad_FileExtName" id="rad_FileExtName">
          <option value="0" <%if Fs_News.fileExtName=0 then Response.Write("selected")%>>html</option>
          <option value="1" <%if Fs_News.fileExtName=1 then Response.Write("selected")%>>htm</option>
          <option value="2" <%if Fs_News.fileExtName=2 then Response.Write("selected")%>>shtml</option>
          <option value="3" <%if Fs_News.fileExtName=3 then Response.Write("selected")%>>shtm</option>
          <option value="4" <%if Fs_News.fileExtName=4 then Response.Write("selected")%>>asp</option>
        </select>        </td>
    </tr>
    <tr class="hback"> 
      <td align="right">����<% = Fs_news.allInfotitle %>������</td>
      <td> <input name="rad_isOpen" type="radio" value="1" <%if Fs_News.isOpen=1 then Response.Write("checked")%>>
        �� 
        <input type="radio" name="rad_isOpen" value="0" <%if Fs_News.isOpen=0 then Response.Write("checked")%>>
        �� </td>
    </tr>
    <tr class="hback"> 
      <td align="right">��ҳģ���ַ��</td>
      <td><input name="txt_IndexTemplet" type="text" id="txt_IndexTemplet" size="50" value="<%=Fs_News.indexTemplet%>"> 
        <input name="bnt_NewsTemplet" type="button" id="bnt_NewsTemplet" value="ѡ��ģ��"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.SysParaForm.txt_IndexTemplet);document.SysParaForm.txt_IndexTemplet.focus();"> 
        <span class="style2">*</span><span id="span_IndexTemplet_Alert"></span></td>
    </tr>
    <tr class="hback"> 
      <td align="right">����·����</td>
      <td> <input type="radio" name="rad_LinkType" value="1" <%if Fs_News.linkType=1 then Response.Write("checked")%>>
        ����·�� 
        <input name="rad_LinkType" type="radio" value="0" <%if Fs_News.linkType=0 then Response.Write("checked")%>>
        ���·�� </td>
    </tr>
    <tr class="hback"> 
      <td align="right">���<% = Fs_news.allInfotitle %>�Ƿ���ˣ�</td>
      <td> <input type="radio" name="rad_isCheck" value="1" <%if Fs_News.isCheck=1 then Response.Write("checked")%>>
        �� 
        <input name="rad_isCheck" type="radio" value="0" <%if Fs_News.isCheck=0 then Response.Write("checked")%>>
        �� </td>
    </tr>
    <tr class="hback" style="display:none;"> 
      <td align="right"><% = Fs_news.allInfotitle %>�����Ƿ���ˣ�</td>
      <td> <input type="radio" name="rad_isReviewCheck" value="1" <%if Fs_News.isReviewCheck=1 then Response.Write("checked")%>>
        �� 
        <input name="rad_isReviewCheck" type="radio" value="0" <%if Fs_News.isReviewCheck=0 then Response.Write("checked")%>>
        �� </td>
    </tr>
    <tr class="hback"> 
      <td align="right">�Ƿ����Ͷ�壺</td>
      <td> <input name="rad_isConstrCheck" type="radio" value="1" <%if Fs_News.isConstrCheck=1 then Response.Write("checked")%>>
        �� 
        <input type="radio" name="rad_isConstrCheck" value="0" <%if Fs_News.isConstrCheck=0 then Response.Write("checked")%>>
        �� </td>
    </tr>
	<tr class="hback"> 
      <td align="right">����Ͷ���ļ���</td>
      <td> <input name="ISCopyFilesTF" type="radio" value="1" <%if Fs_News.CopyFileTF=1 then Response.Write("checked")%>>
        �� 
        <input type="radio" name="ISCopyFilesTF" value="0" <%if Fs_News.CopyFileTF=0 then Response.Write("checked")%>>
        �� 
		<span class="tx" style="margin-left:10px;">[ ���Ͷ��֮�󣬸����ʹ�õ�ͼƬ���ļ��Ƿ��Ƶ�ϵͳ�ļ�Ŀ¼ ]</span></td>
    </tr>
	<tr class="hback"> 
      <td align="right">�����޸�Ͷ�壺</td>
      <td> <input name="EditFileTF" type="radio" value="1" <%if Fs_News.EditFilesTF=1 then Response.Write("checked")%>>
        �� 
        <input type="radio" name="EditFileTF" value="0" <%if Fs_News.EditFilesTF=0 then Response.Write("checked")%>>
        �� 
		<span class="tx" style="margin-left:10px;">[ ���Ͷ��֮���Ƿ����������ٴ��޸Ĵ�Ͷ�岢���¸������Ͷ�� ]</span></td>
    </tr>
    <tr class="hback"> 
      <td align="right"><% = Fs_news.allInfotitle %>��ӷ�ʽ��</td>
      <td> <input type="radio" name="rad_AddNewsType" value="1" <%if Fs_News.addNewsType=1 then Response.Write("checked")%>>
        �߼���ʽ 
        <input name="rad_AddNewsType" type="radio" value="0" <%if Fs_News.addNewsType=0 then Response.Write("checked")%>>
        ��෽ʽ </td>
    </tr>
    <tr class="hback" style="display:none"> 
      <td align="right">ͳ���滻�ַ���</td>
      <td><input name="txt_AllInfotitle" type="text" id="txt_AllInfotitle" value="<%=Fs_News.allInfotitle%>" size="50" ></td>
    </tr>
	
    <tr class="hback"> 
      <td align="right">�Ƿ����ڲ����ӣ�</td>
      <td> <input name="InsideLink" type="radio" value="1" <%if Fs_News.InsideLink=1 then Response.Write("checked")%>>
        �� 
        <input type="radio" name="InsideLink" value="0" <%if Fs_News.InsideLink=0 then Response.Write("checked")%>>
        �� </td>
    </tr>
    <tr class="hback"> 
      <td colspan="2" align="right" class="xingmu"><div align="left">RSS�ۺ�</div></td>
    </tr>
    <tr class="hback"> 
      <td align="right">����RSS</td>
      <td><input name="RSSTF" type="checkbox" id="RSSTF" value="1" <%if Fs_News.RSSTF=1 then response.Write("checked")%>></td>
    </tr>
    <tr class="hback"> 
      <td align="right"><div align="right">��ʾ���¾ۺ϶�����</div></td>
      <td><input name="rssNumber" type="text" id="rssNumber" value="<% = Fs_News.rssNumber%>" onChange="if(/\D/.test(this.value)){alert('ֻ����������');this.value='';}"></td>
    </tr>
    <tr class="hback"> 
      <td align="right"><div align="right">վ��RSS����</div></td>
      <td><textarea name="rssdescript"  style="width:80%" rows="5" id="rssdescript" ><% = Fs_News.rssdescript%></textarea></td>
    </tr>
    <tr class="hback"> 
      <td align="right"><div align="right">����ͼƬ��ַ</div></td>
      <td><input name="RSSPIC" type="text" id="RSSPIC" size="40" value="<% = Fs_News.RSSPIC%>"> <input type="button" name="PPPChoose"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.SysParaForm.RSSPIC);"></td>
    </tr>
    <tr class="hback"> 
      <td align="right"><div align="right">��������</div></td>
      <td><input name="rssContentNumber" type="text" id="rssContentNumber" value="<% = Fs_News.rssContentNumber%>" onChange="if(/\D/.test(this.value)){alert('ֻ����������');this.value='';}"></td>
    </tr>
    <tr class="hback"> 
      <td align="right">&nbsp;</td>
      <td><input type="Button" name="btn_SetSysParam" value=" ���� " onClick="SetSysParam()"/> 
        <input type="reset" name="sub_rest" value=" ���� " /></td>
    </tr>
  </form>
</table>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script> 
</body>
<script language="javascript">
function SetSysParam()
{
	var flag1=isEmpty("txt_SiteName","span_SiteName_Alert");
	var flag2=isEmpty("txt_NewsDir","span_NewsDir_Alert");
	var flag3=isEmpty("txt_IndexPage_Name","span_IndexPage_Name_Alert");
	var flag4=isEmpty("txt_IndexTemplet","span_IndexTemplet_Alert");
	if(flag1&&flag2&&flag3&&flag4)
	{
		document.SysParaForm.submit();
	}
}


//��ʾ��ӦĿ¼���ɹ���ĸ�ʽ
function show_FileDirRule_Detail(param)
{
	if(isNaN(param))
	{
		return;
	}
	switch(parseInt(param))
	{
		case 0:document.getElementById("span_FileDirRule").innerHTML="��ʽ��[ ��ĿӢ��/2006-6-9 ]";break
		case 1:document.getElementById("span_FileDirRule").innerHTML="��ʽ��[ ��ĿӢ��/2006/6/9/ ]";break
		case 2:document.getElementById("span_FileDirRule").innerHTML="��ʽ��[ ��ĿӢ��/2006/6-9/ ]";break
		case 3:document.getElementById("span_FileDirRule").innerHTML="��ʽ��[ ��ĿӢ��/2006-6/9/ ]";break
		case 4:document.getElementById("span_FileDirRule").innerHTML="��ʽ��[ ��ĿӢ��/�ļ��� ]";break
		case 5:document.getElementById("span_FileDirRule").innerHTML="��ʽ��[ ��ĿӢ��/2006/6/ ]";break
		case 6:document.getElementById("span_FileDirRule").innerHTML="��ʽ��[ ��ĿӢ��/2006/69/ ]";break
  		case 7:document.getElementById("span_FileDirRule").innerHTML="��ʽ��[ ��ĿӢ��/200669/ ]";break
	}
}
show_FileDirRule_Detail(<%=Fs_News.fileDirRule%>)

function show_ClassSaveType_Detail(param)
{
	if(isNaN(param))
		return;
	switch(parseInt(param))
	{
		case 0:document.getElementById("span_ClassSaveType").innerHTML="��ʽ��[ ��ĿӢ��/index.html ]";break
		case 1:document.getElementById("span_ClassSaveType").innerHTML="��ʽ��[ ��ĿӢ��/��ĿӢ��.html ]";break
		case 2:document.getElementById("span_ClassSaveType").innerHTML="��ʽ��[ ��ĿӢ��.html ]";break
	}
}
show_ClassSaveType_Detail(<%=Fs_News.classSaveType%>)

function clearAll(radio,check)
{
	var RadioArray=document.all(radio);
	for(var i=0;i<RadioArray.length;i++)
	{
		RadioArray[i].checked=false;
	}
	document.all(check).checked=false;
}

function checkIt(radio,check)
{
	var RadioArray=document.all(radio);
	var checkedTF=false;
	for(var i=0;i<RadioArray.length;i++)
	{
		if("<%=FileNameRuleArray(2)%>"==(2+i).toString())
		RadioArray[i].checked=true;
	}
	if("<%=FileNameRuleArray(3)%>"=="1")
		document.all(check).checked=true;
	for(var i=0;i<RadioArray.length;i++)
	{
		if(RadioArray[i].checked)
		{
			checkedTF=true;
		}
	}
	if(!checkedTF)RadioArray[2].checked=true;
}
checkIt('rad_FileNameRule_Rnd','chk_FileNameRule_UseWord')

<%
	Conn.close
	set Conn=nothing
	Set Fs_News=nothing
%>
</script>
</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





