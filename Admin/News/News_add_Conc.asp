<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<%
	Dim Conn,User_Conn,sRootDir,tmp_sFileExtName,tmp_sTemplets,tmp_defineid
	Dim Fs_news,str_ClassID,news_SQL,obj_news_rs,icNum,isUrlStr,str_Href,str_Href_title,obj_news_rs_1,str_action,obj_cnews_rs,news_count,str_CurrPath
	Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
	Temp_Admin_Name = Session("Admin_Name")
	Temp_Admin_Is_Super = Session("Admin_Is_Super")
	Temp_Admin_FilesTF = Session("Admin_FilesTF")
	MF_Default_Conn
	MF_User_Conn
	MF_GetUserGroupID
	'session�ж�
	MF_Session_TF 
	'Ȩ���ж�
	set Fs_news = new Cls_News
	Fs_News.GetSysParam()
	if Fs_News.isOpen=0 then
		Response.Redirect("lib/error.asp?ErrCodes=<li>���ŷ������ܹر�!</li>&ErrorURL=")
		Response.End()
	end if
	If Not Fs_news.IsSelfRefer Then response.write "�Ƿ��ύ����":Response.end
	str_ClassID = NoSqlHack(Request.QueryString("ClassID"))
	if not Get_SubPop_TF(str_ClassID,"NS001","NS","news") then Err_Show
	if Trim(str_ClassID)<>"" then
		Dim tmp_class_obj
		set tmp_class_obj = conn.execute("select FileExtName,NewsTemplet,DefineID from FS_NS_NewsClass where classID='"& str_ClassID &"'")
		if not tmp_class_obj.eof then
			Select Case tmp_class_obj(0)
					Case "html"
						tmp_sFileExtName = 0
					Case "htm"
						tmp_sFileExtName =1
					Case "shtml"
						tmp_sFileExtName = 2
					Case "shtm"
						tmp_sFileExtName = 3
					Case "asp"
						tmp_sFileExtName = 4
			End Select
			tmp_sTemplets = tmp_class_obj(1)
			tmp_defineid = tmp_class_obj(2)
			tmp_class_obj.close:set tmp_class_obj=nothing
		Else
			tmp_defineid = 0
			tmp_sFileExtName = fs_news.fileExtName
			tmp_sTemplets = Replace("/"& G_TEMPLETS_DIR &"/NewsClass/news.htm","//","/")
		end if
	Else
		tmp_sFileExtName = fs_news.fileExtName
		tmp_sTemplets = Replace("/"& G_TEMPLETS_DIR &"/NewsClass/news.htm","//","/")
	End if
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
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<link rel="Stylesheet" href="../../word_edit/images/word_edit.css" type="text/css" media="all" />
</head>
<body>
<script language="JavaScript" src="js/Public.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/CheckJs.js" type="text/JavaScript"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<iframe width="260" height="165" id="colorPalette" src="lib/selcolor.htm" style="visibility:hidden; position: absolute;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="1"></td>
	</tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
	<form action="News_Save.asp" name="NewsForm" method="post">
		<tr class="xingmu">
			<td colspan="4" class="xingmu">���
				<% = Fs_news.allInfotitle %>
				<a href="../../help?Lable=NS_News_add" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr class="hback">
			<td class="hback">
				<div align="right">
					<table width="95" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td height="1"></td>
						</tr>
					</table>
					<% = Fs_news.allInfotitle %>
					���� </div>
			</td>
			<td colspan="3" class="hback">
				<input name=NewsType type=radio id="NewsType" onClick="SwitchNewsType('TextNews');" value="TextNews" checked>
				��ͨ
				<input name=NewsType type=radio id="NewsType" onClick="SwitchNewsType('PicNews');" value="PicNews">
				ͼƬ
				<input name=NewsType type=radio id="NewsType" onClick="SwitchNewsType('TitleNews');" value="TitleNews">
				���� ��&nbsp;&nbsp;����
				<input name="isdraft" type="checkbox" id="isdraft" value="1">
				�浽�ݸ�����</td>
		</tr>
		<tr class="hback">
			<td width="12%" class="hback">
				<div align="right">
					<% = Fs_news.allInfotitle %>
					����</div>
			</td>
			<td colspan="3" class="hback">
				<input name="NewsTitle" type="text" id="NewsTitle" size="40" maxlength="255">
				<input name="TitleColor" id="TitleColor" type="hidden">
				<img src="images/rectNoColor.gif" width="18" height="17" border=0 align="absmiddle" id="TitleColorShow" style="cursor:pointer;background-color:;" title="ѡȡ��ɫ!" onClick="GetColor(document.getElementById('TitleColorShow'),'TitleColor');">
				<input name="titleBorder" type="checkbox" id="titleBorder" value="1">
				����
				<input name="TitleItalic" type="checkbox" id="TitleItalic" value="1">
				б��
				<input name="isShowReview" type="checkbox" id="isShowReview" value="1">
				�������� ��Ȩ��
				<select name="PopID" id="PopID">
					<option value="5">���ö�</option>
					<option value="4">��Ŀ�ö�</option>
					<option value="0" selected>һ��</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback">
				<div align="right">ѡ����Ŀ</div>
			</td>
			<td colspan="3" class="hback">
				<input name="ClassName" type="text" id="ClassName5" style="width:50%" value="<%=Fs_News.GetAdd_ClassName(str_ClassID)%>" readonly>
				<input name="ClassID" type="hidden" id="ClassID" value="<% = str_ClassID %>">
				<input type="button" name="Submit" value="ѡ����Ŀ"   onClick="SelectClass();">
			</td>
		</tr>
		<tr class="hback" style="display:none">
			<td class="hback">
				<div align="right">ѡ��ר��</div>
			</td>
			<td colspan="3" class="hback">
				<input name="SpecialID" type="text" id="SpecialID" style="width:50%" readonly>
				<input name="SpecialID_EName" type="hidden" id="SpecialID_EName">
				<span class="tx"> </span>
				<input type="button" name="Submit" value="ѡ��ר��"   onClick="SelectSpecial();">
				<span class="tx">
				<input name="Submit" type="button" id="Submit" onClick="dospclear();" value="���ר��">
				</span> <span class="tx"> �ɶ�ѡ</span> &nbsp;&nbsp;<a href="../../help?Lable=NS_News_add_special" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr class="hback" id="str_URLAddress" style="display:none;">
			<td class="hback">
				<div align="right">���ӵ�ַ </div>
			</td>
			<td colspan="3" class="hback">
				<input name="URLAddress" type="text" id="URLAddress"  style="width:65%" maxlength="255"><font color="#FF0000">ע:�ⲿ��ַ,�����м���http://</font></td>
		</tr>
		<tr class="hback" id="str_CurtTitle" style="display:none">
			<td class="hback">
				<div align="right"> ������</div>
			</td>
			<td width="38%" class="hback">
				<input name="CurtTitle" type="text" id="CurtTitle" size="40" maxlength="255">
			</td>
			<td width="10%" class="hback">
				<div align="right">�ؼ���</div>
			</td>
			<td width="40%" class="hback">
				<input name="KeywordText" type="text" id="KeywordText" size="15" maxlength="255">
				<input name="KeyWords" type="hidden" id="KeyWords">
				<select name="selectKeywords" id="selectKeywords" style="width:120px" onChange=Dokesite(this.options[this.selectedIndex].value)>
					<option value="" selected>ѡ��ؼ���</option>
					<option value="Clean" style="color:red">���</option>
					<%=Fs_news.GetKeywordslist("",1)%>
				</select>
				<input name="KeywordSaveTF" type="checkbox" id="KeywordSaveTF" value="1">
				����</td>
		</tr>
		<tr class="hback"  id="str_Templet">
			<td class="hback">
				<div align="right">ģ���ַ</div>
			</td>
			<td colspan="3" class="hback">
				<input name="Templet" type="text" id="Templet" style="width:85%" value="<%=tmp_sTemplets%>" maxlength="255" readonly>
				<input name="Submit5" type="button" id="selNewsTemplet" value="ѡ��ģ��"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.NewsForm.Templet);document.NewsForm.Templet.focus();">
			</td>
		</tr>
		<tr class="hback" id="str_NewsSmallPicFile" style="display:none">
			<td class="hback">
				<div align="right">ͼƬ(С)</div>
			</td>
			<td colspan="3" class="hback">
				<input name="NewsSmallPicFile" type="text" id="NewsSmallPicFile" style="width:85%" maxlength="255">
				<input type="button" name="PPPChoose"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.NewsForm.NewsSmallPicFile);">
			</td>
		</tr>
		<tr class="hback" id="str_NewsPicFile" style="display:none">
			<td class="hback">
				<div align="right">ͼƬ(��)</div>
			</td>
			<td colspan="3" class="hback">
				<input name="NewsPicFile" type="text" id="NewsPicFile" style="width:85%" maxlength="255">
				<input type="button" name="PPPChoose2"  value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.NewsForm.NewsPicFile);">
			</td>
		</tr>
		<tr class="hback" id="str_PicborderCss" style="display:none">
			<td class="hback">
				<div align="right">ͼƬCSS</div>
			</td>
			<td colspan="3" class="hback">
				<select name="PicborderCss" id="PicborderCss">
					<option value="1" selected>ͼƬCSS��ʽһ</option>
					<option value="2">ͼƬCSS��ʽ��</option>
					<option value="3">ͼƬCSS��ʽ��</option>
					<option value="4">ͼƬCSS��ʽ��</option>
					<option value="5">ͼƬCSS��ʽ��</option>
				</select>
				ͼƬǰ̨��ʾCSS��ʽ������ӱ߿򡡡�<a href="../../help?Lable=NS_News_add_PicCSS" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr class="hback" id="str_Author" style="display:none">
			<td class="hback">
				<div align="right">
					<% = Fs_news.allInfotitle %>
					����</div>
			</td>
			<td class="hback">
				<input name="Author" type="text" id="Author" size="15" maxlength="255">
				<select name="selectAuthor" id="selectAuthor" style="width:120px"  onChange="document.NewsForm.Author.value=this.options[this.selectedIndex].text;">
					<option style="color:red"> </option>
					<option value="����">����</option>
					<option value="��վ">��վ</option>
					<option value="δ֪">δ֪</option>
					<%=Fs_news.GetKeywordslist("",3)%>
				</select>
				<input name="AuthorSaveTF" type="checkbox" id="AuthorSaveTF" value="1">
				����</td>
			<td class="hback">
				<div align="right">
					<% = Fs_news.allInfotitle %>
					��Դ</div>
			</td>
			<td class="hback">
				<input name="Source" type="text" id="Source" size="15" maxlength="255">
				<select name="selectSource" id="selectSource" style="width:120px"  onChange="document.NewsForm.Source.value=this.options[this.selectedIndex].text;">
					<option value="" selected> </option>
					<option value="��վԭ��">��վԭ��</option>
					<option value="����">����</option>
					<%=Fs_news.GetKeywordslist("",2)%>
				</select>
				<input name="SourceSaveTF" type="checkbox" id="SourceSaveTF" value="1">
				����</td>
		</tr>
		<tr class="hback" style="display:none">
			<td class="hback">
				<div align="right">
					<% = Fs_news.allInfotitle %>
					����</div>
			</td>
			<td colspan="3" class="hback">
				<div align="left">
					<textarea name="NewsNaviContent" rows="6" id="NewsNaviContent" style="width:96%"></textarea>
				</div>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback">
				<div align="right">����</div>
			</td>
			<td colspan="3" class="hback">
				<div align="left">
					<input name="NewsProperty_Rec" type="checkbox" id="NewsProperty" value="1">
					�Ƽ�
					<input name="NewsProperty_mar" type="checkbox" id="NewsProperty" value="1" checked>
					����
					<input name="NewsProperty_rev" type="checkbox" id="NewsProperty" value="1" checked>
					��������
					<input name="NewsProperty_constr" type="checkbox" id="NewsProperty" value="1">
					Ͷ��
					<input name="NewsProperty_tt" type="checkbox" id="NewsProperty" value="1"  onClick="ChooseTodayNewsType();">
					ͷ��
					<input name="NewsProperty_hots" type="checkbox" id="NewsProperty" value="1" disabled="disabled">
					�ȵ�
					<input name="NewsProperty_jc" type="checkbox" id="NewsProperty" value="1">
					����
					<input name="NewsProperty_unr" type="checkbox" id="NewsProperty" value="1">
					������
					<input name="NewsProperty_ann" type="checkbox" id="NewsProperty" value="1">
					���� <span id="str_filt" style="display:none">
					<input name="NewsProperty_filt" type="checkbox" id="NewsProperty" value="1">
					�õ�</span></div>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" id="TodayNews" style="display:none;">
			<td colspan="4" class="hback">
				<table width="100%" border="0" cellspacing="1" cellpadding="2" class="table">
					<tr>
						<td height="26" align="center" width="120" class="xingmu">ͷ��
							<% = Fs_news.allInfotitle %>
							���ͣ�</td>
						<td height="26" class="hback">
							<input name="TodayNewsPicTF" value="" type="radio" checked onClick="if(this.checked){document.getElementById('TodayPicParam').style.display='none';}">
							����ͷ��
							<input name="TodayNewsPicTF" value="FoosunCMS" type="radio" onClick="if(this.checked){document.getElementById('TodayPicParam').style.display='';}">
							ͼƬͷ�� ����<a href="../../help?Lable=NS_News_add_tt" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
					</tr>
					<tr id="TodayPicParam" style="display:none;">
						<td width="120" height="26" align="center"  class="xingmu">ͷ��
							<% = Fs_news.allInfotitle %>
							������</td>
					  <td height="26" class="hback">&nbsp;&nbsp;����
							<SELECT name="FontFace" id="FontFace">
								<option value="����">����</option>
								<option value="����_GB2312">����</option>
								<option value="������">������</option>
								<option value="����">����</option>
								<option value="����">����</option>
								<OPTION value="Andale Mono">Andale Mono</OPTION>
								<OPTION value="Arial">Arial</OPTION>
								<OPTION value="Arial Black">Arial Black</OPTION>
								<OPTION value="Book Antiqua" >Book Antiqua</OPTION>
								<OPTION value="Century Gothic">Century Gothic</OPTION>
								<OPTION value="Comic Sans MS">Comic Sans MS</OPTION>
								<OPTION value="Courier New">Courier New</OPTION>
								<OPTION value="Georgia">Georgia</OPTION>
								<OPTION value="Impact">Impact</OPTION>
								<OPTION value="Tahoma">Tahoma</OPTION>
								<OPTION value="Times New Roman">Times New Roman</OPTION>
								<OPTION value="Trebuchet MS">Trebuchet MS</OPTION>
								<OPTION value="Script MT Bold">Script MT Bold</OPTION>
								<OPTION value="Stencil">Stencil</OPTION>
								<OPTION value="Verdana">Verdana</OPTION>
								<OPTION value="Lucida Console">Lucida Console</OPTION>
							</SELECT>
							<select name="FontSize">
								<option value="8">8px</option>
								<option value="9">9px</option>
								<option value="10">10px</option>
								<option value="12">12px</option>
								<option value="18">18px</option>
								<option value="20">20px</option>
								<option value="24">24px</option>
								<option value="28">28px</option>
								<option value="30">30px</option>
								<option value="32">32px</option>
								<option value="36">36px</option>
								<option value="40">40px</option>
								<option value="48">48px</option>
								<option value="54">54px</option>
								<option value="60">60px</option>
								<option value="72">72px</option>
							</select>
							<input type="text" name="FontColor" maxlength=7 size=7 id="FontColor" value="000000">
							<img src="images/rect.gif" width="18" height="17" border=0 align=absmiddle id="FontColorShow" style="cursor:pointer;background-Color:#000000;" title="ѡȡ��ɫ!" onClick="GetColor(this,'FontColor');"> ���� �����ࣺ
							<INPUT TYPE="text" maxlength="3" NAME="FontSpace" size=3 value="12">
							px ͼƬ����ɫ
							<input type="text" name="FontBgColor" maxlength=7 size=7 id="FontBgColor" value="FFFFFF">
						  <img src="images/rect.gif" width="18" height="17" border=0 align=absmiddle id="FontBgColorShow" style="cursor:pointer;background-Color:;" title="ѡȡ��ɫ!" onClick="GetColor(this,'FontBgColor');"> <br>
						  &nbsp;&nbsp;ͼƬͷ�����⣺
						  <input name="PicTitle" type="text" id="PicTitle" size="40" maxlength="255">
						  &nbsp;&nbsp;ͼƬ��ȣ�
						  <input name="PicTitlewidth" type="text" id="PicTitlewidth" size="10" maxlength="10">
px </td>
					</tr>
				</table>
			</td>
		</tr>
		<tr  id="str_Content" style="display:" >
			<td height="303" class="hback">
				<div align="left">
					���ݷ�ҳ��ǩ[FS:PAGE]<br>
                    <a href="javascript:void(0);" onClick="InsertHTML('[FS:PAGE]','NewsContent')"><span class="tx">�����ҳ��ǩ</span></a><br>
                    <br>
              <input name="NewsProperty_Remote" type="checkbox" id="NewsProperty_Remote" value="1">
					Զ�̴�ͼ <br>
					<span class="tx">���ô˹��ܺ������������վ�ϸ��Ƶ��ұߵı༭���У������а�����ͼƬ����ϵͳ���ڱ�������ʱ�Զ������ͼƬ���Ƶ���վ�������ϡ�<br>
					ϵͳ����������ͼƬ�Ĵ�С��Ӱ���ٶȣ�����ͼƬ�϶�ʱ��Ҫʹ�ô˹��ܡ�</span> </div>
		  </td>
			<td height="303" colspan="3" class="hback">
				<!--�༭����ʼ-->
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=Content' frameborder=0 scrolling=no width='100%' height='480'></iframe>
				<input type="hidden" name="Content" value="">
                <!--�༭������-->
             </td>
		</tr>
		<tr class="hback" id="str_GroupID" style="display:none">
			<td class="hback">
				<div align="right">�������</div>
			</td>
			<td colspan="3" class="hback">
				<input name="PointNumber" type="text" id="PointNumber2" size="16"  onChange="ChooseExeName();">
				���
				<input name="Money" type="text" id="Money2" size="16"  onChange="ChooseExeName();">
				���Ȩ��
				<input name="BrowPop"  id="BrowPop" type="text" onMouseOver="this.title=this.value;" readonly>
				<select name="selectPop" id="selectPop" style="overflow:hidden;" onChange="ChooseExeName();">
					<option value="" selected>ѡ���Ա��</option>
					<option value="del" style="color:red;">���</option>
					<% = MF_GetUserGroupID %>
				</select>
				<a href="../../help?Lable=NS_News_add_pop" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr class="hback" id="str_FileName" style="display:none">
			<td class="hback">
				<div align="right">�ļ���</div>
			</td>
			<td class="hback">
				<%
	  Dim RoTF
	  if instr(Fs_News.strFileNameRule(Fs_News.fileNameRule,0,0),"�Զ����ID")>0 then:RoTF="Readonly":End if
	  Response.Write"<input name=""FileName"" type=""text"" id=""FileName"" size=""40"" "& RoTF &" maxlength=""255"" value="""&Fs_News.strFileNameRule(Fs_News.fileNameRule,0,0)&""" title=""��������������趨Ϊ�Զ���ţ��������޸�"">"
	  %>
			</td>
			<td class="hback">
				<div align="right">��չ��</div>
			</td>
			<td class="hback">
				<select name="FileExtName" id="FileExtName">
					<option value="html" <%if tmp_sFileExtName = 0 then response.Write("selected")%>>.html</option>
					<option value="htm" <%if tmp_sFileExtName = 1 then response.Write("selected")%>>.htm</option>
					<option value="shtml" <%if tmp_sFileExtName = 2 then response.Write("selected")%>>.shtml</option>
					<option value="shtm" <%if tmp_sFileExtName = 3 then response.Write("selected")%>>.shtm</option>
					<option value="asp" <%if tmp_sFileExtName = 4 then response.Write("selected")%>>.asp</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback">
				<div align="right">�������</div>
			</td>
			<td class="hback">
				<input name="addtime" type="text" id="addtime" value="<% = now %>" size="40" maxlength="255">
			</td>
			<td class="hback">
				<div align="right">�������</div>
			</td>
			<td class="hback">
				<input name="Hits" type="text" id="Hits" value="0" size="20">
			</td>
		</tr>
		<tr class="hback">
			<td class="hback">
				<div align="right"></div>
			</td>
			<td colspan="3" class="hback">
				<input type="button" name="Submit" value="ȷ�ϱ���<% = Fs_news.allInfotitle %>" onClick="SubmitFun(this.form);">
				<input name="d_Id" type="hidden" id="d_Id" value="<%=tmp_defineid%>">
				<input type="reset" name="Submit4" value="��������">
				<input name="News_Action" type="hidden" id="News_Action2" value="add_Save">
			</td>
		</tr>
	</form>
</table>
</body>
</html>
<%
set tmp_class_obj = nothing		
set Fs_news = nothing
%>
<script language="JavaScript" type="text/JavaScript">
function SubmitFun(FormObj)
{
	if(FormObj.NewsTitle.value == "")
	{
		alert("����д���⣡");
		FormObj.NewsTitle.focus();
		return false;
	}
	if(FormObj.ClassName.value=="")
	{
		alert("��ѡ����Ŀ��");
		FormObj.ClassName.focus();
		return false;
	}
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}
	FormObj.Content.value=frames["NewsContent"].GetNewsContentArray();
	FormObj.submit();
}
function SwitchNewsType(NewsType)
{
	switch (NewsType)
	{
	case "TitleNews":
		document.getElementById('str_UrLaddress').style.display='';
		document.getElementById('str_NewsPicFile').style.display='none';
		document.getElementById('str_NewsSmallPicFile').style.display='none';
		document.getElementById('str_Content').style.display='none';
		document.getElementById('str_GroupID').style.display='none';
		document.getElementById('str_filt').style.display='none';
		break;
	case "PicNews":
		document.getElementById('str_UrLaddress').style.display='none';
		document.getElementById('str_NewsPicFile').style.display='';
		document.getElementById('str_NewsSmallPicFile').style.display='';
		document.getElementById('str_Content').style.display='';
		document.getElementById('str_filt').style.display='';
		
		//setChecked('PicNews',true);
		//setChecked('TitleNews',false);
		//setChecked('TextNews',false);
		//ChoosePicType();
		break;
	default :
		document.getElementById('str_UrLaddress').style.display='none';
		document.getElementById('str_NewsPicFile').style.display='none';
		document.getElementById('str_NewsSmallPicFile').style.display='none';
		document.getElementById('str_Content').style.display='';
		document.getElementById('str_filt').style.display='none';
	}
}
function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}
function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}
function setColor(color)
{
	if(ColorImg.id=='FontColorShow' && color=="#") color='#000000';
	if(ColorImg.id=='FontBgColorShow' && color=="#") color='#FFFFFF';
	if (ColorValue){ColorValue.value = color.substr(1);}
	if (ColorImg && color.length>1){
		ColorImg.src='Images/Rect.gif';
		ColorImg.style.backgroundColor = color;
	}else if(color=='#'){ ColorImg.src='Images/rectNoColor.gif';}
	document.getElementById("colorPalette").style.visibility="hidden";
}
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectClassFrame.asp',400,300,window);
	try {
		$("ClassID").value= ReturnValue[0][0];
		$("ClassName").value= ReturnValue[1][0];
		$("Templet").value= ReturnValue[2][0];
	}
	catch (ex) { }
}
function SelectSpecial()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectspecialFrame.asp',400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		if (document.NewsForm.SpecialID.value.search(TempArray[1])==-1)
		{
		if(document.all.SpecialID.value=='') document.all.SpecialID.value=TempArray[1];
		else document.all.SpecialID.value=document.all.SpecialID.value+','+TempArray[1];
		if(document.all.SpecialID_EName.value=='') document.all.SpecialID_EName.value=TempArray[0];
		else document.all.SpecialID_EName.value=document.all.SpecialID_EName.value+','+TempArray[0];
		}
	}
}
function ChooseTodayNewsType()
{
	if (document.NewsForm.NewsProperty_tt.checked==true) document.getElementById('TodayNews').style.display='';
	else document.getElementById('TodayNews').style.display='none';
}
function GetColor(img_val,input_val)
{
	var PaletteLeft,PaletteTop
	var obj = document.getElementById("colorPalette");
	ColorImg = img_val;
	ColorValue = document.getElementById(input_val);
	if (obj){
		PaletteLeft = getOffsetLeft(ColorImg)
		PaletteTop = (getOffsetTop(ColorImg) + ColorImg.offsetHeight)
		if (PaletteLeft+150 > parseInt(document.body.clientWidth)) PaletteLeft = parseInt(event.clientX)-260;
		if (PaletteTop > parseInt(document.body.clientHeight)) PaletteTop = parseInt(document.body.clientHeight)-165;
		obj.style.left = PaletteLeft + "px";
		obj.style.top = PaletteTop + "px";
		if (obj.style.visibility=="hidden")
		{
			obj.style.visibility="visible";
		}else {
			obj.style.visibility="hidden";
		}
	}
}
</script>
<SCRIPT language="JavaScript">
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	ChooseExeName();
}
function ChooseExeName()
{
  var ObjValue = document.NewsForm.selectPop.options[document.NewsForm.selectPop.selectedIndex].value;
  if (ObjValue!='')
  {
	if (document.NewsForm.BrowPop.value=='')
		document.NewsForm.BrowPop.value = ObjValue;
	else if(document.NewsForm.BrowPop.value.indexOf(ObjValue)==-1)
		document.NewsForm.BrowPop.value = document.NewsForm.BrowPop.value+","+ObjValue;
	if (ObjValue=='del')
  		document.NewsForm.BrowPop.value ='';
  }
   CheckNumber(document.NewsForm.PointNumber,"����۵�ֵ");
  if (document.NewsForm.PointNumber.value>32767||document.NewsForm.PointNumber.value<-32768||document.NewsForm.PointNumber.value=='0')
	{
		alert('����۵�ֵ��������Χ��\n���32767���Ҳ���Ϊ0');
		document.NewsForm.PointNumber.value='';
		document.NewsForm.PointNumber.focus();
	}
   CheckNumber(document.NewsForm.Money,"������ֵ");
  if (document.NewsForm.Money.value>32767||document.NewsForm.Money.value<-32768||document.NewsForm.Money.value=='0')
	{
		alert('������ֵ��������Χ��\n���32767���Ҳ���Ϊ0');
		document.NewsForm.Money.value='';
		document.NewsForm.Money.focus();
	}
  if (document.NewsForm.BrowPop.value!=''||document.NewsForm.PointNumber.value!=''||document.NewsForm.Money.value!=''){document.NewsForm.FileExtName.options[4].selected=true;document.NewsForm.FileExtName.readonly=true;}
  else {document.NewsForm.FileExtName.readonly=false;}
}
function CheckFileExtName(Obj)
{
	if (Obj.value!='')
	{
		for (var i=0;i<document.all.FileExtName.length;i++)
		{
			if (document.all.FileExtName.options(i).value=='asp') document.all.FileExtName.options(i).selected=true;
		}
		document.all.FileExtName.readonly=true;
	}
	else
	{
		document.all.FileExtName.readonly=false;
	}
}
</SCRIPT>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






