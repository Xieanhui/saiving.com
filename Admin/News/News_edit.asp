<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<%
Dim Conn,User_Conn,sRootDir,strShowErr
MF_Default_Conn
MF_User_Conn
MF_GetUserGroupID
MF_Session_TF 
Dim Fs_news,obj_newsedit_rs,str_CurrPath,str_display
Dim str_PopId,str_ClassID,str_NewsID,str_SpecialEName,str_NewsTitle,str_CurtTitle,str_NewsNaviContent,str_isShowReview,str_TitleColor,str_titleBorder,str_TitleItalic,str_IsURL,str_URLAddress
Dim str_Content,str_isPicNews,str_NewsPicFile,str_NewsSmallPicFile,str_PicborderCss,str_Templet,str_GroupID,str_PointNumber,str_Money,str_Source,str_Editor,str_Keywords,str_Author
Dim str_Hits,str_FileName,str_FileExtName,str_NewsProperty,str_DefineID,str_TodayNewsPic,str_isLock,str_isRecyle,str_addtime,str_NewsType,str_isdraft
Dim str_UrLaddress_1,str_CurtTitle_1,str_NewsPicFile_1,str_NewsSmallPicFile_1,str_Templet_1,str_Content_1,str_PicborderCss_1,str_Author_1,str_GroupID_1,str_FileName_1,str_filt_1,IsAdPic,AdPicWH,AdPicLink,AdPicAdress
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
set Fs_news = new Cls_News
Fs_News.GetSysParam()
If Not Fs_news.IsSelfRefer Then response.write "�Ƿ��ύ����":Response.end
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
Set obj_newsedit_rs = server.CreateObject(G_FS_RS)
obj_newsedit_rs.Open "Select NewsID,PopId,ClassID,SpecialEName,NewsTitle,CurtTitle,NewsNaviContent,isShowReview,TitleColor,titleBorder,TitleItalic,IsURL,URLAddress,Content,isPicNews,NewsPicFile,NewsSmallPicFile,PicborderCss,Templet,isPop,Source,Editor,Keywords,Author,Hits,FileName,FileExtName,NewsProperty,TodayNewsPic,isLock,isRecyle,addtime,isdraft,IsAdPic,AdPicWH,AdPicLink,AdPicAdress from FS_NS_News where ClassID='"& NoSqlHack(Request.QueryString("ClassID")) &"' and NewsID='"& NoSqlHack(Request.QueryString("NewsID")) &"'",Conn,1,3
If obj_newsedit_rs.eof then 
	strShowErr = "<li>�Ƿ�����,�Ҳ������ݿ��¼</li>"
	Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
Else
	str_NewsID = obj_newsedit_rs("NewsID")
	Dim obj_tmppop_rs
	set obj_tmppop_rs = Conn.execute("select GroupName,PointNumber,FS_Money,InfoID,PopType,isClass From FS_MF_POP where InfoID='"& str_NewsID &"' and isClass=0 and PopType='NS'")
	if obj_tmppop_rs.eof then
			str_GroupID = ""
			str_PointNumber=""
			str_Money = ""
			obj_tmppop_rs.close:set obj_tmppop_rs = nothing
	Else
			str_GroupID = obj_tmppop_rs("GroupName")
			if obj_tmppop_rs("PointNumber") = 0 or isnull(trim(obj_tmppop_rs("PointNumber"))) then:str_PointNumber="" else:str_PointNumber=obj_tmppop_rs("PointNumber"):end if
			if obj_tmppop_rs("FS_Money") = 0 or isnull(trim(obj_tmppop_rs("FS_Money"))) then:str_Money="" else:str_Money=obj_tmppop_rs("FS_Money"):end if
			obj_tmppop_rs.close:set obj_tmppop_rs = nothing
	End if
	str_PopId = obj_newsedit_rs("PopId")
	str_ClassID = obj_newsedit_rs("ClassID")
	if not Get_SubPop_TF(str_ClassID,"NS002","NS","news") then Err_Show
	str_SpecialEName = obj_newsedit_rs("SpecialEName")
	str_NewsTitle = obj_newsedit_rs("NewsTitle")
	str_isdraft = obj_newsedit_rs("isdraft")
	str_CurtTitle = obj_newsedit_rs("CurtTitle")
	str_NewsNaviContent = obj_newsedit_rs("NewsNaviContent")
	if obj_newsedit_rs("isShowReview") = 1 then:str_isShowReview = 1:Else:str_isShowReview = 0:End if
	str_TitleColor = obj_newsedit_rs("TitleColor")
	str_titleBorder = obj_newsedit_rs("titleBorder")
	str_TitleItalic = obj_newsedit_rs("TitleItalic")
	if obj_newsedit_rs("IsURL") = 1 then:str_IsURL = 1:else:str_IsURL = 0:end if
	str_URLAddress = obj_newsedit_rs("URLAddress")
	str_Content = obj_newsedit_rs("Content")
	if obj_newsedit_rs("isPicNews") then:str_isPicNews =1:else:str_isPicNews = 0:end if
	str_NewsPicFile = obj_newsedit_rs("NewsPicFile")
	str_NewsSmallPicFile = obj_newsedit_rs("NewsSmallPicFile")
	str_PicborderCss = obj_newsedit_rs("PicborderCss")
	str_Templet =  obj_newsedit_rs("Templet")
	str_Source = obj_newsedit_rs("Source")
	str_Editor = obj_newsedit_rs("Editor")
	str_Keywords = obj_newsedit_rs("Keywords")
	str_Author = obj_newsedit_rs("Author")
	if obj_newsedit_rs("Hits") ="" then:str_Hits=0:else:str_Hits = obj_newsedit_rs("Hits"):end if
	str_FileName = obj_newsedit_rs("FileName")
	str_FileExtName = obj_newsedit_rs("FileExtName")
	str_NewsProperty = obj_newsedit_rs("NewsProperty")
	str_TodayNewsPic = obj_newsedit_rs("TodayNewsPic")
	
	IsAdPic = obj_newsedit_rs("IsAdPic")
	AdPicWH = obj_newsedit_rs("AdPicWH")
	AdPicLink = obj_newsedit_rs("AdPicLink")
	AdPicAdress = obj_newsedit_rs("AdPicAdress")
			
	if obj_newsedit_rs("isLock") = 1 then:str_islock=1:else:str_isLock = obj_newsedit_rs("isLock"):end if
	if obj_newsedit_rs("isRecyle") = 1 then:str_isRecyle=1:else:str_isRecyle = obj_newsedit_rs("isRecyle"):end if
	if trim(obj_newsedit_rs("addtime"))="" then:str_addtime=now:else:str_addtime = obj_newsedit_rs("addtime"):end if
	if obj_newsedit_rs("isRecyle") = 1 then
		strShowErr = "<li>�ڻ���վ�е�"& Fs_news.allInfotitle&"���ܱ༭</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
End if

'�����ʾ�����
if str_IsURL = 1 then
	str_NewsType = "TitleNews"
	str_UrLaddress_1 = ""
	str_CurtTitle_1="none"
	str_NewsPicFile_1=""
	str_NewsSmallPicFile_1=""
	str_Templet_1="none"
	str_Content_1 ="none"
	str_PicborderCss_1="none"
	str_Author_1="none"
	str_GroupID_1="none"
	str_FileName_1="none"
	str_filt_1="none"
Elseif str_isPicNews = 1 then
	str_NewsType = "PicNews"
	str_UrLaddress_1 = "none"
	str_CurtTitle_1=""
	str_NewsPicFile_1=""
	str_NewsSmallPicFile_1=""
	str_Templet_1=""
	str_Content_1 =""
	str_PicborderCss_1=""
	str_Author_1=""
	str_GroupID_1=""
	str_FileName_1=""
	str_filt_1=""
Else
	str_NewsType = "TextNews"
	str_UrLaddress_1 = "none"
	str_CurtTitle_1=""
	str_NewsPicFile_1="none"
	str_NewsSmallPicFile_1="none"
	str_Templet_1=""
	str_Content_1 =""
	str_PicborderCss_1=""
	str_Author_1=""
	str_GroupID_1=""
	str_FileName_1=""
	str_filt_1=""
End if
'On Error Resume Next
'��ȡ�����ֶ���Ϣ,���浽����CustColumnArr��
dim c_rs,tmp_defineid,i
Set c_rs = Conn.execute("select DefineID from FS_NS_NewsClass where Classid='"& str_ClassID &"'")
tmp_defineid = c_rs(0)
c_rs.close:set c_rs=nothing
if not isnull(trim(tmp_defineid)) or trim(tmp_defineid)>0 then
	Dim CustColumnRs,CustSql,CustColumnArr
	CustSql="select DefineID,ClassID,D_Name,D_Coul,D_Type,D_isNull,D_Value,D_Content,D_SubType from [FS_MF_DefineTable] Where D_SubType='NS' and  Classid="& CintStr(tmp_defineid) &""
	Set CustColumnRs=CreateObject(G_FS_RS)
	CustColumnRs.Open CustSql,Conn,1,3
	If Not CustColumnRs.Eof Then
		CustColumnArr=CustColumnRs.GetRows()
	End If
	CustColumnRs.close:Set CustColumnRs = Nothing
end if
'=====================================
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<script language="JavaScript" src="js/Public.js"></script>
<script language="JavaScript" src="../../FS_Inc/CheckJs.js"></script>
<script language="JavaScript" src="../../FS_Inc/Prototype.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
<iframe width="260" height="165" id="colorPalette" src="lib/selcolor.htm" style="visibility:hidden; position: absolute;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="1"></td>
	</tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
	<form action="News_Save.asp" name="NewsForm" method="post" onSubmit="return CheckForm(this);">
		<tr class="xingmu">
			<td colspan="4" class="xingmu">�༭/�޸�
				<% = Fs_news.allInfotitle %>
				<a href="../../help?Lable=NS_News_add" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr >
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
				<input name=NewsType type=radio id="NewsType" onClick="SwitchNewsType('TextNews');" value="TextNews" <%If str_NewsType = "TextNews" then Response.Write("checked")%>>
				��ͨ
				<input name=NewsType type=radio id="NewsType" onClick="SwitchNewsType('PicNews');" value="PicNews" <%If str_NewsType = "PicNews" then Response.Write("checked")%>>
				ͼƬ
				<input name=NewsType type=radio id="NewsType" onClick="SwitchNewsType('TitleNews');" value="TitleNews" <%If str_NewsType = "TitleNews" then Response.Write("checked")%>>
				���� ��&nbsp;&nbsp;
				<%if  str_isdraft = 1 then%>
				<input name="isdraft" type="checkbox" id="isdraft" value="1" <%if str_isdraft = 1 then response.Write("checked")%>>
				�浽�ݸ�����
				<%end if%>
			</td>
		</tr>
		<tr >
			<td width="12%" class="hback">
				<div align="right">
					<% = Fs_news.allInfotitle %>
					����</div>
			</td>
			<td colspan="3" class="hback">
				<input name="NewsTitle" type="text" id="NewsTitle" size="40"  value="<%=str_NewsTitle%>" maxlength="255" onFocus="Do.these('NewsTitle',function(){return isEmpty('NewsTitle','span_NewsTitle')})" onKeyUp="Do.these('NewsTitle',function(){return isEmpty('NewsTitle','span_NewsTitle')})" style="background-image:url(../Images/bg.gif);">
				<span id="span_NewsTitle"></span>
				<input name="TitleColor" id="TitleColor" type="hidden" value="<% = str_TitleColor %>">
				
				<img src="images/rect<%if str_TitleColor="" Then Response.write("NoColor")%>.gif" id="TitleColorShow" style="cursor:pointer;background-color:<%=str_TitleColor%>;" title="ѡȡ��ɫ!" onClick="GetColor(document.getElementById('TitleColorShow'),'TitleColor');">
				<input name="titleBorder" type="checkbox" id="titleBorder" value="1"  <%if str_titleBorder=1 then response.Write("checked") %>>
				����
				<input name="TitleItalic" type="checkbox" id="TitleItalic" value="1"  <%if str_TitleItalic=1 then response.Write("checked") %>>
				б��
				<input name="isShowReview" type="checkbox" id="isShowReview" value="1"<%if str_isShowReview=1 then response.Write("checked") %>>
				�������ӡ���Ȩ��
				<select name="PopID" id="PopID">
					<option value="5" <%if str_PopID=5 then response.Write("selected") %>>���ö�</option>
					<option value="4" <%if str_PopID=4 then response.Write("selected") %>>��Ŀ�ö�</option>
					<option value="0" <%if str_PopID=0 then response.Write("selected") %>>һ��</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">������ </div>
			</td>
			<td colspan="3" class="hback"><input name="CurtTitle" type="text" id="CurtTitle" style="width:85%;" maxlength="255" value="<%=str_CurtTitle%>"></td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right">ѡ����Ŀ</div>
			</td>
			<td colspan="3" class="hback">
				<input name="ClassName" type="text" id="ClassName5" style="width:45%" value="<%=Fs_News.GetAdd_ClassName(str_ClassID)%>" readonly>
				<input name="ClassID" type="hidden" id="ClassID" value="<% = str_ClassID %>">
				<input type="button" name="Submit" value="ѡ����Ŀ"   onClick="SelectClass();">
				<input type="button" name="Submit2" value="�����Ŀ" onClick="window.location.href='Class_add.asp?ClassID=<%=str_ClassID %>&Action=add'">
			</td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right">ѡ��ר��</div>
			</td>
			<%
		  if str_SpecialEName<>"" then
			  dim sp_rs,sp_array,sp_i,sp_char
			  sp_char=""
			  sp_array = split(str_SpecialEName,",")
			  for sp_i = 0 to ubound(sp_array)
				  set sp_rs=Conn.execute("select SpecialCName From FS_NS_Special where SpecialEName='"& NoSqlHack(sp_array(sp_i)) &"' ")
				   if not sp_rs.eof then
						if sp_i = ubound(sp_array) then
							sp_char = sp_char&sp_rs("SpecialCName")
						else
							sp_char = sp_char&sp_rs("SpecialCName")&","
						end if
				   end if
				   sp_rs.close:set sp_rs = nothing
			  next
			  sp_char = sp_char
		  end if
	  %>
			<td colspan="3" class="hback">
				<input name="SpecialID" type="text" id="SpecialID" style="width:45%" readonly value="<% = sp_char%>">
				<input name="SpecialID_EName" type="hidden" id="SpecialID_EName" value="<%=str_SpecialEName%>">
				<span class="tx"> </span>
				<input type="button" name="Submit" value="ѡ��ר��"   onClick="SelectSpecial();">
				<span class="tx">
				<input name="Submit" type="button" id="Submit" onClick="dospclear();" value="���ר��">
				</span> <span class="tx"> ���޸��뱣��Ϊ��,�ɶ�ѡ</span> &nbsp;&nbsp;<a href="../../help?Lable=NS_News_add_special" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr  id="str_URLAddress" style="display:<%=str_URLAddress_1%>;" >
			<td class="hback">
				<div align="right">���ӵ�ַ </div>
			</td>
			<td colspan="3" class="hback">
				<input name="URLAddress" type="text" id="URLAddress"  style="width:96%" maxlength="255" value="<%=str_URLAddress%>">
			</td>
		</tr>
		<tr  id="str_CurtTitle" style="display:<%=str_CurtTitle_1%>" >
			<td class="hback">
				<div align="right">�ؼ���</div>
			</td>
			<td colspan="3" class="hback">
				<input name="KeywordText" type="text" id="KeywordText" size="15" maxlength="255" value="<%=str_Keywords%>">
				<input name="KeyWords" type="hidden" id="KeyWords" value="<%=str_Keywords%>">
				<select name="selectKeywords" id="selectKeywords" style="width:120px" onChange=Dokesite(this.options[this.selectedIndex].value)>
					<option value="" selected>ѡ��ؼ���</option>
					<option value="Clean" style="color:red">���</option>
					<%=Fs_news.GetKeywordslist("",1)%>
				</select>
				<input name="KeywordSaveTF" type="checkbox" id="KeywordSaveTF" value="1">
				����</td>
		</tr>
		<tr  id="str_Templet" style="display:<%=str_Templet_1%>" >
			<td class="hback">
				<div align="right">ģ���ַ</div>
			</td>
			<td colspan="3" class="hback">
				<input name="Templet" type="text" id="Templet" style="width:85%"  value="<%=str_Templet%>"  maxlength="255" readonly>
				<input name="Submit5" type="button" id="selNewsTemplet" value="ѡ��ģ��"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.NewsForm.Templet);document.NewsForm.Templet.focus();">
			</td>
		</tr>
		<tr   id="str_NewsSmallPicFile" style="display:<%=str_NewsSmallPicFile_1%>">
			<td class="hback">
				<div align="right">ͼƬ(С)</div>
				<div align="right">ͼƬ(��)</div>
			</td>
			<td colspan="3" class="hback">
				<table width="417" border="0" cellspacing="1" cellpadding="5">
					<tr>
						<td width="50%">
							<table width="10" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
								<tr>
									<%if isnull(trim(str_NewsSmallPicFile)) or str_NewsSmallPicFile="" then%>
									<td class="hback"><img src="../Images/nopic_supply.gif" id="pic_p_1" onload="Javascript:if(this.width > 90 || this.height > 90){if(this.width > this.height){this.width=90;}else{this.height=90;}}" /></td>
									<%else%>
									<td class="hback"><img src="<%=str_NewsSmallPicFile%>" id="pic_p_1" /></td>
									<%end if%>
								</tr>
							</table>
							<div align="center">
								<div align="center">
									<input name="NewsSmallPicFile" type="hidden" id="NewsSmallPicFile" style="width:85%" maxlength="255" value="<%=str_NewsSmallPicFile%>">
									<img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>',500,320,window,document.NewsForm.NewsSmallPicFile);" style="cursor:hand;"> ��<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_1();" style="cursor:hand;"> </div>
							</div>
						</td>
						<td width="50%">
							<table width="10" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
								<tr>
									<%if isnull(trim(str_NewsPicFile)) or str_NewsPicFile="" then%>
									<td class="hback"><img src="../Images/nopic_supply.gif" id="pic_p_2" onload="Javascript:if(this.width > 90 || this.height > 90){if(this.width > this.height){this.width=90;}else{this.height=90;}}" /></td>
									<%else%>
									<td class="hback"><img src="<%=str_NewsPicFile%>" id="pic_p_2" /></td>
									<%end if%>
								</tr>
							</table>
							<div align="center">
								<div align="center">
									<input name="NewsPicFile" type="hidden" id="NewsPicFile" style="width:85%" maxlength="255" value="<%=str_NewsPicFile%>">
									<img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>',500,320,window,document.NewsForm.NewsPicFile);" style="cursor:hand;"> ��<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_2();" style="cursor:hand;"> </div>
							</div>
						</td>
					</tr>
					<tr>
						<td class="hback">
							<div align="center">Сͼ��ַ</div>
						</td>
						<td class="hback">
							<div align="center">��ͼ��ַ</div>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr  id="str_Author" style="display:<%=str_Author_1%>" >
			<td class="hback">
				<div align="right">
					<% = Fs_news.allInfotitle %>
					����</div>
			</td>
			<td class="hback">
				<input name="Author" type="text" id="Author" size="15" maxlength="255"  value="<%=str_Author%>" >
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
				<input name="Source" type="text" id="Source" size="15" maxlength="255" value="<%=str_Source%>">
				<select name="selectSource" id="selectSource" style="width:120px"  onChange="document.NewsForm.Source.value=this.options[this.selectedIndex].text;">
					<option value="" selected> </option>
					<option value="��վԭ��">��վԭ��</option>
					<option value="����">����</option>
					<%=Fs_news.GetKeywordslist("",2)%>
				</select>
				<input name="SourceSaveTF" type="checkbox" id="SourceSaveTF" value="1">
				����</td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right">����</div>
			</td>
			<td colspan="3" class="hback">
				<div align="left">
					<input name="NewsProperty_Rec" type="checkbox" id="NewsProperty" value="1" <%if split(str_NewsProperty,",")(0)="1" then Response.Write("checked")%>>
					�Ƽ�
					<input name="NewsProperty_mar" type="checkbox" id="NewsProperty" value="1"  <%if split(str_NewsProperty,",")(1)="1" then Response.Write("checked")%>>
					����
					<input name="NewsProperty_rev" type="checkbox" id="NewsProperty" value="1"  <%if split(str_NewsProperty,",")(2)="1" then Response.Write("checked")%>>
					��������
					<input name="NewsProperty_constr" type="checkbox" id="NewsProperty" value="1" <%if split(str_NewsProperty,",")(3)="1" then Response.Write("checked")%>>
					Ͷ��
					<input name="NewsProperty_tt" type="checkbox" id="NewsProperty" value="1"  onClick="ChooseTodayNewsType();" <%if split(str_NewsProperty,",")(5)="1" then Response.Write("checked")%>>
					ͷ��
					<input name="NewsProperty_hots" type="checkbox" id="NewsProperty" value="1" <%if split(str_NewsProperty,",")(6)="1" then Response.Write("checked")%> disabled="disabled">
					�ȵ�
					<input name="NewsProperty_jc" type="checkbox" id="NewsProperty" value="1" <%if split(str_NewsProperty,",")(7)="1" then Response.Write("checked")%>>
					����
					<input name="NewsProperty_unr" type="checkbox" id="NewsProperty" value="1" <%if split(str_NewsProperty,",")(8)="1" then Response.Write("checked")%>>
					������
					<input name="NewsProperty_ann" type="checkbox" id="NewsProperty" value="1" <%if split(str_NewsProperty,",")(9)="1" then Response.Write("checked")%>>
					���� <span id="str_filt" style="display:<%=str_filt_1%>">
					<input name="NewsProperty_filt" type="checkbox" id="NewsProperty" value="1" <%if split(str_NewsProperty,",")(10)="1" then Response.Write("checked")%>>
					�õ�</span></div>
			</td>
		</tr>
		<tr  id="TodayNews" style="display:<%if split(str_NewsProperty,",")(5) =1 then:response.Write(";"):else:Response.Write("none;"):end if%>" >
			<td colspan="4" class="hback">
				<table width="100%" border="0" cellspacing="1" cellpadding="2" class="table">
					<tr>
						<td height="26" align="center" width="120" class="xingmu">ͷ�����ͣ�</td>
						<td height="26" class="hback">
							<input name="TodayNewsPicTF" value="" type="radio" checked onClick="if(this.checked){document.getElementById('TodayPicParam').style.display='none';}">
							����ͷ��
							<input name="TodayNewsPicTF" value="FoosunCMS" type="radio" onClick="if(this.checked){document.getElementById('TodayPicParam').style.display='';}" <%if str_TodayNewsPic=1 then Response.Write("checked")%>>
							ͼƬͷ�� ����<a href="../../help?Lable=NS_News_add_tt" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
					</tr>
					<tr id="TodayPicParam" style="display:<%if str_TodayNewsPic =1 then:response.Write(";"):else:Response.Write("none;"):end if%>;" >
						<td width="120" height="26" align="center"  class="xingmu">ͷ��������</td>
					  <td height="26" class="hback">&nbsp;&nbsp;����
							<%
							Dim Get_TodayPic,FontFace,FontSize,FontColor,FontSpace,FontBgColor,TodayTitle,Todaywidth
							set Get_TodayPic = Conn.execute("select TodayPic_font,TodayPic_size,TodayPic_color,TodayPic_space,TodayPic_PicColor,TodayTitle,Todaywidth From  FS_NS_TodayPic where NewsID='"& str_NewsID &"'")
							if not Get_TodayPic.eof  then
								FontFace = Get_TodayPic("TodayPic_font")
								FontSize =  Get_TodayPic("TodayPic_size")
								FontColor =  Get_TodayPic("TodayPic_color")
								FontSpace =  Get_TodayPic("TodayPic_space")
								FontBgColor =  Get_TodayPic("TodayPic_PicColor")
								TodayTitle = Get_TodayPic("TodayTitle")
								Todaywidth = Get_TodayPic("Todaywidth")
								Get_TodayPic.close:set Get_TodayPic = nothing
							else
								FontFace = "黑体"
								FontSize = "12"
								FontColor = "000000"
								FontSpace = "12"
								FontBgColor ="FFFFFF"
								TodayTitle = ""
								Todaywidth = 300
								Get_TodayPic.close:set Get_TodayPic = nothing
							end if
							%>
							<SELECT name="FontFace" id="FontFace">
								<option value="����" <%if FontFace = "����" then response.Write("selected")%>>����</option>
								<option value="����_GB2312" <%if FontFace = "����_GB2312" then response.Write("selected")%>>����_GB2312</option>
								<option value="������" <%if FontFace = "������" then response.Write("selected")%>>������</option>
								<option value="����" <%if FontFace = "����" then response.Write("selected")%>>����</option>
								<option value="����" <%if FontFace = "����" then response.Write("selected")%>>����</option>
								<OPTION value="Andale Mono" <%if FontFace = "Andale Mono" then response.Write("selected")%>>Andale 
								Mono</OPTION>
								<OPTION value="Arial" <%if FontFace = "Arial" then response.Write("selected")%>>Arial</OPTION>
								<OPTION value="Arial Black" <%if FontFace = "Arial Black" then response.Write("selected")%>>Arial 
								Black</OPTION>
								<OPTION value="Book Antiqua"  <%if FontFace = "Book Antiqua" then response.Write("selected")%>>Book 
								Antiqua</OPTION>
								<OPTION value="Century Gothic" <%if FontFace = "Century Gothic" then response.Write("selected")%>>Century 
								Gothic</OPTION>
								<OPTION value="Comic Sans MS" <%if FontFace = "Comic Sans MS" then response.Write("selected")%>>Comic 
								Sans MS</OPTION>
								<OPTION value="Courier New" <%if FontFace = "Courier New" then response.Write("selected")%>>Courier 
								New</OPTION>
								<OPTION value="Georgia" <%if FontFace = "Georgia" then response.Write("selected")%>>Georgia</OPTION>
								<OPTION value="Impact" <%if FontFace = "Impact" then response.Write("selected")%>>Impact</OPTION>
								<OPTION value="Tahoma" <%if FontFace = "Tahoma" then response.Write("selected")%>>Tahoma</OPTION>
								<OPTION value="Times New Roman" <%if FontFace = "Times New Roman" then response.Write("selected")%>>Times 
								New Roman</OPTION>
								<OPTION value="Trebuchet MS" <%if FontFace = "Trebuchet MS" then response.Write("selected")%>>Trebuchet 
								MS</OPTION>
								<OPTION value="Script MT Bold" <%if FontFace = "Script MT Bold" then response.Write("selected")%>>Script 
								MT Bold</OPTION>
								<OPTION value="Stencil" <%if FontFace = "Stencil" then response.Write("selected")%>>Stencil</OPTION>
								<OPTION value="Verdana" <%if FontFace = "Verdana" then response.Write("selected")%>>Verdana</OPTION>
								<OPTION value="Lucida Console" <%if FontFace = "Lucida Console" then response.Write("selected")%>>Lucida 
								Console</OPTION>
							</SELECT>
							<select name="FontSize">
								<option value="8" <%if FontSize = "8" then response.Write("selected")%>>8px</option>
								<option value="9" <%if FontSize = "9" then response.Write("selected")%>>9px</option>
								<option value="10" <%if FontSize = "10" then response.Write("selected")%>>10px</option>
								<option value="12" <%if FontSize = "12" then response.Write("selected")%>>12px</option>
								<option value="18" <%if FontSize = "18" then response.Write("selected")%>>18px</option>
								<option value="20" <%if FontSize = "20" then response.Write("selected")%>>20px</option>
								<option value="24" <%if FontSize = "24" then response.Write("selected")%>>24px</option>
								<option value="28" <%if FontSize = "28" then response.Write("selected")%>>28px</option>
								<option value="30" <%if FontSize = "30" then response.Write("selected")%>>30px</option>
								<option value="32" <%if FontSize = "32" then response.Write("selected")%>>32px</option>
								<option value="36" <%if FontSize = "36" then response.Write("selected")%>>36px</option>
								<option value="40" <%if FontSize = "40" then response.Write("selected")%>>40px</option>
								<option value="48" <%if FontSize = "48" then response.Write("selected")%>>48px</option>
								<option value="54" <%if FontSize = "54" then response.Write("selected")%>>54px</option>
								<option value="60" <%if FontSize = "60" then response.Write("selected")%>>60px</option>
								<option value="72" <%if FontSize = "72" then response.Write("selected")%>>72px</option>
							</select>
							<input type="text" name="FontColor" maxlength=7 size=7 id="FontColor" value="<% = FontColor %>">
							<img src="images/rect.gif" width="18" height="17" border=0 align=absmiddle id="FontColorShow" style="cursor:pointer;background-Color:#<% = FontColor %>;" title="ѡȡ��ɫ!" onClick="GetColor(this,'FontColor');"> ���� �����ࣺ
							<INPUT TYPE="text" maxlength="3" NAME="FontSpace" size=3 value="<% = FontSpace %>">
							px ͼƬ����ɫ
							<input type="text" name="FontBgColor" maxlength=7 size=7 id="FontBgColor" value="<% = FontBgColor %>">
						  <img src="images/rect.gif" width="18" height="17" border=0 align=absmiddle id="FontBgColorShow" style="cursor:pointer;background-Color:<% = FontBgColor %>;" title="ѡȡ��ɫ!" onClick="GetColor(this,'FontBgColor');"> <br>
ͼƬͷ�����⣺
<input name="PicTitle" type="text" id="PicTitle" value="<% = TodayTitle %>" size="40" maxlength="255">
&nbsp;&nbsp;ͼƬ��ȣ�
<input name="PicTitlewidth" type="text" id="PicTitlewidth" value="<% = Todaywidth %>" size="10" maxlength="10">
px </td>
					</tr>
				</table>
			</td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right">
					<% = Fs_news.allInfotitle %>
					����</div>
			</td>
			<td colspan="3" class="hback">
				<div align="left">
					<textarea name="NewsNaviContent" rows="6" id="NewsNaviContent" style="width:96%"><%=str_NewsNaviContent%></textarea>
				</div>
			</td>
		</tr>
		<tr id="str_Content" style="display:<%=str_Content_1%>" >
			<td class="hback">
				<div align="left">				  ���ݷ�ҳ��ǩ[FS:PAGE]<br>
                  <a href="javascript:void(0);" onClick="InsertHTML('[FS:PAGE]','NewsContent')"><span class="tx">�����ҳ��ǩ</span></a><br>
			    <input name="NewsProperty_Remote" type="checkbox" id="NewsProperty_Remote" value="1" <%if split(str_NewsProperty,",")(4)="1" then Response.Write("checked")%>>
					���������е�ͼƬ������<br>
          <span class="tx">���ô˹��ܺ������������վ�ϸ������ݵ��ұߵı༭���У����������а�����ͼƬ����ϵͳ���ڱ�������ʱ�Զ������ͼƬ���Ƶ���վ�������ϡ�<br>
          ϵͳ����������ͼƬ�Ĵ�С��Ӱ���ٶȣ�����ͼƬ�϶�ʱ��Ҫʹ�ô˹��ܡ�</span><br>
          <input name="ClearAllPage" type="checkbox" id="ClearAllPage" value="1">���¼����Զ���ҳ(<%=G_FS_Page_Txtlength&"�ֽ�/ҳ"%>)</div>
		  </td>
			<%
                Dim  WZApic,temp
                if instr(str_Content,"<!---���ֻ��л�star---->")>0 then
                     WZApic=right(str_Content,len(str_Content)-instr(str_Content,"<!---���ֻ��л�star---->")-len("<!---���ֻ��л�star---->")+1)
                     WZApic=left(WZApic,instr(WZApic,"<!---���ֻ��л�end--->")-1)
                     str_Content=replace(str_Content,"<!---���ֻ��л�star---->"&WZApic&"<!---���ֻ��л�end--->","")
                     
                end if
			%>
			<td colspan="3" class="hback">
                <!--�༭����ʼ--><input name="ParentDispalyNone" type="hidden" value="<%=str_Content_1%>">
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=Content' frameborder=0 scrolling=no width='100%' height='440'></iframe>
				<input type="hidden" name="Content" value="<% = HandleEditorContent(str_Content) %>">
                <!--�༭������-->
			</td>
		</tr>
		<!--�Զ����Զο�ʼ-->
		<%
	If IsArray(CustColumnArr) Then
		response.Write"<tr><td colspan=""4"" class=""hback_1"">�Զ��忪ʼ</td></tr>"
		Dim NewsAuxiInfoRs,NewsAuxiInfoSql
		Dim InputModeStr,AuxiInfoList,AuxiListArr,k
		For i = 0 to Ubound(CustColumnArr,2)
			NewsAuxiInfoSql="select ColumnValue From FS_MF_DefineData Where InfoID='"&str_NewsID&"' and TableEName='" & NoSqlHack(CustColumnArr(3,i)) & "' And InfoType='NS'"
			Set NewsAuxiInfoRs=Conn.Execute(NewsAuxiInfoSql)
			dim dvalues
			if not NewsAuxiInfoRs.eof then
				dvalues=NewsAuxiInfoRs(0)
			else
				dvalues=""
			end if

			Select Case CStr(CustColumnArr(4,i))		'����ѡ������������������뷽ʽ
				Case 1	'��������
					If Not NewsAuxiInfoRs.Eof Then 		'���Ϊ��ǰ��ӵ�������س��ֶ��������ݵ�������������
						 InputModeStr="<Textarea Name=""FS_NS_Define_"&CustColumnArr(3,i)&""" style=""width:70%;height:80;"">"&dvalues&"</Textarea>"
					Else
						InputModeStr="<Textarea Name=""FS_NS_Define_"&CustColumnArr(3,i)&""" style=""width:70%;height:80;""></Textarea>"
					End If
				Case 4	'�б�ѡ��
					Dim AuxiDictRs
					InputModeStr="<Select Name=""FS_NS_Define_"&CustColumnArr(3,i)&""" style=""width:70%"">"&vbcrlf
					AuxiListArr=Split(CustColumnArr(6,i),vbcrlf)
					For k = 0 to UBound(AuxiListArr)
						If AuxiListArr(k)<>"" Then
							If Not NewsAuxiInfoRs.eof  Then
								If trim(AuxiListArr(k))=trim(NewsAuxiInfoRs(0)) Then
									InputModeStr=InputModeStr&"<Option value="""&AuxiListArr(k)&""" selected>"&AuxiListArr(k)&"</option>"&vbcrlf
								Else
									InputModeStr=InputModeStr&"<Option value="""&AuxiListArr(k)&""">"&AuxiListArr(k)&"</option>"&vbcrlf
								End IF
							Else
								If k=0 Then InputModeStr=InputModeStr&"<Option value="""" selected> </option>"&vbcrlf 
								InputModeStr=InputModeStr&"<Option value="""&AuxiListArr(k)&""">"&AuxiListArr(k)&"</option>"&vbcrlf
							End If
						End If
					Next
					InputModeStr=InputModeStr&"</select>"
				Case 7
						InputModeStr ="<Input Type=""text"" Name=""FS_NS_Define_"&CustColumnArr(3,i)&""" value="""&dvalues&""" style=""width:70%""><input name=""SelectAdPic"" type=""button"" id=""SelectAdPic"" onClick=""OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath="&str_CurrPath&"',500,300,window,document.NewsForm.FS_NS_Define_"&CustColumnArr(3,i)&");""  value=""ѡ��ͼƬ"">"&vbcrlf
				Case 8
						InputModeStr ="<Input Type=""text"" Name=""FS_NS_Define_"&CustColumnArr(3,i)&""" value="""&dvalues&""" style=""width:70%""> <input name=""SelectAdPic"" type=""button"" id=""SelectAdPic"" onClick=""OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath="&str_CurrPath&"',500,300,window,document.NewsForm.FS_NS_Define_"&CustColumnArr(3,i)&");""  value=""ѡ�񸽼�"">"&vbcrlf
				Case 9
						InputModeStr ="<Input Type=""text"" Name=""FS_NS_Define_"&CustColumnArr(3,i)&""" value="""&dvalues&""" style=""width:70%"">"&vbcrlf
				Case 10
						dim scriptSTR,DefineFieldTextAreaID
						DefineFieldTextAreaID = CustColumnArr(3,i)
						scriptSTR = "<iframe id='DefineFieldContent' src='../Editer/AdminEditer.asp?id=FS_NS_Define_" & DefineFieldTextAreaID & "' frameborder=0 scrolling=no width='560' height='300'></iframe>"
						InputModeStr ="<textarea name=""FS_NS_Define_" & DefineFieldTextAreaID & """ id=""FS_NS_Define_" & DefineFieldTextAreaID & """ style=""display: none"">" & HandleEditorContent(dvalues) & "</textarea>"&scriptSTR&""&vbcrlf
				Case Else	'���У����֣�����
					If Not NewsAuxiInfoRs.Eof Then
						InputModeStr="<Input Type=""text"" Name=""FS_NS_Define_"&CustColumnArr(3,i)&""" value="""&dvalues&""" style=""width:70%"">"
					Else
						InputModeStr="<Input Type=""text"" Name=""FS_NS_Define_"&CustColumnArr(3,i)&""" value="""" style=""width:70%"">"
					End If
			End Select
			If CStr(CustColumnArr(7,i))<>"" Then
				InputModeStr=InputModeStr&"<span class=""tx"">"&CustColumnArr(7,i)&"</span>"
			End If
			Response.Write "<tr >"&vbcrlf
			Response.Write "<td width=""10%"" align=""right""  class=""hback"">"&CustColumnArr(2,i)&"</td>"&vbcrlf
			Response.Write "<td width=""90%"" colspan=""3"" align=""left""    class=""hback"">"&vbcrlf&InputModeStr&"</td>"&vbcrlf
			Response.Write "</tr>"&vbcrlf
			NewsAuxiInfoRs.Close
		Next
			response.Write"<tr><td colspan=""4"" class=""hback_1"">�Զ������</td></tr>"
	End If
	Set NewsAuxiInfoRs=Nothing
	%>
		<!--�Զ����Զν���-->
		<tr  id="str_GroupID" style="display:<%=str_GroupID_1%>" >
			<td class="hback">
				<div align="right">�������</div>
			</td>
			<td colspan="3" class="hback">
				<input name="PointNumber" type="text" id="PointNumber2" size="16"  onChange="ChooseExeName();" value="<% = str_PointNumber%>">
				���
				<input name="Money" type="text" id="Money2" size="16"  onChange="ChooseExeName();" value="<% = str_Money%>">
				���Ȩ��
				<input name="BrowPop"  id="BrowPop" type="text" onMouseOver="this.title=this.value;" readonly value="<% = str_GroupID%>">
				<select name="selectPop" id="selectPop" style="overflow:hidden;" onChange="ChooseExeName();">
					<option value="" selected>ѡ���Ա��</option>
					<option value="del" style="color:red;">���</option>
					<% = MF_GetUserGroupID %>
				</select>
				<a href="../../help?Lable=NS_News_add_pop" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr  id="str_FileName" style="display:<%=str_FileName_1%>" >
			<td class="hback">
				<div align="right">�ļ���</div>
			</td>
			<td class="hback">
				<input name="FileName" type="text" value="<% = str_FileName %>" readonly>
				<span class="tx">�����޸�</span> </td>
			<td class="hback">
				<div align="right">��չ��</div>
			</td>
			<td class="hback">
				<input type="hidden" name="DefaultFileExtName" id="DefaultFileExtName" value="<% = str_FileExtName %>">
				<select name="FileExtName" id="FileExtName">
					<option value="html" <%if str_FileExtName = "html" then response.Write("selected")%>>.html</option>
					<option value="htm" <%if str_FileExtName = "htm" then response.Write("selected")%>>.htm</option>
					<option value="shtml" <%if str_FileExtName = "shtml" then response.Write("selected")%>>.shtml</option>
					<option value="shtm" <%if str_FileExtName = "shtm" then response.Write("selected")%>>.shtm</option>
					<option value="asp" <%if str_FileExtName = "asp" then response.Write("selected")%>>.asp</option>
				</select>
			</td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right">�������</div>
			</td>
			<td class="hback">
				<input name="addtime" type="text" id="addtime" value="<% = str_addtime %>" size="40" maxlength="255">
			</td>
			<td class="hback">
				<div align="right">�������</div>
			</td>
			<td class="hback">
				<input name="Hits" type="text" id="Hits"  value="<% = str_hits %>" size="20" onFocus="Do.these('Hits',function(){return isNumber('Hits','span_Hits','����д��ȷ�ĸ�ʽ',true)})" onKeyUp="Do.these('Hits',function(){return isNumber('Hits','span_Hits','����д��ȷ�ĸ�ʽ',true)})">
				<span id="span_Hits"></span></td>
		</tr>
		<tr id="IsShowAdpic">
			<td class="hback"><div align="right">�Ƿ���ʾ���л�</div></td>
			<td colspan="3" class="hback"><input name="IsAdPic" type="checkbox" id="IsAdPic" value="1" onClick="javascript:ShowAdpicInfo();" <% If Cint(IsAdPic)=1 or Cint(IsAdPic)=2 Then Response.Write("checked") %>></td>
		</tr>
		<tr id="selectAp" style="display:none" class="hback">
		<td class="hback"></td>
		    <td  colspan="2" class="hback" align="left"> ͼƬ���л�
		
                <input id="Checkbox1" name="Checkbox1" type="checkbox" onClick="javascript:ShowAdpicInfo1();" <% If Cint(IsAdPic)=1 Then Response.Write("checked") %>> &nbsp;&nbsp;&nbsp;���ֻ��л�
		     
                <input id="Checkbox2" name="Checkbox2" type="checkbox"  onClick="javascript:ShowAdpicInfo2();" <% If  Cint(IsAdPic)=2 Then Response.Write("checked") %>>
             </td>
             		<td class="hback"></td>

		</tr>
		<tr id="Adpic" style="display:none" class="hback">
			<td class="hback" colspan="4"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
              <tr>
                <td width="12%" height="2" class="hback"><div align="right">���л���������</div></td>
                <td width="88%" height="2" class="hback"><input name="AdPicWH" type="text" id="AdPicWH" size="20" maxlength="20" value="<%if Cint(IsAdPic)=1 then response.write(AdPicWH) end if%>">
(���,�߶�,��(1)��(0),����λ������������ǰ������(������)����100,100,1,400)</td>
              </tr>
              <tr>
                <td height="5" class="hback"><div align="right">ͼƬ��ַ</div></td>
                <td height="5" class="hback"><input name="AdPicAdress" type="text" id="AdPicAdress"  size="20" maxlength="250" readonly value="<%=AdPicAdress%>">
                    <input name="SelectAdPic" type="button" id="SelectAdPic" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath%>',500,300,window,document.NewsForm.AdPicAdress);"  value="ѡ��ͼƬ��FLASH">
                  ���ӵ�ַ
                  <input name="AdPicLink" type="text" id="AdPicLink"  size="36" maxlength="250" value="<%=AdPicLink%>"></td>
              </tr>
            </table></td>
		</tr>
		
		<tr id="wzPic" style="display:none" class="hback">
		         <td colspan="4">
		        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		                  <tr>
                <td width="12%" height="2" class="hback"><div align="right">���л�����</div></td>
                <td width="88%" height="2" class="hback"><input name="AdPicWHw" type="text" id="Text2" size="20" maxlength="20" value="<%if Cint(IsAdPic)=2 then response.write(AdPicWH) end if%>">
                (����λ������������ǰ������(������)����100</td>
              </tr>
              <tr>
	         <td class="hback" align="right">���л�����
		     </td>
		     <td class="hback" colspan="3"  align="left">
                <textarea id="IsApicArea" name="IsApicArea" cols="80" rows="10"><%
				if WZApic<>""  then
					WZApic=replace(WZApic,"<table width=0 border=0 align="&G_CodeContentAlign&"><tr><td>","")
					WZApic=replace(WZApic,"</td></tr></table>","")
					WZApic=replace(WZApic,"<!---���ֻ��л�star---->","")
					WZApic=replace(WZApic,"<!---���ֻ��л�end--->","")
				end if
				response.Write(WZApic)
				%></textarea>
		      </td>
		     </tr>
		     </table>
		    </td>
		</tr>
		
		<tr >
			<td class="hback">
				<div align="right"></div>
			</td>
			<td colspan="3" class="hback">
			<script language="javascript">
			function SubmitFun()
			{
				if (frames["NewsContent"].g_currmode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}
				document.NewsForm.Content.value=frames["NewsContent"].GetNewsContentArray();
				if(frames["DefineFieldContent"])document.NewsForm.FS_NS_Define_<% = DefineFieldTextAreaID %>.value=frames["DefineFieldContent"].GetNewsContentArray();
				document.NewsForm.submit();
			}
			</script>
				<input type="button" name="Submit" value="ȷ�ϱ���<% = Fs_news.allInfotitle %>" onClick="SubmitFun();">
				<input type="reset" name="Submit" value="��������">
				<input name="News_Action" type="hidden" id="News_Action2" value="Edit_Save">
				<input name="NewsID" type="hidden" id="NewsID" value="<% = str_NewsID %>">
				<input name="d_Id" type="hidden" id="d_Id" value="<%=tmp_defineid%>">
			</td>
		</tr>
	</form>	
</table>
</body>
</html>
<%

If Cint(IsAdPic)=1 Then Response.Write("<script language=""javascript"">document.getElementById('Adpic').style.display='';</script>")
set Fs_news = nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckForm(FormObj)
{
	if(FormObj.ClassName.value=="")
	{
		alert("��ѡ����Ŀ��");
		FormObj.ClassName.focus();
		return false;
	}
	if(FormObj.ClassID.value=="")
	{
		alert("��Ŀ��������");
		FormObj.ClassName.focus();
		return false;
	}
	if(document.NewsForm.NewsTitle.value == "")
	{
		alert("����д���⣡");
		FormObj.NewsTitle.focus();
		return false;
	}
	return true;
}

function SwitchNewsType(NewsType)
{
	switch (NewsType)
	{
	case "TitleNews":
		document.getElementById('str_UrLaddress').style.display='';
		document.getElementById('str_CurtTitle').style.display='none';
		document.getElementById('str_NewsSmallPicFile').style.display='';
		document.getElementById('str_Templet').style.display='none';
		document.getElementById('str_Content').style.display='none';
		document.getElementById('str_Author').style.display='none';
		document.getElementById('str_GroupID').style.display='none';
		document.getElementById('str_FileName').style.display='none';
		document.getElementById('str_filt').style.display='none';
		document.getElementById('IsShowAdpic').style.display='none';
		document.getElementById('Adpic').style.display='none';
		break;
	case "PicNews":
		document.getElementById('str_UrLaddress').style.display='none';
		document.getElementById('str_CurtTitle').style.display='';
		document.getElementById('str_NewsSmallPicFile').style.display='';
		document.getElementById('str_Templet').style.display='';
		document.getElementById('str_Content').style.display='';
		document.getElementById('str_Author').style.display='';
		document.getElementById('str_GroupID').style.display='';
		document.getElementById('str_FileName').style.display='';
		document.getElementById('str_filt').style.display='';
		document.getElementById('IsShowAdpic').style.display='';
		document.getElementById('Adpic').style.display='none';
		break;
	default :
		document.getElementById('str_UrLaddress').style.display='none';
		document.getElementById('str_CurtTitle').style.display='';
		document.getElementById('str_NewsSmallPicFile').style.display='none';
		document.getElementById('str_Templet').style.display='';
		document.getElementById('str_Content').style.display='';
		document.getElementById('str_Author').style.display='';
		document.getElementById('str_GroupID').style.display='';
		document.getElementById('str_FileName').style.display='';
		document.getElementById('str_filt').style.display='none';
		document.getElementById('IsShowAdpic').style.display='';
		document.getElementById('Adpic').style.display='none';
	}
}
function SetEditerFrame(){
	setTimeout("document.frames['NewsContent'].LayoutAndSetContent();",100);
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
new Form.Element.Observer($('NewsSmallPicFile'),1,pics_1);
	function pics_1()
		{
			if ($('NewsSmallPicFile').value=='')
			{
				$('pic_p_1').src='../Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_1').src=$('NewsSmallPicFile').value
			}
		} 
new Form.Element.Observer($('NewsPicFile'),1,pics_2);
	function pics_2()
		{
			if($('NewsPicFile').value=='')
			{
			$('pic_p_2').src='../Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_2').src=$('NewsPicFile').value
			}
		} 
function dels_1()
	{
		document.NewsForm.NewsSmallPicFile.value=''
	}
function dels_2()
	{
		document.NewsForm.NewsPicFile.value=''
	}
function ShowAdpicInfo()
{
	if (document.all.IsAdPic.checked==true)
    {
        document.all.selectAp.style.display="";
        document.all.Checkbox1.value="0";
        document.all.Checkbox2.value="0";
    }
    else
    {
        document.all.selectAp.style.display="none";
        document.all.wzPic.style.display="none";
        document.all.Checkbox2.checked=false;
        document.all.Checkbox1.checked=false;
        document.all.Checkbox1.value="0";
        document.all.Checkbox2.value="0";
        document.all.Adpic.style.display="none";
    }
}
function ShowAdpicInfo1()
{
	if (document.all.Checkbox1.checked==true)
    {   
        document.all.Checkbox1.value="1";
         document.all.Checkbox2.value="0";
        document.all.Adpic.style.display="";
        document.all.Checkbox2.checked=false;
        document.all.wzPic.style.display="none";
        document.all.IsAdPic.checked=true;
    }
    else
    {
        document.all.Checkbox1.value="0";
        document.all.Adpic.style.display="none";
    }
}
function ShowAdpicInfo2()
{
	if (document.all.Checkbox2.checked==true)
    {
        document.all.Checkbox2.value="1";        
        document.all.wzPic.style.display="";
        document.all.Checkbox1.checked=false;
         document.all.Checkbox1.value="0";
        document.all.Adpic.style.display="none";
        document.all.IsAdPic.checked=true
    }
    else
    {
        document.all.Checkbox2.value="0";
        document.all.wzPic.style.display="none";
    }
}


function Pic_Pic(obj)
{
	if(obj.width > 90 || obj.height > 90)
	{
		if(obj.width > obj.height)
		{
			obj.width=90;
		}
		else
		{
			obj.height=90;
		}
	}
}

function setPicToSmallPic(Pic)
{
	var _Pic = $(Pic);
	var _width = _Pic.width;
	var _height = _Pic.height;
	if (_width > 90 || _height > 90)
	{
		if (_width > _height)
		{
			_Pic.width = 90;
		}
		else
		{
			_Pic.height = 90
		}
	}
}
window.setTimeout("setPicToSmallPic('pic_p_1')",100);
window.setTimeout("setPicToSmallPic('pic_p_2')",100);
if (document.all.IsAdPic.checked==true)
{
	document.all.selectAp.style.display="";
}
else
{
	document.all.selectAp.style.display="none";
	document.all.wzPic.style.display="none";
	document.all.Checkbox1.value="0";
	document.all.Checkbox2.value="0";
	document.all.Adpic.style.display="none";
}
if (document.all.Checkbox1.checked==true)
{   
	document.all.Checkbox1.value="1";
	 document.all.Checkbox2.value="0";
	document.all.Adpic.style.display="";
	document.all.Checkbox2.checked=false;
	document.all.wzPic.style.display="none";
	document.all.IsAdPic.checked=true;
}
else
{
	document.all.Checkbox1.value="0";
	document.all.Adpic.style.display="none";
}
if (document.all.Checkbox2.checked==true)
{
	document.all.Checkbox2.value="1";        
	document.all.wzPic.style.display="";
	document.all.Checkbox1.checked=false;
	 document.all.Checkbox1.value="0";
	document.all.Adpic.style.display="none";
	document.all.IsAdPic.checked=true
}
else
{
	document.all.Checkbox2.value="0";
	document.all.wzPic.style.display="none";
}
</SCRIPT>