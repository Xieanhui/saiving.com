<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<%
	Dim Conn,User_Conn,sRootDir,tmp_sFileExtName,tmp_sTemplets
	Dim Fs_news,str_ClassID,news_SQL,obj_news_rs,icNum,isUrlStr,str_Href,str_Href_title,obj_news_rs_1,str_action,obj_cnews_rs,news_count,str_CurrPath,IsAdPic,AdPicWH,AdPicLink,AdPicAdress
	Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
	Temp_Admin_Name = Session("Admin_Name")
	Temp_Admin_Is_Super = Session("Admin_Is_Super")
	Temp_Admin_FilesTF = Session("Admin_FilesTF")
	MF_Default_Conn
	MF_User_Conn
	MF_GetUserGroupID
	MF_Session_TF 
	if not MF_Check_Pop_TF("NS013") then
	if not MF_Check_Pop_TF("NS001") then Err_Show
	end if
	set Fs_news = new Cls_News
	Fs_News.GetSysParam()
	if Fs_News.isOpen=0 then
		Response.Redirect("lib/error.asp?ErrCodes=<li>新闻发布功能关闭!</li>&ErrorURL=")
		Response.End()
	end if
	str_ClassID = NoSqlHack(Request.QueryString("ClassID"))
	if str_ClassId<>"" then
		if not Get_SubPop_TF(str_ClassID,"NS001","NS","news") then Err_Show
		Dim GetClassAdPicInfoRs,GetClassAdPicInfoSql
		GetClassAdPicInfoSql="Select IsAdPic,AdPicWH,AdPicLink,AdPicAdress From FS_NS_NewsClass Where ClassID='"&str_ClassId&"'"
		Set GetClassAdPicInfoRs=Conn.execute(GetClassAdPicInfoSql)
		If Not GetClassAdPicInfoRs.Eof Then
			IsAdPic=GetClassAdPicInfoRs("IsAdPic")
			AdPicWH=GetClassAdPicInfoRs("AdPicWH")
			AdPicLink=GetClassAdPicInfoRs("AdPicLink")
			AdPicAdress=GetClassAdPicInfoRs("AdPicAdress")			
		End If
		GetClassAdPicInfoRs.Close
		Set GetClassAdPicInfoSql=Nothing
	end if
	if Trim(str_ClassID)<>"" then
		Dim tmp_class_obj,tmp_defineid
		set tmp_class_obj = conn.execute("select FileExtName,NewsTemplet,DefineID from FS_NS_NewsClass where classID='"& str_ClassID &"'")
		if tmp_class_obj.eof then
			tmp_class_obj.close:set tmp_class_obj = nothing
			response.Write "错误的参数,找不到栏目"
			Response.end
		end if
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
		set tmp_class_obj = nothing
	Else
		tmp_defineid = 0
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
	'获取辅助字段信息,保存到数组CustColumnArr中if not isnull(trim(tmp_defineid)) or trim(tmp_defineid)>0 then
		Dim CustColumnRs,CustSql,CustColumnArr
		CustSql="select DefineID,ClassID,D_Name,D_Coul,D_Type,D_isNull,D_Value,D_Content,D_SubType from [FS_MF_DefineTable] Where D_SubType='NS' and  Classid="& tmp_defineid &""
		Set CustColumnRs=CreateObject(G_FS_RS)
		CustColumnRs.Open CustSql,Conn,1,3
		If Not CustColumnRs.Eof Then
			CustColumnArr=CustColumnRs.GetRows()
		End If
		CustColumnRs.close:Set CustColumnRs = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="js/Public.js"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/CheckJs.js" type="text/JavaScript"></script>
<script language="JavaScript" type="text/javascript" src="../../FS_Inc/Get_Domain.asp"></script>
</head>
<body>
<iframe width="260" height="165" id="colorPalette" src="lib/selcolor.htm" style="visibility:hidden; position: absolute;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="1"></td>
	</tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<tr>
		<td class="hback"><a href="News_Manage.asp">栏目管理</a>┆<a href="News_Manage.asp?ClassId=<%=Request.QueryString("ClassId")%>">返回本类栏目管理</a></td>
	</tr>
</table>
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
	<form action="News_Save.asp" name="NewsForm" method="post">
		<tr class="xingmu">
			<td colspan="4" class="xingmu">添加
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
					类型 </div>			</td>
			<td colspan="3" class="hback">
				<input name=NewsType type=radio onClick="SwitchNewsType('TextNews');" value="TextNews" checked>
				普通
				<input name=NewsType type=radio onClick="SwitchNewsType('PicNews');" value="PicNews">
				图片
				<input name=NewsType type=radio onClick="SwitchNewsType('TitleNews');" value="TitleNews">
				标题 　&nbsp;&nbsp;　
				<input name="isdraft" type="checkbox" id="isdraft" value="1">
				存到草稿箱中</td>
		</tr>
		<tr >
			<td width="12%" class="hback">
				<div align="right"> 标题</div>			</td>
			<td colspan="3" class="hback">
				<input name="NewsTitle" type="text" id="NewsTitle" size="40" maxlength="255" onFocus="Do.these('NewsTitle',function(){return isEmpty('NewsTitle','span_NewsTitle')})" onKeyUp="Do.these('NewsTitle',function(){return isEmpty('NewsTitle','span_NewsTitle')})" style="background-image:url(../Images/bg.gif);">
				<span id="span_NewsTitle"></span>
				<input name="TitleColor" id="TitleColor" type="hidden">
				<img src="images/rectNoColor.gif" width="18" height="17" border=0 align="absmiddle" id="TitleColorShow" style="cursor:pointer;background-color:;" title="选取颜色!" onClick="GetColor(document.getElementById('TitleColorShow'),'TitleColor');">
				<input name="titleBorder" type="checkbox" id="titleBorder" value="1">
				粗体
				<input name="TitleItalic" type="checkbox" id="TitleItalic" value="1">
				斜体
				<input name="isShowReview" type="checkbox" id="isShowReview" value="1">
				评论连接　　权重
				<select name="PopID" id="PopID">
					<option value="5">总置顶</option>
					<option value="4">栏目置顶</option>
					<option value="0" selected>一般</option>
				</select>			</td>
		</tr>
		<tr>
			<td class="hback">
				<div align="right">副标题 </div></td>
			<td colspan="3" class="hback"><input name="CurtTitle" type="text" id="CurtTitle" style="width:85%;" maxlength="255"></td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right">选择栏目</div>			</td>
			<td colspan="3" class="hback">
				<input name="ClassName" type="text" id="ClassName" style="width:50%" value="<%=Fs_News.GetAdd_ClassName(str_ClassID)%>" readonly onFocus="Do.these('ClassName',function(){return isEmpty('ClassName','span_ClassName')})" onKeyUp="Do.these('ClassName',function(){return isEmpty('ClassName','span_ClassName')})">
				<input name="ClassID" type="hidden" id="ClassID" value="<% = str_ClassID %>">
				<input type="button" name="Submit" value="选择栏目"   onClick="SelectClass();">
				<input type="button" name="Submit2" value="添加栏目" onClick="window.location.href='Class_add.asp?ClassID=<%=str_ClassID %>&Action=add'">
				<span id="span_ClassName"></span> </td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right">选择专题</div>			</td>
			<td colspan="3" class="hback">
				<input name="SpecialID" type="text" id="SpecialID" style="width:50%" readonly>
				<input name="SpecialID_EName" type="hidden" id="SpecialID_EName">
				<span class="tx"> </span>
				<input type="button" name="Submit" value="选择专题"   onClick="SelectSpecial();">
				<span class="tx">
				<input name="Submit" type="button" id="Submit" onClick="dospclear();" value="清除专题">
				</span> <span class="tx"> 可多选</span> &nbsp;&nbsp;<a href="../../help?Lable=NS_News_add_special" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr  id="str_URLAddress" style="display:none;" >
			<td class="hback">
				<div align="right">连接地址 </div>			</td>
			<td colspan="3" class="hback">
				<input name="URLAddress" type="text" id="URLAddress"  style="width:65%" maxlength="255"><font color="#FF0000">注:外部地址,请自行加入http://</font></td>
		</tr>
		<tr  id="str_CurtTitle" style="display:" >
			<td class="hback">
				<div align="right">关键字</div>			</td>
			<td colspan="3" class="hback">
				<input name="KeywordText" type="text" id="KeywordText" size="15" maxlength="255">
				<input name="KeyWords" type="hidden" id="KeyWords">
				<select name="selectKeywords" id="selectKeywords" style="width:120px" onChange=Dokesite(this.options[this.selectedIndex].value)>
					<option value="" selected>选择关键字</option>
					<option value="Clean" style="color:red">清空</option>
					<%=Fs_news.GetKeywordslist("",1)%>
				</select>
				<input name="KeywordSaveTF" type="checkbox" id="KeywordSaveTF" value="1">
				记忆</td>
		</tr>
		<tr   id="str_Templet" style="display:" >
			<td class="hback">
				<div align="right">模板地址</div>			</td>
			<td colspan="3" class="hback">
				<input name="Templet" type="text" id="Templet" style="width:85%" value="<%=tmp_sTemplets%>" maxlength="255" readonly>
				<input name="Submit5" type="button" id="selNewsTemplet" value="选择模板"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.NewsForm.Templet);document.NewsForm.Templet.focus();">			</td>
		</tr>
		<tr   id="str_NewsSmallPicFile" style="display:none">
			<td class="hback">
				<div align="right">选择图片地址</div>			</td>
			<td colspan="3" class="hback">
				<table width="417" border="0" cellspacing="1" cellpadding="5">
					<tr>
						<td width="50%">
							<table width="10" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
								<tr>
									<td class="hback"><img src="../Images/nopic_supply.gif" id="pic_p_1" onload="Javascript:if(this.width > 90 || this.height > 90){if(this.width > this.height){this.width=90;}else{this.height=90;}}" /></td>
								</tr>
							</table>
							<div align="center">
								<input name="NewsSmallPicFile" type="hidden" id="NewsSmallPicFile" style="width:85%" maxlength="255">
								<img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>',500,320,window,document.NewsForm.NewsSmallPicFile);" style="cursor:hand;"> 　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_1();" style="cursor:hand;"> </div>						</td>
						<td width="50%">
							<table width="10" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
								<tr>
									<td class="hback"><img src="../Images/nopic_supply.gif" id="pic_p_2" onload="Javascript:if(this.width > 90 || this.height > 90){if(this.width > this.height){this.width=90;}else{this.height=90;}}" /></td>
								</tr>
							</table>
							<div align="center">
								<input name="NewsPicFile" type="hidden" id="NewsPicFile" style="width:85%" maxlength="255">
								<img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<% = str_CurrPath %>',500,320,window,document.NewsForm.NewsPicFile);" style="cursor:hand;"> 　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_2();" style="cursor:hand;"> </div>						</td>
					</tr>
					<tr>
						<td class="hback">
							<div align="center">小图地址</div>						</td>
						<td class="hback">
							<div align="center">大图地址</div>						</td>
					</tr>
				</table>			</td>
		</tr>
		<tr  id="str_Author" style="display:" >
			<td class="hback">
				<div align="right"> 作者</div>			</td>
			<td class="hback">
				<input name="Author" type="text" id="Author" size="15" maxlength="255">
				<select name="selectAuthor" id="selectAuthor" style="width:120px"  onChange="document.NewsForm.Author.value=this.options[this.selectedIndex].text;">
					<option style="color:red"> </option>
					<option value="佚名">佚名</option>
					<option value="本站">本站</option>
					<option value="未知">未知</option>
					<%=Fs_news.GetKeywordslist("",3)%>
				</select>
				<input name="AuthorSaveTF" type="checkbox" id="AuthorSaveTF" value="1">
				记忆</td>
			<td class="hback">
				<div align="right"> 来源</div>			</td>
			<td class="hback">
				<input name="Source" type="text" id="Source" size="15" maxlength="255">
				<select name="selectSource" id="selectSource" style="width:120px"  onChange="document.NewsForm.Source.value=this.options[this.selectedIndex].text;">
					<option value="" selected> </option>
					<option value="本站原创">本站原创</option>
					<option value="不详">不详</option>
					<%=Fs_news.GetKeywordslist("",2)%>
				</select>
				<input name="SourceSaveTF" type="checkbox" id="SourceSaveTF" value="1">
				记忆</td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right"> 导读</div>			</td>
			<td colspan="3" class="hback">
				<div align="left">
					<textarea name="NewsNaviContent" rows="6" id="NewsNaviContent" style="width:96%"></textarea>
				</div>			</td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right">类型</div>			</td>
			<td colspan="3" class="hback">
				<div align="left">
					<input name="NewsProperty_Rec" type="checkbox" id="NewsProperty" value="1">
					推荐
					<input name="NewsProperty_mar" type="checkbox" id="NewsProperty" value="1">
					滚动
					<input name="NewsProperty_rev" type="checkbox" id="NewsProperty" value="1" checked>
					允许评论
					<input name="NewsProperty_constr" type="checkbox" id="NewsProperty" value="1">
					投稿
					<input name="NewsProperty_tt" type="checkbox" id="NewsProperty" value="1"  onClick="ChooseTodayNewsType();">
					头条
					<input name="NewsProperty_hots" type="checkbox" id="NewsProperty" value="1" disabled="disabled">
					热点
					<input name="NewsProperty_jc" type="checkbox" id="NewsProperty" value="1">
					精彩
					<input name="NewsProperty_unr" type="checkbox" id="NewsProperty" value="1">
					不规则
					<input name="NewsProperty_ann" type="checkbox" id="NewsProperty" value="1">
					公告 <span id="str_filt" style="display:none">
					<input name="NewsProperty_filt" type="checkbox" id="NewsProperty" value="1">
					幻灯</span></div>			</td>
		</tr>
		<tr  id="TodayNews" style="display:none;" >
			<td colspan="4" class="hback">
				<table width="100%" border="0" cellspacing="1" cellpadding="2" class="table">
					<tr>
						<td height="26" align="center" width="120" class="xingmu">头条类型：</td>
						<td height="26" class="hback">
							<input name="TodayNewsPicTF" value="" type="radio" checked onClick="if(this.checked){document.getElementById('TodayPicParam').style.display='none';}">
							文字头条
							<input name="TodayNewsPicTF" value="FoosunCMS" type="radio" onClick="if(this.checked){document.getElementById('TodayPicParam').style.display='';}">
							图片头条 　　<a href="../../help?Lable=NS_News_add_tt" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
					</tr>
					<tr id="TodayPicParam" style="display:none;" >
						<td width="120" height="26" align="center"  class="xingmu">头条参数：</td>
					  <td height="26" class="hback">&nbsp;&nbsp;字体
							<SELECT name="FontFace" id="FontFace">
								<option value="宋体">宋体</option>
								<option value="楷体_GB2312">楷体</option>
								<option value="新宋体">新宋体</option>
								<option value="黑体">黑体</option>
								<option value="隶书">隶书</option>
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
							<img src="images/rect.gif" width="18" height="17" border=0 align=absmiddle id="FontColorShow" style="cursor:pointer;background-Color:#000000;" title="选取颜色!" onClick="GetColor(this,'FontColor');"> 　　 字体间距：
							<INPUT TYPE="text" maxlength="3" NAME="FontSpace" size=3 value="12">
							px 图片背景色
							<input type="text" name="FontBgColor" maxlength=7 size=7 id="FontBgColor" value="FFFFFF">
						  <img src="images/rect.gif" width="18" height="17" border=0 align=absmiddle id="FontBgColorShow" style="cursor:pointer;background-Color:;" title="选取颜色!" onClick="GetColor(this,'FontBgColor');"><br>
						  &nbsp;&nbsp;图片头条标题：
						  <input name="PicTitle" type="text" id="PicTitle" size="40" maxlength="255">
						  &nbsp;&nbsp;图片宽度：
						  <input name="PicTitlewidth" type="text" id="PicTitlewidth" size="10" maxlength="10">
px </td>
					</tr>
				</table>			</td>
		</tr>
		<tr id="str_Content">
			<td class="hback">
				<div align="left">内容分页标签[FS:PAGE]<br>
				  <a href="javascript:void(0);" onClick="InsertHTML('[FS:PAGE]','NewsContent')"><span class="tx">插入分页标签</span></a><br>
			      <br>
        <input name="NewsProperty_Remote" type="checkbox" id="NewsProperty_Remote" value="1">
					远程存图 <br>
					<span class="tx">启用此功能后，如果从其它网站上复制到右边的编辑器中，并且中包含有图片，本系统会在保存文章时自动把相关图片复制到本站服务器上。<br>
					系统会因所下载图片的大小而影响速度，建议图片较多时不要使用此功能。</span> </div>			</td>
			<td colspan="3" class="hback">
                <!--编辑器开始-->
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=Content' frameborder=0 scrolling=no width='100%' height='480'></iframe>
				<input type="hidden" name="Content" value="">
                <!--编辑器结束-->
            </td>
		</tr>
		<!--自定义自段开始-->
		<%If IsArray(CustColumnArr) Then
			response.Write"<tr><td colspan=""4"" class=""hback_1""><b>自定义开始</b></td></tr>"
			Dim InputModeStr,AuxiInfoList,AuxiListArr,k,tmp_k,i,tmp_nulls_span,tmp_nulls
			For i = 0 to Ubound(CustColumnArr,2)
				if CustColumnArr(5,i)=0 then
					tmp_nulls="onFocus=""Do.these('FS_NS_Define_"&CustColumnArr(3,i)&"',function(){return isEmpty('FS_NS_Define_"&CustColumnArr(3,i)&"','span_FS_NS_Define_"&CustColumnArr(3,i)&"')})"" onKeyUp=""Do.these('FS_NS_Define_"&CustColumnArr(3,i)&"',function(){return isEmpty('FS_NS_Define_"&CustColumnArr(3,i)&"','span_FS_NS_Define_"&CustColumnArr(3,i)&"')})"""
					tmp_nulls_span="&nbsp;<span id=""span_FS_NS_Define_"&CustColumnArr(3,i)&"""></span>"
				else
					tmp_nulls=""
					tmp_nulls_span=""
				end if
				'0:单行文本,1:多行文本,4:下拉列表,5:数字类型,6:日期类型,7:图片类型,8:附件类型,9:邮件类型,10:多行文本带编辑器
				Select Case CStr(CustColumnArr(4,i))	'根据选择的输入类型生成输入方式
					Case 0
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><input name=""FS_NS_Define_"&CustColumnArr(3,i)&""" type=""test"" style=""width:70%""  value="""&CustColumnArr(6,i)&""" "& tmp_nulls &">"& tmp_nulls_span &"&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
					case 1
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><textarea rows=""4"" name=""FS_NS_Define_"&CustColumnArr(3,i)&""" style=""width:70%"" "& tmp_nulls &">"&CustColumnArr(6,i)&"</textarea>"& tmp_nulls_span &"&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
					Case 4
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><Select Name=""FS_NS_Define_"&CustColumnArr(3,i)&""" style=""width:70%"">"&vbcrlf
							AuxiListArr=Split(CustColumnArr(6,i),vbcrlf)
							For tmp_k = 0 to UBound(AuxiListArr)	'读辅助字典的选项信息
								If AuxiListArr(tmp_k)<>"" Then 
									Response.Write"<Option value="""&AuxiListArr(tmp_k)&""">"&AuxiListArr(tmp_k)&"</option>"&vbcrlf
								End if
							Next
							Response.Write "</Select>&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
					Case 7
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><input name=""FS_NS_Define_"&CustColumnArr(3,i)&""" type=""test"" style=""width:70%""  value="""&CustColumnArr(6,i)&""" "& tmp_nulls &"> <input name=""SelectAdPic"" type=""button"" id=""SelectAdPic"" onClick=""OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath="&str_CurrPath&"',500,300,window,document.NewsForm.FS_NS_Define_"&CustColumnArr(3,i)&");""  value=""选择图片""> "& tmp_nulls_span &"&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
					Case 8
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><input name=""FS_NS_Define_"&CustColumnArr(3,i)&""" type=""test"" style=""width:70%""  value="""&CustColumnArr(6,i)&""" "& tmp_nulls &"> <input name=""SelectAdPic"" type=""button"" id=""SelectAdPic"" onClick=""OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath="&str_CurrPath&"',500,300,window,document.NewsForm.FS_NS_Define_"&CustColumnArr(3,i)&");""  value=""选择附件""> "& tmp_nulls_span &"&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
					Case 9
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><input name=""FS_NS_Define_"&CustColumnArr(3,i)&""" type=""test"" style=""width:70%""  value=""mailto:"&CustColumnArr(6,i)&""" "& tmp_nulls &">"& tmp_nulls_span &"&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
					Case 10
							dim scriptSTR,DefineFieldTextAreaID
							DefineFieldTextAreaID = CustColumnArr(3,i)
							scriptSTR = "<iframe id='DefineFieldContent' src='../Editer/AdminEditer.asp?id=FS_NS_Define_" & DefineFieldTextAreaID & "' frameborder=0 scrolling=no width='560' height='300'></iframe>"
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><textarea name=""FS_NS_Define_" & DefineFieldTextAreaID & """ id=""FS_NS_Define_" & DefineFieldTextAreaID & """ style=""display: none"">" & HandleEditorContent(CustColumnArr(6,i)) & "</textarea>"&scriptSTR&" "& tmp_nulls_span &"&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
					Case Else	'单行，数字，日期
							Response.Write "<tr ><td class=""hback"" align=""right"">"&CustColumnArr(2,i)&"</td><td colspan=""3"" class=""hback""><input name=""FS_NS_Define_"&CustColumnArr(3,i)&""" type=""test"" style=""width:70%""  value="""&CustColumnArr(6,i)&""" "& tmp_nulls &">"& tmp_nulls_span &"&nbsp;<span class=""tx"">"&CustColumnArr(7,i)&"</span></td></tr>"&vbcrlf
				End Select
			Next
			response.Write"<tr><td colspan=""4"" class=""hback_1""><b>自定义结束</b></td></tr>"
		End If
	%>
     
		<!--自定义自段结束-->
		<tr  id="str_GroupID" style="display:" >
			<td class="hback">
				<div align="right">浏览点数</div>			</td>
			<td colspan="3" class="hback">
				<input name="PointNumber" type="text" id="PointNumber2" size="16"  onChange="ChooseExeName();">
				金币
				<input name="Money" type="text" id="Money2" size="16"  onChange="ChooseExeName();">
				浏览权限
				<input name="BrowPop"  id="BrowPop" type="text" onMouseOver="this.title=this.value;" readonly>
				<select name="selectPop" id="selectPop" style="overflow:hidden;" onChange="ChooseExeName();">
					<option value="" selected>选择会员组</option>
					<option value="del" style="color:red;">清空</option>
					<% = MF_GetUserGroupID %>
				</select>
				<a href="../../help?Lable=NS_News_add_pop" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr id="str_FileName" style="display:" >
			<td class="hback">
				<div align="right">文件名</div>			</td>
			<td class="hback">
				<%
	  Dim RoTF
	  if instr(Fs_News.strFileNameRule(Fs_News.fileNameRule,0,0),"自动编号ID")>0 or instr(Fs_News.strFileNameRule(Fs_News.fileNameRule,0,0),"唯一NewsID") then:RoTF="Readonly":End if
	  Response.Write"<input name=""FileName"" type=""text"" id=""FileName"" size=""40"" "& RoTF &" maxlength=""255"" value="""&Fs_News.strFileNameRule(Fs_News.fileNameRule,0,0)&""" title=""如果参数设置中设定为自动编号，将不能修改"">"
	  %>			</td>
			<td class="hback">
				<div align="right">扩展名</div>			</td>
			<td class="hback">
				<select name="FileExtName" id="FileExtName">
					<option value="html" <%if tmp_sFileExtName = 0 then response.Write("selected")%>>.html</option>
					<option value="htm" <%if tmp_sFileExtName = 1 then response.Write("selected")%>>.htm</option>
					<option value="shtml" <%if tmp_sFileExtName = 2 then response.Write("selected")%>>.shtml</option>
					<option value="shtm" <%if tmp_sFileExtName = 3 then response.Write("selected")%>>.shtm</option>
					<option value="asp" <%if tmp_sFileExtName = 4 then response.Write("selected")%>>.asp</option>
				</select>			</td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right">添加日期</div>			</td>
			<td class="hback">
				<input name="addtime" type="text" id="addtime" value="<% = now %>" size="40" maxlength="255">			</td>
			<td class="hback">
				<div align="right">点击次数</div>			</td>
			<td class="hback">
				<input name="Hits" type="text" id="Hits" value="0" size="20" onFocus="Do.these('Hits',function(){return isNumber('Hits','span_Hits','请填写正确的格式',true)})" onKeyUp="Do.these('Hits',function(){return isNumber('Hits','span_Hits','请填写正确的格式',true)})">
				<span id="span_Hits"></span></td>
		</tr>
		<tr id="IsShowAdpic">
			<td class="hback"><div align="right">是否显示画中画</div></td>
			<td colspan="3" class="hback"><input name="IsAdPic" type="checkbox" id="IsAdPic" value="1" onClick="javascript:ShowAdpicInfo();" <%If IsAdPic="1" or IsAdPic="2" Then Response.Write("checked")%>></td>
		</tr>
		<tr id="selectAp" style="display:none" class="hback">
		<td class="hback"></td>
		    <td  colspan="2" class="hback" align="left"> 图片画中画
		
                <input id="Checkbox1" name="Checkbox1" type="checkbox" onClick="javascript:ShowAdpicInfo1();" <% If Cint(IsAdPic)="1" Then Response.Write("checked") %>/> &nbsp;&nbsp;&nbsp;文字画中画
		     
                <input id="Checkbox2" name="Checkbox2" type="checkbox"  onClick="javascript:ShowAdpicInfo2();" <% If Cint(IsAdPic)="2" Then Response.Write("checked") %>/>
             </td>
             		<td class="hback"></td>

		</tr>
		<tr id="Adpic" style="display:none" class="hback">
			<td class="hback" colspan="4"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
              <tr>
                <td width="12%" height="2" class="hback"><div align="right">画中画参数设置</div></td>
                <td width="88%" height="2" class="hback"><input name="AdPicWH" type="text" id="AdPicWH" size="20" maxlength="20" value="<%if Cint(IsAdPic)=1 then response.write(AdPicWH) else response.Write("100,100,1,400") end if%>">
                (宽度,高度,左(1)右(0),插入位置在新闻内容前多少字(正整数)：如100,100,1,400)</td>
              </tr>
              <tr>
                <td height="5" class="hback"><div align="right">图片地址</div></td>
                <td height="5" class="hback"><input name="AdPicAdress" type="text" id="AdPicAdress"  size="20" maxlength="250" readonly value="<%=AdPicAdress%>">
                    <input name="SelectAdPic" type="button" id="SelectAdPic" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath%>',500,300,window,document.NewsForm.AdPicAdress);"  value="选择图片或FLASH">
                  链接地址
                  <input name="AdPicLink" type="text" id="AdPicLink"  size="36" maxlength="250" value="<%if Cint(IsAdPic)=1 then response.write(AdPicLink) end if%>"></td>
              </tr>
            </table></td>
		</tr>
		<tr id="wzPic" style="display:none" class="hback">
		         <td colspan="4">
		        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
		                  <tr>
                <td width="12%" height="2" class="hback"><div align="right">画中画设置</div></td>
                <td width="88%" height="2" class="hback"><input name="AdPicWHw" type="text" id="Text2" size="20" maxlength="20" value="<%if Cint(IsAdPic)=2 then response.write(AdPicWH) end if%>">
                (插入位置在新闻内容前多少字(正整数)：如100</td>
              </tr>
              <tr>
	         <td class="hback" align="right">画中画代码
		     </td>
		     <td class="hback" colspan="3"  align="left">
                <textarea id="IsApicArea" name="IsApicArea" cols="80" rows="10"><%if Cint(IsAdPic)=2 then response.write(AdPicLink) end if%></textarea>
		      </td>
		     </tr>
		     </table>
		    </td>
		</tr>
		<tr >
			<td class="hback">
				<div align="right"></div>			</td>
			<td colspan="3" class="hback"><input type="button" name="Submit3" value="确认保存<% = Fs_news.allInfotitle %>"  onClick="SubmitFun(this.form);">
			  <input type="reset" name="Submit4" value="重新设置">
              <input name="News_Action" type="hidden" id="News_Action" value="add_Save">
              <input name="d_Id" type="hidden" id="d_Id" value="<%=tmp_defineid%>"></td>
		</tr>
	</form>
</table>
</body>
</html>
<%
If IsAdPic="1" Then Response.Write("<script language=""javascript"">document.getElementById('Adpic').style.display="""";</script>")
set tmp_class_obj = nothing
set tmp_class_obj = nothing
set Fs_news = nothing
%><script language="JavaScript" type="text/JavaScript">
function SubmitFun(FormObj)
{
	if(FormObj.NewsTitle.value == "")
	{
		alert("请填写标题！");
		FormObj.NewsTitle.focus();
		return false;
	}
	if(FormObj.ClassName.value=="")
	{
		alert("请选择栏目！");
		FormObj.ClassName.focus();
		return false;
	}
	if(FormObj.ClassID.value=="")
	{
		alert("栏目参数错误！");
		FormObj.ClassName.focus();
		return false;
	}
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
	FormObj.Content.value=frames["NewsContent"].GetNewsContentArray();
	if(frames["DefineFieldContent"])FormObj.FS_NS_Define_<% = DefineFieldTextAreaID %>.value=frames["DefineFieldContent"].GetNewsContentArray();
	FormObj.submit();
}

function SwitchNewsType(NewsType)
{
	switch (NewsType)
	{
	case "TitleNews":
		document.getElementById('str_UrLaddress').style.display='';
		document.getElementById('str_CurtTitle').style.display='';
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
   CheckNumber(document.NewsForm.PointNumber,"浏览扣点值");
  if (document.NewsForm.PointNumber.value>32767||document.NewsForm.PointNumber.value<-32768||document.NewsForm.PointNumber.value=='0')
	{
		alert('浏览扣点值超过允许范围！\n最大32767，且不能为0');
		document.NewsForm.PointNumber.value='';
		document.NewsForm.PointNumber.focus();
	}
   CheckNumber(document.NewsForm.Money,"浏览金币值");
  if (document.NewsForm.Money.value>32767||document.NewsForm.Money.value<-32768||document.NewsForm.Money.value=='0')
	{
		alert('浏览金币值超过允许范围！\n最大32767，且不能为0');
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
//加载时，自动显示画中画
//2009-5-26 by SamJun
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

</SCRIPT><!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->