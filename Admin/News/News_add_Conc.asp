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
	'session判断
	MF_Session_TF 
	'权限判断
	set Fs_news = new Cls_News
	Fs_News.GetSysParam()
	if Fs_News.isOpen=0 then
		Response.Redirect("lib/error.asp?ErrCodes=<li>新闻发布功能关闭!</li>&ErrorURL=")
		Response.End()
	end if
	If Not Fs_news.IsSelfRefer Then response.write "非法提交数据":Response.end
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
<title>管理___Powered by foosun Inc.</title>
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
			<td colspan="4" class="xingmu">添加
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
					类型 </div>
			</td>
			<td colspan="3" class="hback">
				<input name=NewsType type=radio id="NewsType" onClick="SwitchNewsType('TextNews');" value="TextNews" checked>
				普通
				<input name=NewsType type=radio id="NewsType" onClick="SwitchNewsType('PicNews');" value="PicNews">
				图片
				<input name=NewsType type=radio id="NewsType" onClick="SwitchNewsType('TitleNews');" value="TitleNews">
				标题 　&nbsp;&nbsp;　　
				<input name="isdraft" type="checkbox" id="isdraft" value="1">
				存到草稿箱中</td>
		</tr>
		<tr class="hback">
			<td width="12%" class="hback">
				<div align="right">
					<% = Fs_news.allInfotitle %>
					标题</div>
			</td>
			<td colspan="3" class="hback">
				<input name="NewsTitle" type="text" id="NewsTitle" size="40" maxlength="255">
				<input name="TitleColor" id="TitleColor" type="hidden">
				<img src="images/rectNoColor.gif" width="18" height="17" border=0 align="absmiddle" id="TitleColorShow" style="cursor:pointer;background-color:;" title="选取颜色!" onClick="GetColor(document.getElementById('TitleColorShow'),'TitleColor');">
				<input name="titleBorder" type="checkbox" id="titleBorder" value="1">
				粗体
				<input name="TitleItalic" type="checkbox" id="TitleItalic" value="1">
				斜体
				<input name="isShowReview" type="checkbox" id="isShowReview" value="1">
				评论连接 　权重
				<select name="PopID" id="PopID">
					<option value="5">总置顶</option>
					<option value="4">栏目置顶</option>
					<option value="0" selected>一般</option>
				</select>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback">
				<div align="right">选择栏目</div>
			</td>
			<td colspan="3" class="hback">
				<input name="ClassName" type="text" id="ClassName5" style="width:50%" value="<%=Fs_News.GetAdd_ClassName(str_ClassID)%>" readonly>
				<input name="ClassID" type="hidden" id="ClassID" value="<% = str_ClassID %>">
				<input type="button" name="Submit" value="选择栏目"   onClick="SelectClass();">
			</td>
		</tr>
		<tr class="hback" style="display:none">
			<td class="hback">
				<div align="right">选择专题</div>
			</td>
			<td colspan="3" class="hback">
				<input name="SpecialID" type="text" id="SpecialID" style="width:50%" readonly>
				<input name="SpecialID_EName" type="hidden" id="SpecialID_EName">
				<span class="tx"> </span>
				<input type="button" name="Submit" value="选择专题"   onClick="SelectSpecial();">
				<span class="tx">
				<input name="Submit" type="button" id="Submit" onClick="dospclear();" value="清除专题">
				</span> <span class="tx"> 可多选</span> &nbsp;&nbsp;<a href="../../help?Lable=NS_News_add_special" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr class="hback" id="str_URLAddress" style="display:none;">
			<td class="hback">
				<div align="right">连接地址 </div>
			</td>
			<td colspan="3" class="hback">
				<input name="URLAddress" type="text" id="URLAddress"  style="width:65%" maxlength="255"><font color="#FF0000">注:外部地址,请自行加入http://</font></td>
		</tr>
		<tr class="hback" id="str_CurtTitle" style="display:none">
			<td class="hback">
				<div align="right"> 副标题</div>
			</td>
			<td width="38%" class="hback">
				<input name="CurtTitle" type="text" id="CurtTitle" size="40" maxlength="255">
			</td>
			<td width="10%" class="hback">
				<div align="right">关键字</div>
			</td>
			<td width="40%" class="hback">
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
		<tr class="hback"  id="str_Templet">
			<td class="hback">
				<div align="right">模板地址</div>
			</td>
			<td colspan="3" class="hback">
				<input name="Templet" type="text" id="Templet" style="width:85%" value="<%=tmp_sTemplets%>" maxlength="255" readonly>
				<input name="Submit5" type="button" id="selNewsTemplet" value="选择模板"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.NewsForm.Templet);document.NewsForm.Templet.focus();">
			</td>
		</tr>
		<tr class="hback" id="str_NewsSmallPicFile" style="display:none">
			<td class="hback">
				<div align="right">图片(小)</div>
			</td>
			<td colspan="3" class="hback">
				<input name="NewsSmallPicFile" type="text" id="NewsSmallPicFile" style="width:85%" maxlength="255">
				<input type="button" name="PPPChoose"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.NewsForm.NewsSmallPicFile);">
			</td>
		</tr>
		<tr class="hback" id="str_NewsPicFile" style="display:none">
			<td class="hback">
				<div align="right">图片(大)</div>
			</td>
			<td colspan="3" class="hback">
				<input name="NewsPicFile" type="text" id="NewsPicFile" style="width:85%" maxlength="255">
				<input type="button" name="PPPChoose2"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.NewsForm.NewsPicFile);">
			</td>
		</tr>
		<tr class="hback" id="str_PicborderCss" style="display:none">
			<td class="hback">
				<div align="right">图片CSS</div>
			</td>
			<td colspan="3" class="hback">
				<select name="PicborderCss" id="PicborderCss">
					<option value="1" selected>图片CSS样式一</option>
					<option value="2">图片CSS样式二</option>
					<option value="3">图片CSS样式三</option>
					<option value="4">图片CSS样式四</option>
					<option value="5">图片CSS样式五</option>
				</select>
				图片前台显示CSS样式，方便加边框　　<a href="../../help?Lable=NS_News_add_PicCSS" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
		</tr>
		<tr class="hback" id="str_Author" style="display:none">
			<td class="hback">
				<div align="right">
					<% = Fs_news.allInfotitle %>
					作者</div>
			</td>
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
				<div align="right">
					<% = Fs_news.allInfotitle %>
					来源</div>
			</td>
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
		<tr class="hback" style="display:none">
			<td class="hback">
				<div align="right">
					<% = Fs_news.allInfotitle %>
					导读</div>
			</td>
			<td colspan="3" class="hback">
				<div align="left">
					<textarea name="NewsNaviContent" rows="6" id="NewsNaviContent" style="width:96%"></textarea>
				</div>
			</td>
		</tr>
		<tr class="hback">
			<td class="hback">
				<div align="right">类型</div>
			</td>
			<td colspan="3" class="hback">
				<div align="left">
					<input name="NewsProperty_Rec" type="checkbox" id="NewsProperty" value="1">
					推荐
					<input name="NewsProperty_mar" type="checkbox" id="NewsProperty" value="1" checked>
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
					幻灯</span></div>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" id="TodayNews" style="display:none;">
			<td colspan="4" class="hback">
				<table width="100%" border="0" cellspacing="1" cellpadding="2" class="table">
					<tr>
						<td height="26" align="center" width="120" class="xingmu">头条
							<% = Fs_news.allInfotitle %>
							类型：</td>
						<td height="26" class="hback">
							<input name="TodayNewsPicTF" value="" type="radio" checked onClick="if(this.checked){document.getElementById('TodayPicParam').style.display='none';}">
							文字头条
							<input name="TodayNewsPicTF" value="FoosunCMS" type="radio" onClick="if(this.checked){document.getElementById('TodayPicParam').style.display='';}">
							图片头条 　　<a href="../../help?Lable=NS_News_add_tt" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
					</tr>
					<tr id="TodayPicParam" style="display:none;">
						<td width="120" height="26" align="center"  class="xingmu">头条
							<% = Fs_news.allInfotitle %>
							参数：</td>
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
						  <img src="images/rect.gif" width="18" height="17" border=0 align=absmiddle id="FontBgColorShow" style="cursor:pointer;background-Color:;" title="选取颜色!" onClick="GetColor(this,'FontBgColor');"> <br>
						  &nbsp;&nbsp;图片头条标题：
						  <input name="PicTitle" type="text" id="PicTitle" size="40" maxlength="255">
						  &nbsp;&nbsp;图片宽度：
						  <input name="PicTitlewidth" type="text" id="PicTitlewidth" size="10" maxlength="10">
px </td>
					</tr>
				</table>
			</td>
		</tr>
		<tr  id="str_Content" style="display:" >
			<td height="303" class="hback">
				<div align="left">
					内容分页标签[FS:PAGE]<br>
                    <a href="javascript:void(0);" onClick="InsertHTML('[FS:PAGE]','NewsContent')"><span class="tx">插入分页标签</span></a><br>
                    <br>
              <input name="NewsProperty_Remote" type="checkbox" id="NewsProperty_Remote" value="1">
					远程存图 <br>
					<span class="tx">启用此功能后，如果从其它网站上复制到右边的编辑器中，并且中包含有图片，本系统会在保存文章时自动把相关图片复制到本站服务器上。<br>
					系统会因所下载图片的大小而影响速度，建议图片较多时不要使用此功能。</span> </div>
		  </td>
			<td height="303" colspan="3" class="hback">
				<!--编辑器开始-->
				<iframe id='NewsContent' src='../Editer/AdminEditer.asp?id=Content' frameborder=0 scrolling=no width='100%' height='480'></iframe>
				<input type="hidden" name="Content" value="">
                <!--编辑器结束-->
             </td>
		</tr>
		<tr class="hback" id="str_GroupID" style="display:none">
			<td class="hback">
				<div align="right">浏览点数</div>
			</td>
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
		<tr class="hback" id="str_FileName" style="display:none">
			<td class="hback">
				<div align="right">文件名</div>
			</td>
			<td class="hback">
				<%
	  Dim RoTF
	  if instr(Fs_News.strFileNameRule(Fs_News.fileNameRule,0,0),"自动编号ID")>0 then:RoTF="Readonly":End if
	  Response.Write"<input name=""FileName"" type=""text"" id=""FileName"" size=""40"" "& RoTF &" maxlength=""255"" value="""&Fs_News.strFileNameRule(Fs_News.fileNameRule,0,0)&""" title=""如果参数设置中设定为自动编号，将不能修改"">"
	  %>
			</td>
			<td class="hback">
				<div align="right">扩展名</div>
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
				<div align="right">添加日期</div>
			</td>
			<td class="hback">
				<input name="addtime" type="text" id="addtime" value="<% = now %>" size="40" maxlength="255">
			</td>
			<td class="hback">
				<div align="right">点击次数</div>
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
				<input type="button" name="Submit" value="确认保存<% = Fs_news.allInfotitle %>" onClick="SubmitFun(this.form);">
				<input name="d_Id" type="hidden" id="d_Id" value="<%=tmp_defineid%>">
				<input type="reset" name="Submit4" value="重新设置">
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
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
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
</SCRIPT>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






