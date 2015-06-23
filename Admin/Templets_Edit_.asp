<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/md5.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn
MF_Default_Conn
'得到session
MF_Session_TF
if not MF_Check_Pop_TF("MF_Templet") then Err_Show
if not MF_Check_Pop_TF("MF006") then Err_Show
if not MF_Check_Pop_TF("MF001") then Err_Show
Dim FileName,EditFile,FileContent
Dim NewsTempletpath,ParentPath,tmp_path
Dim tempRootPath,str_CurrPath,sRootDir

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
FileName = Request.QueryString("File")
NewsTempletPath = Replace(Request.QueryString("Dir"),"..","")
If Not IsSelfRefer Then response.write "非法提交数据":Response.end
if Trim(NewsTempletPath)="" then Response.Write("错误参数"):response.end
if Lcase(instr(Lcase(NewsTempletPath),Lcase(G_TEMPLETS_DIR)))=0 then Response.Write("非法请求"):response.end
if G_VIRTUAL_ROOT_DIR = "" then:tempRootPath = Replace(Replace("/" &NewsTempletPath,"//","/"),"//","/"):else:tempRootPath =Replace(Replace( "/" & NewsTempletPath,"//","/"),"//","/"):end if
If IsValidStr(tempRootPath,"(\/){2,}|(\\)+") Then:Response.write "路径错误":Response.end
EditFile = Server.MapPath(Replace(Replace(tempRootPath&"/"&FileName,"//","/"),"//","/"))
Dim FsoObj,FileObj,FileStreamObj,strShowErr,FileCont
Set FsoObj = Server.CreateObject(G_FS_FSO)
If Request.Form("Action")="Save" Then
	On Error Resume Next
	set FileObj= FsoObj.openTextFile(EditFile,2,1)
	FileCont=Request.Form("FileContent")
	FileObj.write FileCont
    FileObj.close
	set fileObj = nothing
	if Err.number<>0 then
		strShowErr = "<li>保存失败，可能是您没开启模板的写入权限.</li><li>可能是你服务器不支持FSO组件</li><li><a href=""javascript:history.back()"">继续编辑</a>&nbsp;&nbsp;</li>"
		Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Templets_List.asp?ShowPath="&Request.QueryString("Dir")&"")
		Response.end
	else
	    '更新缓存中的模板内容。Fsj 08.11.17
	    Application.Lock()
		Application("Temp_Templet_Path")=Replace(Replace(tempRootPath&"/"&FileName,"//","/"),"//","/")
		Application("Temp_Templet_Content")=FileCont
		Application.UnLock()
		
		strShowErr = "<li>保存模板成功.</li><li><a href=""javascript:history.back()"">继续编辑</a>&nbsp;&nbsp;</li>"
		Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=Templets_List.asp?ShowPath="&Request.QueryString("Dir")&"")
		Response.end
	end if
End IF
Set FileObj = FsoObj.GetFile(EditFile)
Set FileStreamObj = FileObj.OpenAsTextStream(1)
if Not FileStreamObj.AtEndOfStream then:FileContent = FileStreamObj.ReadAll:else:FileContent = "":end If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/Prototype.js"></script>
<script language="JavaScript" type="text/javascript" src="../FS_Inc/Get_Domain.asp"></script>
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY  LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table" >
	<form name="form1" method="post" action="">
		<tr class="hback">
			<td height="20" colspan="5"  class="xingmu">
				<div align="left">在线编辑模板　<a href="../help?label=MF_Templet_Edit" target="_blank" style="cursor:help;'" class="sd"><img src="Images/_help.gif" width="50" height="17" border="0"></a>　　<a href="Templets_List.asp" class="sd">模板管理首页</a>　　　文件名：<%=Request.QueryString("File")%> 　　　路径:
					<%if Trim(Request.QueryString("Dir")) = "" then:Response.Write("根目录"):else:Response.Write Trim(Request.QueryString("Dir")):end if%>
				</div>
			</td>
		</tr>
		<tr class="hback">
			<td colspan="5" valign="top">
				<table width="100%" border="0" cellspacing="0" cellpadding="2">
					<tr class="hback">
						<td height="33">
							<label>
							<select name="Lable_Name" id="Lable_Name">
								<option value="">==选择标签==</option>			
			 <%
			  dim rs
			  set rs = Conn.execute("select LableName From FS_MF_Lable where isDel=0 order by id desc")
			  do while not rs.eof
			  	response.Write "<option value="""&rs("LableName")&""">┄"&replace(replace(rs("LableName"),"{FS400_",""),"}","")&"</option>"
			  rs.movenext
			  loop
			  rs.close:set rs = nothing
			  %>	</select>
							<input type="button" onClick="Insertlabel_Sel($('Lable_Name'));" value=" 插入标签 ">
							<input type="button" name="Button" value="选择标签" onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/Label_List.asp',400,350,'obj');">
							</label>
						</td>
					</tr>
				</table>
				
                <!--编辑器开始-->
				<iframe id='NewsContent' src='Editer/AdminEditer.asp?id=FileContent' frameborder=0 scrolling=no width='100%' height='480'></iframe>
				<input type="hidden" name="FileContent" value="<%=HandleEditorContent(FileContent)%>">
                <!--编辑器结束-->		
			</td>
		</tr>
		<tr class="hback">
			<td colspan="5" height="39">
				<div align="left">
					<input type="button" name="Submit" onClick="CheckForm(this.form);" value="保存模板">
					<input type="reset" name="Submit" value="恢复模板"  onclick="restValue()">
					<input name="Action" type="hidden" id="Action" value="Save">
					<input name="CodeTF" type="checkbox" id="CodeTF" value="1" style="display:none">
				</div>
			</td>
		</tr>
	</form>
</table>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function CheckForm(FormObj)
{
	if (frames["NewsContent"].g_currmode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
	FormObj.FileContent.value=frames["NewsContent"].GetNewsContentArray();
	FormObj.submit();
}

function Insertlabel_News(URL,widthe,heighte,obj)
{
  var obj=window.OpenWindowAndSetValue("../Fs_Inc/convert.htm?"+URL,widthe,heighte,'window',obj);
  if (obj==undefined)return false;
  if (obj!='')InsertEditor(obj);
	
}
function restValue()
{
   this.location.reload();
}
function Insertlabel_Sel(Lable_obj)
{
	if(Lable_obj.options[Lable_obj.selectedIndex].value==''){
	return false;
	}else{
	InsertEditor(Lable_obj.options[Lable_obj.selectedIndex].value);
	}
}
function InsertEditor(InsertValue)
{
	InsertHTML(InsertValue,"NewsContent");
}
function opencat(cat)
{
  if(cat.style.display=="none"){
     cat.style.display="";
  } else {
     cat.style.display="none"; 
  }
}
</script>
<%
Set FsoObj = Nothing
Set FileObj = Nothing
Set FileStreamObj = Nothing
%>






