<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
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
Dim tempRootPath '模板根路径
FileName = Request.QueryString("File")
NewsTempletPath = Replace(Request.QueryString("Dir"),"..","")
If Not IsSelfRefer Then response.write "非法提交数据":Response.end
if Trim(NewsTempletPath)="" then Response.Write("错误参数"):response.end
if Lcase(instr(Lcase(NewsTempletPath),Lcase(G_TEMPLETS_DIR)))=0 then Response.Write("非法请求"):response.end
if G_VIRTUAL_ROOT_DIR = "" then:tempRootPath = Replace(Replace("/" &NewsTempletPath,"//","/"),"//","/"):else:tempRootPath =Replace(Replace( "/" & NewsTempletPath,"//","/"),"//","/"):end if
If IsValidStr(tempRootPath,"(\/){2,}|(\\)+") Then:Response.write "路径错误":Response.end
EditFile = Server.MapPath(Replace(Replace(tempRootPath&"/"&FileName,"//","/"),"//","/"))
Dim FsoObj,FileObj,FileStreamObj,strShowErr
Set FsoObj = Server.CreateObject(G_FS_FSO)
If Request.Form("Action")="Save" Then
	On Error Resume Next
	set FileObj= FsoObj.openTextFile(EditFile,2,1)
	FileObj.write Request.Form("FileContent")
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
		Application("Temp_Templet_Content")=Request.Form("FileContent")
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
      <td height="20" colspan="5"  class="xingmu"> <div align="left">文本编辑模板　<a href="../help?label=MF_Templet_Edit" target="_blank" style="cursor:help;'" class="sd"><img src="Images/_help.gif" width="50" height="17" border="0"></a>　　<a href="Templets_List.asp" class="sd">模板管理首页</a>　　　文件名：<%=Request.QueryString("File")%> 
          　　　路径:<%if Trim(Request.QueryString("Dir")) = "" then:Response.Write("根目录"):else:Response.Write Trim(Request.QueryString("Dir")):end if%></div></td>
    </tr>
    <tr class="hback"> 
      <td height="239" colspan="5" valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="2">
          <tr class="hback">
            <td height="33"><label>
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
			  %>
              </select>
              <input type="button" onClick="if(this.form.Lable_Name.options[this.form.Lable_Name.selectedIndex].value==''){return false;}else{insert(this.form.Lable_Name.options[this.form.Lable_Name.selectedIndex].value);}" value=" 插入标签 ">
              <input type="button" name="Submit3" value="选择标签" onClick="Insertlabel_News('../<%=G_ADMIN_DIR%>/label/Label_List.asp',400,350,'obj');">
            </label>              <label></label></td>
          </tr>
        </table>
        <textarea name="FileContent" rows="28" style="width:100%"><% = FileContent %></textarea>
      </td>
    </tr>
    <tr class="hback"> 
      <td colspan="5" height="39"> 
        <div align="left">
          <input type="submit" name="Submit" value="保存模板">
          <input type="reset" name="Submit2" value="恢复模板">
          <input name="Action" type="hidden" id="Action" value="Save">
          <input name="CodeTF" type="checkbox" id="CodeTF" value="1" style="display:none">
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function Insertlabel_News(URL,widthe,heighte,obj)
{
  var obj=window.OpenWindowAndSetValue("../Fs_Inc/convert.htm?"+URL,widthe,heighte,'window',obj)
  if (obj==undefined)return false;
  if (obj!='')insert(obj);
  }
function insert(returnValue_lable)
{
	obj=document.getElementById("FileContent");
	obj.focus();
	if(document.selection==null)
	{
		var iStart = obj.selectionStart
		var iEnd = obj.selectionEnd;
		obj.value = obj.value.substring(0, iEnd) +returnValue_lable+ obj.value.substring(iEnd, obj.value.length);
	}else
	{
		var range = document.selection.createRange();
		range.text=returnValue_lable;
	}
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






