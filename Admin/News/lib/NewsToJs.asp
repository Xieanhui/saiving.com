<% Option Explicit %>
<!--#include file="../../../FS_Inc/Const.asp" -->
<!--#include file="../../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../../FS_InterFace/CLS_Foosun.asp" -->
<!--#include file="../../../FS_Inc/md5.asp" -->
<!--#include file="../../../FS_Inc/Function.asp" -->
<!--#include file="cls_main.asp" -->
<!--#include file="Cls_Js.asp"-->
<%
Dim Conn,str_CurrPath,Types,NewsID,RsNewsObj,sRootDir
MF_Default_Conn
MF_Session_TF
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
'权限判断
'Call MF_Check_Pop_TF("NS_Class_000001") 
'得到会员组列表 
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
'判断用户是否为超级管理员，限定访问路径
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if

If Request("NewsID")<>"" and Request("Types")<>"" then
   NewsID = Replace(Replace(Replace(Cstr(Request("NewsID")),")",""),"(",""),"'","")
   Types = Cstr(NoSqlHack(Request("Types")))
Else
	Response.Write("<script>alert(""参数传递错误"");dialogArguments.location.reload();window.close();</script>")
	Response.End
End if
%>
<html>
<head>
<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script src="../js/Public.js" language="JavaScript"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加新闻到自由JS</title>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <form action="" name="ToPicJsForm" method="post" >
    <tr> 
      <td width="7%" height="5">&nbsp;</td>
      <td width="16%" height="5">&nbsp;</td>
      <td width="77%" height="5">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>JS名称</td>
      <td><select name="JSName" id="JSName" style="width:90%" onChange=ChooseJsName(this.options[this.selectedIndex].value)>
          <option value="" <%If Request("JSEName")="" then Response.Write("selected")%>> 
          </option>
      <%
	    Dim PicJsObj
		Set PicJsObj = Conn.Execute("Select EName,CName,Manner from FS_NS_FreeJS order by AddTime desc")
	    Do While Not PicJsObj.eof 
	  %>
          <option value="<%=PicJsObj("EName")&"***"&PicJsObj("Manner")%>" <%If Cstr(Request("JSEName")) = Cstr(PicJsObj("EName")) then Response.Write("selected")%>><%=PicJsObj("CName")%></option>
     <%
			PicJsObj.MoveNext
		Loop
	    PicJsObj.Close
		Set PicJsObj = Nothing
	  %>
        </select> <input name="Manner" type="hidden" id="Manner"> <input name="JSEName" type="hidden" id="JSEName"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>图片地址</td>
      <td>
	  <input name="PicPath" type="text" id="PicPath" size="28" value=""> 
       <input type="button" name="bnt_PicChooseButtonn"  value="选择图片" onClick="OpenWindowAndSetValue('../../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.ToPicJsForm.PicPath);">
	   </td>
    </tr>
    <tr> 
      <td height="5">&nbsp;</td>
      <td height="5">&nbsp;</td>
      <td height="5">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="3"><div align="center"> 
          <input type="button" name="Submit2" value=" 确 定 " onClick="ChoosePicPath();">
          <input name="action" type="hidden" id="action" value="trues">
          <input type="button" name="Submit3" value=" 取 消 " onClick="window.close();">
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<script>
function ChoosePicPath()
{
	var Value=parseInt(document.ToPicJsForm.Manner.value);
  document.ToPicJsForm.submit();
 }
 
function ChooseWordJsName(TempString)
{
   var TempArr=TempString.split("***");
   document.ToWordJsForm.Manner.value=TempArr[1];
   document.ToWordJsForm.JSEName.value=TempArr[0];
 }
 
function ChooseJsName(TempStr)
{
	var TempArray=TempStr.split("***");
	document.ToPicJsForm.Manner.value=TempArray[1];
	document.ToPicJsForm.JSEName.value=TempArray[0];
	var Value=parseInt(TempArray[1]);
	if (Value<6)
	{
		document.ToPicJsForm.PicPath.disabled=true;
		document.ToPicJsForm.bnt_PicChooseButtonn.disabled=true;
	}
	else
	{
		document.ToPicJsForm.PicPath.disabled=false;
		document.ToPicJsForm.bnt_PicChooseButtonn.disabled=false;
	}
}
</script>
<%
If Request.Form("action")="trues" then
  Dim JsFileObj,JsFileSql,TFFlagObj,NewsIDArray,Rt_i,RsNewsTFObj
  If Request.Form("JSEName")="" or isnull(Request.Form("JSEName")) then
	  Response.Write("<script>alert(""请选择自由JS"");</script>")
	  Response.End
  End If
 '=======================================
  NewsIDArray = Split(NewsID,"***")
  For Rt_i = 0 to UBound(NewsIDArray)
	If NewsIDArray(Rt_i) <> "" And IsNumeric(NewsIDArray(Rt_i)) Then
	  Set RsNewsTFObj = Conn.Execute("Select NewsID from FS_NS_FreeJsFile where JSName='"&NoSqlHack(Request.Form("JSEName"))&"' and NewsID=(Select NewsID from FS_NS_News where ID="&NewsIDArray(Rt_i)&")")
		  If RsNewsTFObj.eof Then
		  Set RsNewsObj = Conn.Execute("Select NewsTitle,NewsID,ClassID,addtime from FS_NS_News where isLock=0 and isRecyle=0 and ID="&NewsIDArray(Rt_i))
			 If Not RsNewsObj.eof Then
				  Set JsFileObj = Server.Createobject(G_FS_RS)
				  JsFileSql="select * from FS_NS_FreeJsFile where 1=0"
				  JsFileObj.open JsFileSql,Conn,3,3
				  JsFileObj.AddNew
				  JsFileObj("Title") = RsNewsObj("NewsTitle")
				  JsFileObj("JSName") = NoSqlHack(Request.Form("JSEName"))
				  JsFileObj("NewsID") = RsNewsObj("NewsID")
				  If Trim(Request.Form("PicPath"))<>"" And Not IsNull(Request.Form("PicPath"))  then
						JsFileObj("PicPath")=Request.Form("PicPath")
				  End if
				  JsFileObj("ClassID") = RsNewsObj("ClassID")
				  JsFileObj("NewsTime") = RsNewsObj("Addtime")
				  JsFileObj("ToJsTime") = Now
				  JsFileObj.Update
				  JsFileObj.Close
				  Set JsFileObj = Nothing
			 End if
			 RsNewsObj.Close
			 Set RsNewsObj = Nothing
		 End If
		 RsNewsTFObj.Close
		 Set RsNewsTFObj = Nothing
	   End If		 
	 Next
  
  '----------------生成JS文件-------------
  	Dim JSClassObj,ReturnValue,TempRs
	Set TempRs=Conn.Execute("Select NewsDir from FS_NS_SysParam")
	If TempRs.eof Then
		Response.Redirect("/error.asp?ErrCodes=<li>出现异常</li>")
		Response.End()
	End if
	Set JSClassObj = New Cls_Js
	JSClassObj.SysRootDir=TempRs("NewsDir")
  Select case Request.Form("Manner")
     case "1"   ReturnValue = JSClassObj.WCssA(Request.Form("JSEName"),True)
     case "2"   ReturnValue = JSClassObj.WCssB(Request.Form("JSEName"),True)
     case "3"   ReturnValue = JSClassObj.WCssC(Request.Form("JSEName"),True)
     case "4"   ReturnValue = JSClassObj.WCssD(Request.Form("JSEName"),True)
     case "5"   ReturnValue = JSClassObj.WCssE(Request.Form("JSEName"),True)
     case "6"   ReturnValue = JSClassObj.PCssA(Request.Form("JSEName"),True)
     case "7"   ReturnValue = JSClassObj.PCssB(Request.Form("JSEName"),True)
     case "8"   ReturnValue = JSClassObj.PCssC(Request.Form("JSEName"),True)
     case "9"   ReturnValue = JSClassObj.PCssD(Request.Form("JSEName"),True)
     case "10"  ReturnValue = JSClassObj.PCssE(Request.Form("JSEName"),True)
     case "11"  ReturnValue = JSClassObj.PCssF(Request.Form("JSEName"),True)
     case "12"  ReturnValue = JSClassObj.PCssG(Request.Form("JSEName"),True)
     case "13"  ReturnValue = JSClassObj.PCssH(Request.Form("JSEName"),True)
     case "14"  ReturnValue = JSClassObj.PCssI(Request.Form("JSEName"),True)
     case "15"  ReturnValue = JSClassObj.PCssJ(Request.Form("JSEName"),True)
     case "16"  ReturnValue = JSClassObj.PCssK(Request.Form("JSEName"),True)
     case "17"  ReturnValue = JSClassObj.PCssL(Request.Form("JSEName"),True)
   End Select
   Set JSClassObj = Nothing
  '----------------   Over   -------------
	if ReturnValue <> "" then
		Response.Write("<script>alert('" & ReturnValue & "');window.close();</script>")
	else
	  Response.Write("<script>window.close();</script>")
	end if
end If
Set TempRs=Nothing
Conn.close
Set Conn=nothing
%>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->