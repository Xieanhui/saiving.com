<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%'Copyright (c) 2006 Foosun Inc.
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,DS_Rs,DS_Sql
Dim AutoDelete,Months
MF_Default_Conn 
MF_Session_TF

dim sRootDir,str_CurrPath
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
str_CurrPath = sRootDir &"/"&G_TEMPLETS_DIR

Sub Save()
	Dim sysid,Str_Tmp,Arr_Tmp,IndexPage,IsLockTF
	sysid = NoSqlHack(request.Form("sysid"))
	Str_Tmp = "IPType,IPList,OverDueMode,DownDir,IndexTemplet,LinkType,IsDomain,FileNameRule,FileDirRule,ClassSaveType,IndexPage,NewsCheck,FileExtName"
	Arr_Tmp = split(Str_Tmp,",")
	IndexPage = NoSqlHack(request.Form("IndexPage_Name"))&"."&NoSqlHack(request.Form("IndexPage_ExtName"))
	if trim(Request.Form("FileNameRule_Element_Separator"))<>"" then
		if not chkinputchar(NoSqlHack(Request.Form("FileNameRule_Element_Separator"))) then
			Response.Redirect("../error.asp?ErrCodes=<li>分割符号只允许为：""0-9""，""A-Z""，""-"",""_"","",""."",""@"",""#""</li>")
			Response.End()
		end if
	End if
	fileNameRule=NoSqlHack(Request.Form("FileNameRule_Element_Prefix"))&"$"&NoSqlHack(replace(Request.Form("FileNameRule_Element"),",",""))&"$"&NoSqlHack(Request.Form("FileNameRule_Rnd"))&"$"&NoSqlHack(Request.Form("FileNameRule_UseWord"))&"$"&NoSqlHack(Request.Form("FileNameRule_Element_Separator"))&"$"&NoSqlHack(Request.Form("FileNameRule_UseDownID"))&"$"&NoSqlHack(Request.Form("FileNameRule_DownID"))
	IsLockTF = request.Form("Lock")
	If IsLockTF = "" then
		IsLockTF = 0
	Else
		IsLockTF = CintStr(IsLockTF)	
	End If	
	DS_Sql = "select top 1 "&Str_Tmp&",IndexPage,FileNameRule,Lock  from FS_DS_SysPara"
	'response.Write(DS_Sql)
	Set DS_Rs = CreateObject(G_FS_RS)
	DS_Rs.Open DS_Sql,Conn,3,3
	if DS_Rs.eof then DS_Rs.AddNew
	for each Str_Tmp in Arr_Tmp
		'response.Write(Str_Tmp&":"&NoSqlHack(request.Form(Str_Tmp))&"<br>")
		DS_Rs("Lock") = IsLockTF
		DS_Rs("IndexPage") = IndexPage
		DS_Rs("FileNameRule")  = fileNameRule
		DS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
	next
	'response.End()
	DS_Rs.update
	DS_Rs.close
	DSConfig_Cookies
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Down/FS_DS_SysPara.asp" )&"&ErrCodes=<li>恭喜，修改成功。</li>")
End Sub
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" src="../../FS_Inc/checkJs.js"></script>
<head><body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr  class="hback">
    <td colspan="10" align="left" class="xingmu" >下载系统参数设置</td>
  </tr>
  <tr  class="hback">
    <td colspan="10" height="25"><a href="FS_DS_SysPara.asp">管理首页</a> </td>
  </tr>
</table>
<%

'******************************************************************
if request.QueryString("Act")="Save" then 
	Call Save
else
	Call Add_Edit_Search		
end if
'******************************************************************
Dim Bol_IsEdit
Dim Lock,IPType,IPList,OverDueMode,IsDomain,FileNameRule,FileDirRule,ClassSaveType,FileExtName
Dim DownDir,IndexTemplet,LinkType,IndexPage,NewsCheck ,FileNameRuleArray,IndexPageArray
Sub Add_Edit_Search()
Bol_IsEdit = false
DS_Sql = "select top 1 Lock,IPType,IPList,OverDueMode,DownDir,IndexTemplet,LinkType,IsDomain,FileNameRule,FileDirRule,ClassSaveType,FileExtName,IndexPage,NewsCheck from FS_DS_SysPara"
Set DS_Rs	= CreateObject(G_FS_RS)
DS_Rs.Open DS_Sql,Conn,1,1
if not DS_Rs.eof then 
	Bol_IsEdit = True
	Lock = DS_Rs("Lock")
	IPType = DS_Rs("IPType")
	IPList = DS_Rs("IPList")
	OverDueMode = DS_Rs("OverDueMode")
	IsDomain = DS_Rs("IsDomain")
	FileNameRule = DS_Rs("FileNameRule")
	FileDirRule = DS_Rs("FileDirRule")
	ClassSaveType = DS_Rs("ClassSaveType")
	FileExtName = DS_Rs("FileExtName")
	IndexPage = DS_Rs("IndexPage")
	NewsCheck = DS_Rs("NewsCheck")
	DownDir = DS_Rs("DownDir")
	IndexTemplet = DS_Rs("IndexTemplet")
	LinkType = DS_Rs("LinkType")
else
	Lock = 1
	IPType = 1
	IPList = ""
	OverDueMode = 1
	IsDomain = ""
	FileNameRule = "FS$YMDHIS$2$1$-$1"
	FileDirRule = 0
	ClassSaveType = 0
	FileExtName = 0
	IndexPage = "index,html"
	NewsCheck = 1	
	DownDir = "Down"
	IndexTemplet = ""
	LinkType = 0
end if
FileNameRuleArray=split(FileNameRule,"$")
IndexPage = replace(IndexPage,",",".")
IndexPageArray=split(IndexPage,".")
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="Form" id="Form" method="post" action="?Act=Save">
    <tr  class="hback">
      <td colspan="3" align="left" class="xingmu" >系统参数设置信息</td>
    </tr>
    <tr  class="hback">
      <td width="110" align="right">是否加防盗链</td>
      <td><input <% if Lock=1 then Response.Write("Checked") %>  name="Lock" type="checkbox" id="Lock" value="1">
        <input name="IPList" type="hidden" id="IPList">
      </td>
    </tr>
    <tr  class="hback">
      <td align="right">类型</td>
      <td><select name="IPType">
          <%=PrintOption(IPType,"1:阻止列表,2:允许列表")%>
        </select>
      </td>
    </tr>
    <tr  class="hback">
      <td align="right">IP地址列表</td>
      <td><select name="IPSelectList" size="10" multiple id="IPSelectList" style="width:60%;">
          <%
		  Dim TempArray,i
		  if Not IsNull(IPList) and IPList<>"" then
			  TempArray = Split(IPList,"$")
			  for i=LBound(TempArray) to UBound(TempArray)
			  %>
          <option value="<% = TempArray(i) %>">
          <% = TempArray(i) %>
          </option>
          <%
			  Next
		  end if
		  %>
        </select></td>
    </tr>
    <tr  class="hback">
      <td align="right">起始IP</td>
      <td><input name="BeginIP" type="text" id="BeginIP">
        ---
        <input name="EndIP" type="text" id="EndIP">
        <input type="button" onClick="AddIPList();" name="Submit3" value=" 添 加 ">
        <input type="button" onClick="DelIPList();" name="Submit4" value=" 删 除 ">
      </td>
    </tr>
    <tr  class="hback">
      <td align="right">过期下载处理方式</td>
      <td><select name="OverDueMode">
          <%=PrintOption(OverDueMode,"1:删除,2:提示已过期")%>
        </select>
      </td>
    </tr>
    </tr>
    
    <!--新加的--->
    <tr class="hback">
      <td align="right"> 系统前台目录：</td>
      <td><input name="DownDir" type="text" id="DownDir" value="<%=DownDir%>" size="50" maxlength="20">
        <font color="red">*</font><span id="span_DownDir_Alert"></span></td>
    </tr>
    <tr class="hback">
      <td align="right">启用二级域名：</td>
      <td><input name="IsDomain" type="text" id="IsDomain" value="<%=isDomain%>" size="50">
        <br>
        格式：Down.foosun.cn，不带&quot;http://&quot;或者虚拟目录，后面不带&quot;/&quot;.如果不开启二级域名，空保持为空</td>
    </tr>
    <tr class="hback">
      <td align="right">首页模板地址：</td>
      <td><input name="IndexTemplet" type="text" id="IndexTemplet" size="50" value="<%=indexTemplet%>">
        <input name="bnt_NewsTemplet" type="button" id="bnt_NewsTemplet" value="选择模板"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=str_CurrPath %>',400,300,window,document.Form.IndexTemplet);document.Form.IndexTemplet.focus();">
        <span class="style2">*</span><span id="span_IndexTemplet_Alert"></span></td>
    </tr>
    <tr class="hback">
      <td align="right">连接路径：</td>
      <td><input type="radio" name="LinkType" value="1" <%if linkType=1 then Response.Write("checked")%>>
        绝对路径
        <input name="LinkType" type="radio" value="0" <%if linkType=0 then Response.Write("checked")%>>
        相对路径 </td>
    </tr>
    <tr class="hback">
      <td align="right">文件名前缀：</td>
      <td><input name="FileNameRule_Element_Prefix" type="text" id="FileNameRule_Element_Prefix" value="<%=FileNameRuleArray(0)%>" size="50" maxlength="10"></td>
    </tr>
    <tr class="hback">
      <td align="right">文件名参数：</td>
      <td><input name="FileNameRule_Element" type="checkbox" id="FileNameRule_Element" value="Y" <%if instr(FileNameRuleArray(1),"Y")>0 then Response.Write("checked")%>>
        年
        <input name="FileNameRule_Element" type="checkbox" id="FileNameRule_Element" value="M" <%if instr(FileNameRuleArray(1),"M")>0 then Response.Write("checked")%>>
        月
        <input name="FileNameRule_Element" type="checkbox" id="FileNameRule_Element" value="D" <%if instr(FileNameRuleArray(1),"D")>0 then Response.Write("checked")%>>
        日
        <input name="FileNameRule_Element" type="checkbox" id="FileNameRule_Element" value="H" <%if instr(FileNameRuleArray(1),"H")>0 then Response.Write("checked")%>>
        时
        <input name="FileNameRule_Element" type="checkbox" id="FileNameRule_Element" value="I" <%if instr(FileNameRuleArray(1),"I")>0 then Response.Write("checked")%>>
        分
        <input name="FileNameRule_Element" type="checkbox" id="FileNameRule_Element" value="S" <%if instr(FileNameRuleArray(1),"S")>0 then Response.Write("checked")%>>
        秒 <br>
        <input type="radio" name="FileNameRule_Rnd" id="FileNameRule_Rnd" value="2" <%if FileNameRuleArray(2)="2" then Response.Write("checked")%>>
        2位随机数
        <input type="radio" name="FileNameRule_Rnd" id="FileNameRule_Rnd" value="3" <%if FileNameRuleArray(2)="3" then Response.Write("checked")%>>
        3位随机数
        <input type="radio" name="FileNameRule_Rnd" id="FileNameRule_Rnd" value="4" <%if FileNameRuleArray(2)="4" then Response.Write("checked")%>>
        4位随机数
        <input type="radio" name="FileNameRule_Rnd" id="FileNameRule_Rnd" value="5" <%if FileNameRuleArray(2)="5" then Response.Write("checked")%>>
        5位随机数
        <input name="FileNameRule_UseWord" type="checkbox" id="FileNameRule_UseWord" value="1" <%if ubound(FileNameRuleArray)>=3 then if FileNameRuleArray(3)="1" then Response.Write("checked")%>>
        是否组合字母 </td>
    </tr>
    <tr class="hback">
      <td align="right">分割符号：</td>
      <td><input name="FileNameRule_Element_Separator" type="text" id="FileNameRule_Element_Separator" size="50" value="<%=FileNameRuleArray(4)%>"></td>
    </tr>
    <tr class="hback">
      <td align="right">是否使用自动ID： </td>
      <td><input type="radio" name="FileNameRule_UseDownID" value="1" <%if ubound(FileNameRuleArray)>=5 then if FileNameRuleArray(5)="1" then Response.Write("checked")%> onClick="clearAll('FileNameRule_Rnd','FileNameRule_UseWord')">
        是
        <input type="radio" name="FileNameRule_UseDownID" value="0" <%if Ubound(FileNameRuleArray)>=5 then if FileNameRuleArray(5)="0" then Response.Write("checked")%> onClick="checkIt('FileNameRule_Rnd','FileNameRule_UseWord')">
        否 </td>
    </tr>
    <tr class="hback" style="display:">
      <td align="right">是否使用DownID：</td>
      <td><input type="radio" name="FileNameRule_DownID" value="1" <%if ubound(FileNameRuleArray)>=6 then if FileNameRuleArray(6)="1" then Response.Write("checked")%> onClick="clearAll('FileNameRule_Rnd','FileNameRule_UseWord')">
        是
        <input type="radio" name="FileNameRule_DownID" value="0" <%if Ubound(FileNameRuleArray)>=6 then if FileNameRuleArray(6)="0" then Response.Write("checked")%> onClick="checkIt('FileNameRule_Rnd','FileNameRule_UseWord')">
        否 </td>
    </tr>
    <tr class="hback">
      <td align="right">目录生成规则：</td>
      <td><input name="FileDirRule" type="radio" value="0" onClick="show_FileDirRule_Detail(this.value);" <%if fileDirRule=0 then Response.Write("checked")%>>
        规则1
        <input type="radio" name="FileDirRule" value="1" onClick="show_FileDirRule_Detail(this.value)" <%if fileDirRule=1 then Response.Write("checked")%>>
        规则2
        <input type="radio" name="FileDirRule" value="2" onClick="show_FileDirRule_Detail(this.value)" <%if fileDirRule=2 then Response.Write("checked")%>>
        规则3
        <input type="radio" name="FileDirRule" value="3" onClick="show_FileDirRule_Detail(this.value)" <%if fileDirRule=3 then Response.Write("checked")%>>
        规则4
        <input type="radio" name="FileDirRule" value="4" onClick="show_FileDirRule_Detail(this.value)" <%if fileDirRule=4 then Response.Write("checked")%>>
        规则5
        <input type="radio" name="FileDirRule" value="5" onClick="show_FileDirRule_Detail(this.value)" <%if fileDirRule=5 then Response.Write("checked")%>>
        规则6
        <input type="radio" name="FileDirRule" value="6" onClick="show_FileDirRule_Detail(this.value)" <%if fileDirRule=6 then Response.Write("checked")%>>
        规则7 &nbsp;&nbsp;<span id="span_FileDirRule" style="color:blue"></span> </td>
    </tr>
    <tr class="hback">
      <td align="right">首页生成规则：</td>
      <td><input name="ClassSaveType" type="radio" value="0" onClick="show_ClassSaveType_Detail(this.value)" <%if classSaveType=0 then Response.Write("checked")%>>
        规则1
        <input type="radio" name="ClassSaveType" value="1" onClick="show_ClassSaveType_Detail(this.value)" <%if classSaveType=1 then Response.Write("checked")%>>
        规则2
        <input type="radio" name="ClassSaveType" value="2" onClick="show_ClassSaveType_Detail(this.value)" <%if classSaveType=2 then Response.Write("checked")%>>
        规则3 &nbsp;&nbsp;<span id="span_ClassSaveType" style="color:blue"></span> </td>
    </tr>
    <tr class="hback">
      <td align="right">文件扩展名：</td>
      <td><input type="radio" name="FileExtName" value="0" <%if fileExtName=0 then Response.Write("checked")%>>
        Html
        <input name="FileExtName" type="radio" value="1" <%if fileExtName=1 then Response.Write("checked")%>>
        HTM
        <input type="radio" name="FileExtName" value="2" <%if fileExtName=2 then Response.Write("checked")%>>
        Shtml
        <input type="radio" name="FileExtName" value="3" <%if fileExtName=3 then Response.Write("checked")%>>
        shtm
        <input type="radio" name="FileExtName" value="4" <%if fileExtName=4 then Response.Write("checked")%>>
        asp</td>
    </tr>
    <tr class="hback">
      <td align="right">首页文件名：</td>
      <td><input name="IndexPage_Name" type="text" id="IndexPage_Name" size="50" maxlength="10" value="<%=IndexPageArray(0)%>"
	   onFocus="Do.these('IndexPage_Name',function(){return isEmpty('IndexPage_Name','span_IndexPage_Name_Alert')});" onKeyUp="Do.these('IndexPage_Name',function(){return isEmpty('IndexPage_Name','span_IndexPage_Name_Alert')});">
        <span id="span_IndexPage_Name_Alert"></span>* </td>
    </tr>
    <tr class="hback">
      <td align="right">首页扩展名：</td>
      <td><select name="IndexPage_ExtName" id="IndexPage_ExtName">
          <option value="html" <%if IndexPageArray(1)="html" then Response.Write("selected")%>>html</option>
          <option value="htm" <%if IndexPageArray(1)="htm" then Response.Write("selected")%>>htm</option>
          <option value="shtml" <%if IndexPageArray(1)="shtml" then Response.Write("selected")%>>shtml</option>
          <option value="shtm" <%if IndexPageArray(1)="shtm" then Response.Write("selected")%>>shtm</option>
          <option value="asp" <%if IndexPageArray(1)="asp" then Response.Write("selected")%>>asp</option>
        </select>
      </td>
    </tr>
    <tr class="hback">
      <td align="right">发布是否需审核：</td>
      <td><input name="NewsCheck" type="radio" value="1" <%if NewsCheck=1 then Response.Write("checked")%>>
        是
        <input type="radio" name="NewsCheck" value="0" <%if NewsCheck=0 then Response.Write("checked")%>>
        否 </td>
    </tr>
    <tr  class="hback">
      <td colspan="4"><table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td align="center"><input type="submit" value=" 保存 "  onClick="SetIPList();" />
              &nbsp;
              <input type="reset" value=" 重置 " />
            </td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>
<%
End Sub
set DS_Rs = Nothing
Conn.close
%>
<script language="JavaScript">
function AddIPList()
{
	var BeginIPStr=document.Form.BeginIP.value,EndIPStr=document.Form.EndIP.value;
	if (CheckIP(BeginIPStr))
	{
		if (CheckIP(EndIPStr))
		{
			if (CheckBeginAndEndIP(BeginIPStr,EndIPStr))
			{
				var TempStr=BeginIPStr+'-'+EndIPStr;
				AddList(document.Form.IPSelectList,TempStr,TempStr);
				document.Form.BeginIP.value='';
				document.Form.EndIP.value='';
			}
		}
		else
		{
			alert('结束IP地址不对');
			document.Form.EndIP.focus();
			document.Form.EndIP.select();
		}
	}
	else
	{
		alert('开始IP地址不对');
		document.Form.BeginIP.focus();
		document.Form.BeginIP.select();
	}
}
function DelIPList()
{
	DelList(document.Form.IPSelectList);
}
function SetIPList()
{
	var flag1=isEmpty("IndexPage_Name","span_IndexPage_Name_Alert");
	var flag2=isEmpty("DownDir","span_DownDir_Alert");
	var flag4=isEmpty("IndexTemplet","span_IndexTemplet_Alert");
	if(flag1&&flag2&&flag4)
	{
		var TempStr='',Obj=document.Form.IPSelectList;
		for(var i=0;i<Obj.length;i++)
		{
			if (TempStr=='') TempStr=Obj.options(i).value;
			else TempStr=TempStr+'$'+Obj.options(i).value;
		}
		document.Form.IPList.value=TempStr;
		document.Form.submit();
	}
}
function CheckBeginAndEndIP(BeginIPStr,EndIPStr)
{
	return true;
}
function CheckIP(IPAddress)
{
	var re = IPAddress.split(".");
	var check = function(v){try{return (v<=255 && v>=0)}catch(x){return false}};
	var ip = (re.length==4)?(check(re[0]) && check(re[1]) && check(re[2]) && check(re[3])):false;
	return ip;
}
function AddList(SelectObj,Lable,LableContent)
{
	var i=0,AddOption;
	if (!SearchOptionExists(SelectObj,Lable))
	{
		AddOption = document.createElement("OPTION");
		AddOption.text=Lable;
		AddOption.value=LableContent;
		SelectObj.add(AddOption);
		//SelectObj.options(SelectObj.length-1).selected=true;
	}
}
function SearchOptionExists(Obj,SearchText)
{
	var i;
	for(i=0;i<Obj.length;i++)
	{
		if (Obj.options(i).text==SearchText)
		{
			Obj.options(i).selected=true;
			return true;
		}
	}
	return false;
}
function DelList(SelectObj)
{
	var OptionLength=SelectObj.length;
	for(var i=0;i<OptionLength;i++)
	{
		if (SelectObj.options(SelectObj.length-1).selected==true) SelectObj.options.remove(SelectObj.length-1);
		//OptionLength=SelectObj.length;
	}
}
show_FileDirRule_Detail(<%=FileDirRule%>);
//显示相应目录生成规则的格式
function show_FileDirRule_Detail(param)
{
	if(isNaN(param))
	{
		return;
	}
	switch(parseInt(param))
	{
		case 0:document.getElementById("span_FileDirRule").innerHTML="格式：[ 行业英文/2006-6-9 ]";break
		case 1:document.getElementById("span_FileDirRule").innerHTML="格式：[ 行业英文/2006/6/9/ ]";break
		case 2:document.getElementById("span_FileDirRule").innerHTML="格式：[ 行业英文/2006/6-9/ ]";break
		case 3:document.getElementById("span_FileDirRule").innerHTML="格式：[ 行业英文/2006-6/9/ ]";break
		case 4:document.getElementById("span_FileDirRule").innerHTML="格式：[ 行业英文/文件名 ]";break
		case 5:document.getElementById("span_FileDirRule").innerHTML="格式：[ 行业英文/2006/6/ ]";break
		case 6:document.getElementById("span_FileDirRule").innerHTML="格式：[ 行业英文/200669/ ]";break
	}
}
show_ClassSaveType_Detail(<%=ClassSaveType%>);
function show_ClassSaveType_Detail(param)
{
	if(isNaN(param))
		return;
	switch(parseInt(param))
	{
		case 0:document.getElementById("span_ClassSaveType").innerHTML="格式：[ 栏目英文/index.html ]";break
		case 1:document.getElementById("span_ClassSaveType").innerHTML="格式：[ 栏目英文/栏目英文.html ]";break
		case 2:document.getElementById("span_ClassSaveType").innerHTML="格式：[ 为栏目英文.html ]";break
	}
}
function clearAll(radio,check)
{
	var RadioArray=document.all(radio);
	for(var i=0;i<RadioArray.length;i++)
	{
		RadioArray[i].checked=false;
	}
	document.all(check).checked=false;
}
checkIt('FileNameRule_Rnd','FileNameRule_UseWord')
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
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->






