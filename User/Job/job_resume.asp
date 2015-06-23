<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<% Option Explicit %>
<%Session.CodePage=936%> 
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
dim obj_mf_sys_obj,MF_Domain,MF_Site_Name,tmp_c_path
set obj_mf_sys_obj = Conn.execute("select top 1 MF_Domain,MF_Site_Name from FS_MF_Config")
if obj_mf_sys_obj.eof then
	strShowErr = "<li>找不到主系统配置信息！</li>"
	Response.Redirect("../lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	MF_Domain = obj_mf_sys_obj("MF_Domain")
	MF_Site_Name = obj_mf_sys_obj("MF_Site_Name")
end if
obj_mf_sys_obj.close:set obj_mf_sys_obj = nothing
tmp_c_path =MF_Domain &"/"&G_VIRTUAL_ROOT_DIR


''得到相关表的值。
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if			
	if not This_Fun_Rs.eof then 
		if instr(This_Fun_Sql," in ")>0 then 
			do while not This_Fun_Rs.eof
				Get_OtherTable_Value  = Get_OtherTable_Value &""& This_Fun_Rs(0)	
				This_Fun_Rs.movenext		
			loop
		else
			Get_OtherTable_Value = This_Fun_Rs(0)
		end if
	else
		Get_OtherTable_Value = ""
	end if
	if Err.Number>0 then 
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Value未能得到相关数据。错误描述："&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function

%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=GetUserSystemTitle%>-求职招聘</title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统,风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
<script language="javascript" src="../../FS_Inc/PublicJS.js"></script>
<script language="javascript" src="../../FS_Inc/coolWindowsCalendar.js"></script>
</head>
<body id="mainContainer">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback">
	  <table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td class="hback"><strong>位置：</strong><a href="../">网站首页</a> &gt;&gt; 
            <a href="../main.asp">会员首页</a> &gt;&gt; 求职位</td>
        </tr>
	  <%	  
	  Dim IsCorporation
	  IsCorporation = Get_OtherTable_Value("select IsCorporation from FS_ME_Users where UserNumber='"&Session("FS_UserNumber")&"'")
	  if IsCorporation="1" then 
	 	response.Write("<tr class=""hback""><td>企业会员没有求职服务.若需要个人求职,请另行注册.</td></tr>")   
	  else
	  %>

		
        <tr class="hback">
          <td class="hback"><a href="#" onClick="getSearchPane();hightLightCurrent('search')" id="search">职位搜索</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="Person.asp" target="_blank">预览简历</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="#" onClick="javascript:history.back()">后退</a>
		  [被浏览<%
				Dim clickRs
				Set clickRs=Conn.execute("select click from FS_AP_Resume_BaseInfo where UserNumber='"&session("FS_UserNumber")&"'")
				If Not clickRs.eof Then
					response.write(clickRs("click"))
				End If
				clickRs.close():Set clickRs=nothing
			%>次]
		  </td>
        </tr>
        <tr class="hback">
          <td class="hback">
		  <a href="#" onClick="getResumeForm('resume_container','baseinfo','');hightLightCurrent('baseinfo')" id="baseinfo">基本信息</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','intention','');hightLightCurrent('intention')" id="intention">意向/自我评价</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','position','');hightLightCurrent('position')" id="position">行业/岗位</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','workcity','');hightLightCurrent('workcity')" id="workcity">工作地点</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','workexp','');hightLightCurrent('workexp')" id="workexp">工作经历</a>&nbsp;&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','educateexp','');hightLightCurrent('educateexp')" id="educateexp">教育经历</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','trainexp','');hightLightCurrent('trainexp')" id="trainexp">培训经历</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','language','');hightLightCurrent('language')" id="language">语言能力</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','certificate','');hightLightCurrent('certificate')" id="certificate">证书/荣誉</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','projectexp','');hightLightCurrent('projectexp')" id="projectexp">项目经验</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','other','');hightLightCurrent('other')" id="other">其它信息</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','mail','');hightLightCurrent('mail')" id="mail">求职信</a></td>
        </tr>
      </table>
	  <table border="0" width="98%" align="center">
          <tr>
            <td align="center"><div id="resume_status" align="center"></div></td>
          </tr>
        </table>
	    <table border="0" width="98%" align="center">
          <tr>
            <td align="center"><div id="resume_container" align="center"></div></td>
          </tr>
		<%end if%>
        </table>
		</td>
        </tr>
      </table>
	  </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
</table>
</body>
</html>
<%
Set Fs_User = Nothing
Set Conn=nothing
Set User_Conn=nothing
%>
<script language='javascript'>
var inputRight=false;
//显示状态■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
new Ajax.Updater('resume_status',"lib/AP_BaseInfo_status.asp?and="+Math.random(),{method:'get', parameters:"action=edit"});
//获取相应的输入界面■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
function getResumeForm(container,part,id,action)
{
	switch(part)
	{
		case 'baseinfo': hightLightCurrent('baseinfo');
							var url="lib/AP_Resume_BaseInfo.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_BaseInfo_status.asp?and="+Math.random();
							break;
		case 'intention': hightLightCurrent('intention');
							var url="lib/AP_Resume_Intention.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_Intention_status.asp?and="+Math.random();
							break;
		case 'position': hightLightCurrent('position');
							var url="lib/AP_Resume_position.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_position_status.asp?and="+Math.random();
							break;
		case 'workcity': hightLightCurrent('workcity');
							var url="lib/AP_Resume_workcity.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_workcity_status.asp?and="+Math.random();
							break;
		case 'workexp': hightLightCurrent('workexp');
							var url="lib/AP_Resume_WorkExp.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_WorkExp_status.asp?and="+Math.random();
							break;
		case 'educateexp': hightLightCurrent('educateexp');
							var url="lib/AP_Resume_EducateExp.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_EducateExp_status.asp?and="+Math.random();
							break;
		case 'trainexp': hightLightCurrent('trainexp');
							var url="lib/AP_Resume_TrainExp.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_TrainExp_status.asp?and="+Math.random();
							break;
		case 'language': hightLightCurrent('language');
							var url="lib/AP_Resume_Language.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_Language_status.asp?and="+Math.random();
							break;
		case 'certificate': hightLightCurrent('certificate');
							var url="lib/AP_Resume_Certificate.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_Certificate_status.asp?and="+Math.random();
							break;
		case 'projectexp': hightLightCurrent('projectexp');
							var url="lib/AP_Resume_ProjectExp.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_ProjectExp_status.asp?and="+Math.random();							
							break;
		case 'other': hightLightCurrent('other');
							var url="lib/AP_Resume_Other.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_Other_status.asp?and="+Math.random();
							break;
		case 'mail': hightLightCurrent('mail');
							var url="lib/AP_Resume_Mail.asp?id="+id+"&and="+Math.random();
							var url2="lib/AP_Mail_status.asp?and="+Math.random();
							break;
		case 'search': hightLightCurrent('search');
							var url="lib/AP_Resume_BaseInfo.asp?id="+id+"&and="+Math.random();
							break;
		default:alert('发生异常，请联系技术人员');return;break;
	}
	$(container).innerHTML="<img src='../../sys_Images/progerssbar.gif'/>"
	if($('loading')!=null&&$('innerloading')!=null)
	{
		document.body.removeChild($('loading'));
		document.body.removeChild($('innerloading'));
	}
	var myAjax_input = new Ajax.Updater(container,url,{method:'get', parameters:"action="+action});
	var myAjax_stauts = new Ajax.Updater('resume_status',url2,{method:'get', parameters:"action="+action});
}
//搜索面板■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
function getSearchPane()
{
	var container="resume_container"
	$(container).innerHTML="<img src='../../sys_Images/progerssbar.gif'/>"
	$("resume_status").innerHTML=""
	var myAjax_pane = new Ajax.Updater(container,"job_search.asp?and="+Math.random(),{method:'get', parameters:"action=search"});
}
/*显示搜索项目条件面板*/
function showPane(pane)
{
	var param="";
	var container=pane;
	$(container).innerHTML="<img src='../../sys_Images/progerssbar.gif'/>"
	switch(pane)
	{
		case "div_JobName":   $(pane).style.display="";param="condition=jobname";break;
		case "div_WorkCity":  $(pane).style.display="";param="condition=workcity";break
		case "div_PublicDate":$(pane).style.display="";param="condition=publicdate";break
	}
	var myAjax_search = new Ajax.Updater(container,"getSearchCondition.asp?and="+Math.random(),{method:'get', parameters:param});
}
/*获得选中的值*/
function chooseIt(input,value)
{
	$(input).value=value;
	if(arguments[2]!=""&&!isNaN(arguments[2]))
	{
		$("div_WorkCity_2").style.display="";
		$("div_WorkCity_2").innerHTML="<img src='../../sys_Images/progerssbar.gif'/>"
		var myAjax_search = new Ajax.Updater("div_WorkCity_2","getSearchCondition.asp?and="+Math.random(),{method:'get', parameters:"condition=workcity2&pid="+arguments[2]});
	}
	if(!isNaN(arguments[3]))
	{
		$('hd_PublicDate').value=arguments[3];
	}
}
//高亮显示当前步骤的连接■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
function hightLightCurrent(value)
{
	switch(value)
	{
		case 'baseinfo':cleanAllFightLightCurrent(); $('baseinfo').style.color='red';break;
		case 'intention':cleanAllFightLightCurrent();$('intention').style.color='red';break;
		case 'position':cleanAllFightLightCurrent();$('position').style.color='red';break;
		case 'workcity':cleanAllFightLightCurrent();$('workcity').style.color='red';break;
		case 'workexp': cleanAllFightLightCurrent();$('workexp').style.color='red';break;
		case 'educateexp':cleanAllFightLightCurrent(); $('educateexp').style.color='red';break;
		case 'trainexp': cleanAllFightLightCurrent();$('trainexp').style.color='red';break;
		case 'language':cleanAllFightLightCurrent();$('language').style.color='red';break;
		case 'certificate': cleanAllFightLightCurrent();$('certificate').style.color='red';break;
		case 'projectexp': cleanAllFightLightCurrent();$('projectexp').style.color='red';break;
		case 'other':cleanAllFightLightCurrent(); $('other').style.color='red';break;
		case 'mail': cleanAllFightLightCurrent();$('mail').style.color='red';break;
		case 'search':cleanAllFightLightCurrent(); $('search').style.color='red';break;
	}
}
//清除所有高亮显示连接■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
function cleanAllFightLightCurrent()
{
	$('baseinfo').style.color="";
	$('intention').style.color="";
	$('position').style.color="";
	$('workcity').style.color="";
	$('workexp').style.color="";
	$('educateexp').style.color="";
	$('trainexp').style.color="";
	$('language').style.color="";
	$('certificate').style.color="";
	$('projectexp').style.color="";
	$('other').style.color="";
	$('mail').style.color="";
	$('search').style.color="";
}
//提交表单Post请求■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
//@paramURL:目标URL
//@paramValues：提交的数据
//@form:当前表单
//@action:动作(修改，添加)
function ajaxPost(paramURL,paramValues,form,action,bid)
{
	var element=shadowDiv(form);//加载效果
	var next;
	var part;
	var param;
	/*-获得下一个输入视图------------------------*/
	switch(form)
	{
		case "BaseInfoForm" :next="intention";break;
		case "IntentionForm":next="position";break;
		case "PositionForm":next="workcity";break;
		case "WorkCityForm":next="workexp";break;
		case "WorkExpForm":next="educateexp";break;
		case "EducateExpForm":next="trainexp";break;
		case "TrainExpForm":next="language";break;
		case "LanuageForm":next="certificate";break;
		case "CertificateForm":next="projectexp";break;
		case "ProjectExpForm":next="other";break;
		case "OtherForm":next="mail";break;
		case "MailForm":next="baseinfo";break;
	}
	/*-------------------------*/
	/*-获得下一个输入视图------------------------*/
	switch(form)
	{
		case "BaseInfoForm":part="baseinfo";break
		case "IntentionForm":part="intention";break
		case "PositionForm":part="position";break
		case "WorkCityForm":part="workcity";break
		case "WorkExpForm":part="workexp";break
		case "EducateExpForm":part="educateexp";break
		case "TrainExpForm" :part="trainexp";break;
		case "LanuageForm" :part="language";break;
		case "CertificateForm" :part="certificate";break;
		case "ProjectExpForm":part="projectexp";break;
		case "OtherForm" :part="other";break;
		case "MailForm" :part="mail";break;
	}
	/*-------------------------*/
	param=paramValues+"&action="+action+"&part="+part+"&id="+bid;
	var ajaxRequest=new Ajax.Request(paramURL,{method:'post',parameters:"ran="+Math.random(),onComplete:response,postBody:param});
	function response(originalRequest)
	{
		if(originalRequest.responseText=="ok")
		{
			CleanShadowDiv(element,next,part);
		}
		else
		{  
			if(originalRequest.responseText.indexOf('错误')>-1)			
			{document.body.removeChild($('loading'));document.body.removeChild($('innerloading'));alert(originalRequest.responseText);return false;}
			var errorarray=originalRequest.responseText.split('*')
			var msg="";
			for(var i=0;i<errorarray.length;i++)
			{
				if(errorarray[i]!="")
					msg+=(i+1)+"."+errorNumber(errorarray[i])+"\n\n";
			}
			alert(msg);
			var SelectArray=document.getElementsByTagName("select");
			//显示所有的select元素
			for(var i=0;i<SelectArray.length;i++)
			{
				SelectArray[i].style.display='';
			}
			if($('loading')!=null&&$('innerloading')!=null)
			{
				document.body.removeChild($('loading'));
				document.body.removeChild($('innerloading'));
				Form.enable(form);//激活表单中的所有元素
			}
		}
	}

}
//显示发送请求后的加载层■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
//@form current Form element
function shadowDiv(form)
{
	Form.disable(form);
	//隐藏所有的select元素
	var SelectArray=document.getElementsByTagName("select");
	for(var i=0;i<SelectArray.length;i++)
	{
		SelectArray[i].style.display='none';
	}
	var oElement = document.createElement("<DIV id='loading' style='z-index:10;position:absolute;left:0;top:0;FILTER:alpha(opacity=50);background-color:#efefef;'align='center'></DIV>")
	oElement.style.width=screen.availWidth;
	oElement.style.height=screen.height;
	document.body.appendChild(oElement);
	var innerDIV = document.createElement("<DIV id='innerloading'  style='position:absolute;z-index:100' align='center'></DIV>")
	innerDIV.style.left=(screen.availWidth-280)/2;
	innerDIV.style.top=(screen.height-160)/2;
	innerDIV.style.width=100;
	innerDIV.style.height=100;
	document.body.appendChild(innerDIV);
	var imageElement=document.createElement("<img src='../../sys_Images/progerssbar.gif'/>");
	innerDIV.appendChild(imageElement);
	return innerDIV;
}
//隐藏发送请求后的加载层■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
//@Obj:加载的层
//next:下个输入视图
function CleanShadowDiv(Obj,next,part)
{
	Obj.innerHTML="<table border='0' class='table'><tr><td class='hback'><button onclick=\"getResumeForm('resume_container','"+next+"','','')\" style='width:140;height:80' >保存成功，进入下一步</button>&nbsp;&nbsp;<button onclick=\"getResumeForm('resume_container','"+part+"','','')\" style='width:140;height:80' >保存成功，继续这一步</button></td></table>";
}
//删除操作■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
function Delete(part,id)
{
	if(confirm("确定要删除该条记录？"))
	{
		var url="AP_Resume_Action.asp";
		var pars="action=del&delpart="+part+"&id="+id;
		var myAjax = new Ajax.Request(url,{method: 'get', parameters: pars, onComplete: showResponse});
		var action="";
		function showResponse(originalRequest)
		{
			var result= originalRequest.responseText;
			if(result=="ok")
			{
				alert("删除操作成功！");
				getResumeForm("resume_container",part,id,action);
			}else
			{
				alert(result);
				alert("发生错误，请联系技术人员!");
			}
		}
	}
}
//错误影射表■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
function errorNumber(number)
{
	switch(number)
	{
		case "1":return "用户名不能为空";break;
		case "2":return "年龄应该为数字";break;
		case "3":return "工做年限应为数字";break;
		case "4":return "开始日期不能为空";break;
		case "5":return "公司名不能为空";break;
		case "6":return "职位不能为空";break;
		case "7":return "结束日期不能为空";break;
		case "8":return "学校名不能为空";break;
		case "9":return "培训组织不能为空";break;
		case "10":return "语言名称不能为空";break;
		case "11":return "等级(分数)不能为空";break;
		case "12":return "获证时间不能为空";break;
		case "13":return "证书名称不能为空";break;
		case "14":return "项目名称不能为空";break;
		case "15":return "项目描述不能为空";break;
		case "16":return "责任描述不能为空";break;
		case "17":return "标题不能为空";break;
		case "18":return "内容不能为空";break;
		case "19":return "行业或则岗位不能为空";break;
		case "20":return "行业或则岗位不能重复添加";break;
		case "21":return "省和城市不能为空";break;
		case "22":return "不能重复添加相同数据";break;
		default:return "";break;
	}
}
//根据行业获得岗位■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
function getJob(container,tid)
{
	if(tid=="")
	{
		return false;
	}
	new Ajax.Updater(container,"lib/getJob.asp",{method:'get',parameters:"tid="+tid+"&rnd="+Math.random()})
}
function getCity(container,pid)
{
	if(pid=="")
	{
		return false;
	}
	new Ajax.Updater(container,"lib/getCity.asp",{method:'get',parameters:"pid="+pid+"&rnd="+Math.random()})
}
function setValue(from,to)
{
	var text=from.options[from.selectedIndex].text
	var value=from.options[from.selectedIndex].value
	if(value!="")
	{
		$(to).value=text
	}else
	{
		$(to).value=""
	}
}
</script>







