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
	strShowErr = "<li>�Ҳ�����ϵͳ������Ϣ��</li>"
	Response.Redirect("../lib/error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Response.end
else
	MF_Domain = obj_mf_sys_obj("MF_Domain")
	MF_Site_Name = obj_mf_sys_obj("MF_Site_Name")
end if
obj_mf_sys_obj.close:set obj_mf_sys_obj = nothing
tmp_c_path =MF_Domain &"/"&G_VIRTUAL_ROOT_DIR


''�õ���ر��ֵ��
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
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Valueδ�ܵõ�������ݡ�����������"&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function

%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=GetUserSystemTitle%>-��ְ��Ƹ</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
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
          <td class="hback"><strong>λ�ã�</strong><a href="../">��վ��ҳ</a> &gt;&gt; 
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; ��ְλ</td>
        </tr>
	  <%	  
	  Dim IsCorporation
	  IsCorporation = Get_OtherTable_Value("select IsCorporation from FS_ME_Users where UserNumber='"&Session("FS_UserNumber")&"'")
	  if IsCorporation="1" then 
	 	response.Write("<tr class=""hback""><td>��ҵ��Աû����ְ����.����Ҫ������ְ,������ע��.</td></tr>")   
	  else
	  %>

		
        <tr class="hback">
          <td class="hback"><a href="#" onClick="getSearchPane();hightLightCurrent('search')" id="search">ְλ����</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="Person.asp" target="_blank">Ԥ������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="#" onClick="javascript:history.back()">����</a>
		  [�����<%
				Dim clickRs
				Set clickRs=Conn.execute("select click from FS_AP_Resume_BaseInfo where UserNumber='"&session("FS_UserNumber")&"'")
				If Not clickRs.eof Then
					response.write(clickRs("click"))
				End If
				clickRs.close():Set clickRs=nothing
			%>��]
		  </td>
        </tr>
        <tr class="hback">
          <td class="hback">
		  <a href="#" onClick="getResumeForm('resume_container','baseinfo','');hightLightCurrent('baseinfo')" id="baseinfo">������Ϣ</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','intention','');hightLightCurrent('intention')" id="intention">����/��������</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','position','');hightLightCurrent('position')" id="position">��ҵ/��λ</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','workcity','');hightLightCurrent('workcity')" id="workcity">�����ص�</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','workexp','');hightLightCurrent('workexp')" id="workexp">��������</a>&nbsp;&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','educateexp','');hightLightCurrent('educateexp')" id="educateexp">��������</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','trainexp','');hightLightCurrent('trainexp')" id="trainexp">��ѵ����</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','language','');hightLightCurrent('language')" id="language">��������</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','certificate','');hightLightCurrent('certificate')" id="certificate">֤��/����</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','projectexp','');hightLightCurrent('projectexp')" id="projectexp">��Ŀ����</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','other','');hightLightCurrent('other')" id="other">������Ϣ</a>&nbsp;
		  <a href="#" onClick="getResumeForm('resume_container','mail','');hightLightCurrent('mail')" id="mail">��ְ��</a></td>
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
//��ʾ״̬��������������������������������������������������������������������������������
new Ajax.Updater('resume_status',"lib/AP_BaseInfo_status.asp?and="+Math.random(),{method:'get', parameters:"action=edit"});
//��ȡ��Ӧ����������������������������������������������������������������������������������������
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
		default:alert('�����쳣������ϵ������Ա');return;break;
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
//��������������������������������������������������������������������������������������
function getSearchPane()
{
	var container="resume_container"
	$(container).innerHTML="<img src='../../sys_Images/progerssbar.gif'/>"
	$("resume_status").innerHTML=""
	var myAjax_pane = new Ajax.Updater(container,"job_search.asp?and="+Math.random(),{method:'get', parameters:"action=search"});
}
/*��ʾ������Ŀ�������*/
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
/*���ѡ�е�ֵ*/
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
//������ʾ��ǰ��������ӡ�������������������������������������������������������������������������������
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
//������и�����ʾ���ӡ�������������������������������������������������������������������������������
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
//�ύ��Post�����������������������������������������������������������������������������������
//@paramURL:Ŀ��URL
//@paramValues���ύ������
//@form:��ǰ��
//@action:����(�޸ģ����)
function ajaxPost(paramURL,paramValues,form,action,bid)
{
	var element=shadowDiv(form);//����Ч��
	var next;
	var part;
	var param;
	/*-�����һ��������ͼ------------------------*/
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
	/*-�����һ��������ͼ------------------------*/
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
			if(originalRequest.responseText.indexOf('����')>-1)			
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
			//��ʾ���е�selectԪ��
			for(var i=0;i<SelectArray.length;i++)
			{
				SelectArray[i].style.display='';
			}
			if($('loading')!=null&&$('innerloading')!=null)
			{
				document.body.removeChild($('loading'));
				document.body.removeChild($('innerloading'));
				Form.enable(form);//������е�����Ԫ��
			}
		}
	}

}
//��ʾ���������ļ��ز��������������������������������������������������������������������������������
//@form current Form element
function shadowDiv(form)
{
	Form.disable(form);
	//�������е�selectԪ��
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
//���ط��������ļ��ز��������������������������������������������������������������������������������
//@Obj:���صĲ�
//next:�¸�������ͼ
function CleanShadowDiv(Obj,next,part)
{
	Obj.innerHTML="<table border='0' class='table'><tr><td class='hback'><button onclick=\"getResumeForm('resume_container','"+next+"','','')\" style='width:140;height:80' >����ɹ���������һ��</button>&nbsp;&nbsp;<button onclick=\"getResumeForm('resume_container','"+part+"','','')\" style='width:140;height:80' >����ɹ���������һ��</button></td></table>";
}
//ɾ��������������������������������������������������������������������������������������
function Delete(part,id)
{
	if(confirm("ȷ��Ҫɾ��������¼��"))
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
				alert("ɾ�������ɹ���");
				getResumeForm("resume_container",part,id,action);
			}else
			{
				alert(result);
				alert("������������ϵ������Ա!");
			}
		}
	}
}
//����Ӱ����������������������������������������������������������������������������������
function errorNumber(number)
{
	switch(number)
	{
		case "1":return "�û�������Ϊ��";break;
		case "2":return "����Ӧ��Ϊ����";break;
		case "3":return "��������ӦΪ����";break;
		case "4":return "��ʼ���ڲ���Ϊ��";break;
		case "5":return "��˾������Ϊ��";break;
		case "6":return "ְλ����Ϊ��";break;
		case "7":return "�������ڲ���Ϊ��";break;
		case "8":return "ѧУ������Ϊ��";break;
		case "9":return "��ѵ��֯����Ϊ��";break;
		case "10":return "�������Ʋ���Ϊ��";break;
		case "11":return "�ȼ�(����)����Ϊ��";break;
		case "12":return "��֤ʱ�䲻��Ϊ��";break;
		case "13":return "֤�����Ʋ���Ϊ��";break;
		case "14":return "��Ŀ���Ʋ���Ϊ��";break;
		case "15":return "��Ŀ��������Ϊ��";break;
		case "16":return "������������Ϊ��";break;
		case "17":return "���ⲻ��Ϊ��";break;
		case "18":return "���ݲ���Ϊ��";break;
		case "19":return "��ҵ�����λ����Ϊ��";break;
		case "20":return "��ҵ�����λ�����ظ����";break;
		case "21":return "ʡ�ͳ��в���Ϊ��";break;
		case "22":return "�����ظ������ͬ����";break;
		default:return "";break;
	}
}
//������ҵ��ø�λ��������������������������������������������������������������������������������
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







