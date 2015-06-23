<!-- #include file = "Const.asp" -->
<%
Dim VRootStr
If G_VIRTUAL_ROOT_DIR = "" Then
	VRootStr = ""
Else
	VRootStr = "/" & Replace(Replace(G_VIRTUAL_ROOT_DIR,"//",""),"/","")
End If		
%>

<!--
//2007-01-09 by ken 用于人才搜索

var HiddenDivStr = '';
HiddenDivStr = '<div id="HiddenDiv" style="display:none;z-index:999;left:180px; top:100px; width:600px;position:absolute; border:solid 1px #3399CC;"><div style="width:600px; height:20px; background-color:#3399CC; text-align:right; line-height:20px; border-bottom:solid 1px #3399cc;cursor:move;" onMouseDown="new moveStart(event,\'HiddenDiv\')"><span style="margin-right:20px; font-size:12px; color:#f7f7f7; cursor:hand;" onClick="CloseDiv()">[关闭]</span></div><div style="background-color:#ffffff; width:600px; height:30px;text-align:center; line-height:30px; border-bottom:solid 1px #3399cc" id="FotherDiv">正在加载,请稍候...</div><div style="background-color:#ffffff; width:600px; height:30px;text-align:center; line-height:30px; display:none;" id="ChildDiv"></div></div>';
document.write(HiddenDivStr);

//搜索关键字鼠标点击取消文字
function ClearKey()
{
  if (document.Search_form.Sekey.value == "查询关键字")
  document.Search_form.Sekey.value = "";
}

//关闭隐藏层
function CloseDiv()
{
	document.getElementById('HiddenDiv').style.display='none';
}


//显示隐藏层
function DisDiv(Str)
{
	document.getElementById('HiddenDiv').style.display='';
	document.getElementById('ChildDiv').style.display='none';
	var PageUrl,ReturnStr,XMLobj,StrUrl
	PageUrl = 'APSearch.asp?act='+Str;
	StrUrl = '<% = VRootStr %>/FS_Inc/'+PageUrl;
	var myAjax = new Ajax.Request(
	StrUrl,
	{method:'get',
	parameters:'',
	onComplete:function (Receve){
			ReturnStr = Receve.responseText;
			if (ReturnStr !='')
			{
				document.getElementById('FotherDiv').innerHTML = ReturnStr;
			}
			else
			{
				document.getElementById('FotherDiv').innerHTML = '<font color=red>暂无信息</font>';
			}
		}
	}
	);
}

function DisChildDiv(Str,TypeStr)
{
	document.getElementById('ChildDiv').style.display='';
	var PageUrl,ReturnStr,XMLobj,StrUrl
	PageUrl = 'APSearch.asp?act=GetChlid&StrType='+TypeStr+'&ID='+Str;
	StrUrl = '<% = VRootStr %>/FS_Inc/'+PageUrl;
	var myAjax = new Ajax.Request(
	StrUrl,
	{method:'get',
	parameters:'',
	onComplete:function (Receve){
			ReturnStr = Receve.responseText;
			if (ReturnStr !='')
			{
				document.getElementById('ChildDiv').innerHTML = ReturnStr;
			}
			else
			{
				document.getElementById('ChildDiv').innerHTML = '<font color=red>暂无信息</font>';
			}
		}
	}
	);
}

function GetSelectValue(StrID,TypeStr,Str)
{
	if (TypeStr =='City')
	{
		document.getElementById('JobCity').value = Str;
		document.getElementById('JobCityID').value = StrID;
		document.getElementById('HiddenDiv').style.display='none';
	}
	else if (TypeStr =='Job') 
	{
		document.getElementById('JobType').value = Str;
		document.getElementById('JobTypeID').value = StrID;
		document.getElementById('HiddenDiv').style.display='none';
	}
	else if (TypeStr =='StrTime')
	{
		document.getElementById('JobTime').value = Str;
		document.getElementById('JobTimeID').value = StrID;
		document.getElementById('HiddenDiv').style.display='none';	
	}
}



function moveStart(event,_sId)
{
	var oObj = $(_sId);
	oObj.onmousemove = mousemove;
	oObj.onmouseup = mouseup;
	oObj.setCapture ? oObj.setCapture() : function(){};
	oEvent = window.event ? window.event : event;
	var dragData = {x : oEvent.clientX, y : oEvent.clientY};
	var backData = {x : parseInt(oObj.style.top), y : parseInt(oObj.style.left)};
	function mousemove()
	{
		var oEvent = window.event ? window.event : event;
		var iLeft = oEvent.clientX - dragData["x"] + parseInt(oObj.style.left);
		var iTop = oEvent.clientY - dragData["y"] + parseInt(oObj.style.top);
		oObj.style.left = iLeft;
		oObj.style.top = iTop;
		dragData = {x: oEvent.clientX, y: oEvent.clientY};

	}
	function mouseup()
	{
		var oEvent = window.event ? window.event : event;
		oObj.onmousemove = null;
		oObj.onmouseup = null;
		if(oEvent.clientX < 1 || oEvent.clientY < 1 || oEvent.clientX > document.body.clientWidth || oEvent.clientY > document.body.clientHeight){
			oObj.style.left = backData.y;
			oObj.style.top = backData.x;
		}
		oObj.releaseCapture ? oObj.releaseCapture() : function(){};
	}
}

-->





