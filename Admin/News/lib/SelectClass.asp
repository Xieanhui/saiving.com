<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.Charset="GB2312"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>选择栏目</title>
	<link href="../../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css"
		rel="stylesheet" type="text/css" />
	<style type="text/css">
		.LableSelectItem
		{
			background-color: highlight;
			cursor: pointer;
			color: white;
			text-decoration: none;
		}
		.LableItem
		{
			cursor: pointer;
		}
		.SubItem
		{
			margin-left: 15px;
		}
		.RootItem
		{
			margin-left: 15px;
		}
		body
		{
			margin-left: 0px;
			margin-top: 0px;
			margin-right: 0px;
			margin-bottom: 0px;
			line-height: 120%;
		}
	</style>

	<script type="text/javascript" src="../../../FS_Inc/Prototype.js"></script>

	<script type="text/javascript">
	<!--
		window.SelectSingleClass = [[], [], []];
		window.DbClick = false;
		function SwitchImg(ImgObj, ParentId) {
			var ImgSrc = "", SubImgSrc;
			ImgSrc = ImgObj.src;
			SubImgSrc = ImgSrc.substr(ImgSrc.length - 5, 5);

			if (SubImgSrc == "+.gif") {
				ImgObj.src = ImgObj.src.replace(SubImgSrc, "-.gif");
				ImgObj.alt = "点击收起子栏目";
				SwitchSub(ParentId, true);
			} else {
				if (SubImgSrc == "-.gif") {
					ImgObj.src = ImgObj.src.replace(SubImgSrc, "+.gif");
					ImgObj.alt = "点击展开子栏目";
					SwitchSub(ParentId, false);
				} else {
					return false;
				}
			}
		}
		function SwitchSub(ParentId, ShowFlag) {
			if ($("Parent" + ParentId).attributes.getNamedItem('HasSub').value == "True") {
				if (ShowFlag) {
					$("Parent" + ParentId).style.display = "";

					if ($("Parent" + ParentId).innerHTML == "" || $("Parent" + ParentId).innerHTML == "栏目加载中...") {
						$("Parent" + ParentId).innerHTML = "栏目加载中...";
						GetSubClass(ParentId);
					}
				} else {
					$("Parent" + ParentId).style.display = "none";
				}
			}
		}
		function SelectLable(Obj) {
			var SelectedInfo = "";
			if (window.SelectSingleClass[0].length > 0) {

				$("class_" + window.SelectSingleClass[0][0]).className = 'LableItem';
			}
			Obj.className = 'LableSelectItem';
			window.SelectSingleClass[0][0] = Obj.id.substring(6);
			window.SelectSingleClass[1][0] = Obj.innerHTML;
			window.SelectSingleClass[2][0] = Obj.attributes.getNamedItem('value').value;
		}
		function GetRootClass() {
			GetSubClass("0");
		}


		function GetSubClass(ParentId) {
			var url = "SelectClass_Ajax.asp";
			var Action = "ParentId=" + ParentId;
			if (document.location.search.length > 0) {
				Action += "&" + document.location.search.substring(1);
			}
			var myAjax = new Ajax.Request(
		url,
		{ method: 'get',
			parameters: Action,
			onComplete: GetSubClassOk
		}
		);
		}
		function GetSubClassOk(OriginalRequest) {
			var ClassInfo;
			if (OriginalRequest.responseText != "" && OriginalRequest.responseText.indexOf("|||") > -1) {
				ClassInfo = OriginalRequest.responseText.split("|||");

				if (ClassInfo[0] == "Succee") {
					$("Parent" + ClassInfo[1]).innerHTML = ClassInfo[2];
				} else {
					$("Parent" + ClassInfo[1]).innerHTML = "<a href=\"点击重试\" onclick=\"$('Parent" + ClassInfo[1] + "').innerHTML='栏目加载中...';GetSubClass('" + ClassInfo[1] + "');return false;\">点击重试</a>";
				}
			} else {
				alert("读取栏目错误.\n请联系管理员.");
				return false;
			}
		}


		function SubmitLable(Obj) {
			SelectLable(Obj);
			window.DbClick = true;
			window.top.close();
		}

		function GetSelectedClass() {
			var selecteds = document.getElementsByName('chkNewsClasses');
			var selectedsArray = [[], []];
			if (!window.DbClick) {
				var mulit = false;
				for (var c = 0; c < selecteds.length; c++) {
					if (selecteds.item(c).checked) {
						mulit = true;
						selectedsArray[0].push(selecteds.item(c).attributes.getNamedItem('id').value);
						selectedsArray[1].push(selecteds.item(c).value);
					}
				}
				if (!mulit) {
					selectedsArray = SelectSingleClass;
				} else {
					selectedsArray[0][0] = selectedsArray[0].join(',');
					selectedsArray[1][0] = selectedsArray[1].join(',');
				}
			} else {
				selectedsArray = SelectSingleClass;
			}
			return selectedsArray;
		}

		function SetReturnValue() {
			window.top.returnValue = GetSelectedClass();
		}
		window.onload = GetRootClass;
		window.onunload = SetReturnValue;
//-->
	</script>

</head>
<body ondragstart="return false;" onselectstart="return false;">
	<div id="Parent0" class="RootItem">
		栏目加载中...
	</div>
</body>
</html>
