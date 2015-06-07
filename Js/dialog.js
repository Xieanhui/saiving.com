// JavaScript Document

function setCookie()
{
	var expDate = new Date();	
	expDate.setTime(expDate.getTime() + 4*60*60*1000);
	document.cookie = "isShow=123; expires="+ expDate.toGMTString()+"; domain=saiving.com; path=/";
}

function getCookie(objName)
{
	var aStr = document.cookie;
	var aCookie = aStr.split("; ");
	for(var i=0; i< aCookie.length; i++)
	{
		var aTmp = aCookie[i].split("=");	
		if (aTmp[0]==objName) 
		{				
			return aTmp[1];	
		}					
	}
}

window.onload = function () {
	
	var oDialogBg;
	var oDialog;
	var oClose;
	
	var oClientWidth;
	var oClientHeight;
	var oScrollTop;
	var oScrollLeft;
	
	var isShow = false;

	oDialogBg = document.getElementById("dialogBg");
	oDialog = document.getElementById("dialog");		
	oClose = oDialog.getElementsByTagName("span");
	
	oClientWidth = document.documentElement.clientWidth || document.body.clientWidth;
	oClientHeight = document.documentElement.clientHeight || document.body.clientHeight;
	oScrollTop = document.documentElement.scrollTop || document.body.scrollTop;
	oScrollLeft = document.documentElement.scrollLeft || document.body.scrollLeft;
	
	oDialogBg.style.width = oClientWidth +"px";
	oDialogBg.style.height = oClientHeight +"px";
		
	oClose[0].onclick = function (){
		oDialogBg.style.display = 'none';
		oDialog.style.display = 'none';
		};	
	oClose[0].onmouseover = function(){
		oClose[0].style.background = 'red';
	};
	oClose[0].onmouseout = function(){
		oClose[0].style.background = '#333';
	};
			
	isShow = getCookie('isShow');

	if (isShow)
	{
		//不显示对话框；
		oDialogBg.style.display = 'none';
		oDialog.style.display = 'none';	
	}
	else
	{
		//设置Cookie;并显示对话框;
		setCookie();		
	
		oDialogBg.style.display = 'block';
		oDialog.style.display = 'block';	
		
		oDialog.style.top =  (oClientHeight/2 - oDialog.offsetHeight/2) + "px";
		oDialog.style.left = (oClientWidth/2 - oDialog.offsetWidth/2) + "px";		
	}	
};