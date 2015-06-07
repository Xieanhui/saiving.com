var oAskOnline = document.getElementsByName("askOnline") || null;
if(oAskOnline != null){
	var sUrl = 'http://qiao.baidu.com/v3/?module=default&controller=im&action=index&ucid=7225288&type=n&siteid=4628760';
	var sName = 'ÔÚÏß×ÉÑ¯'; 
	var iWidth = 800; 
	var iHeight = 550; 
	var iTop = (window.screen.height-iHeight)/2; 
	var iLeft = (window.screen.width-iWidth)/2;
	var sParam = 'height='+iHeight+',width='+iWidth+',top='+iTop+',left='+iLeft+',toolbar=no,menubar=no,scrollbars=no, resizable=no,location=no,status=no';
	var iLength = oAskOnline.length;		
	var cssDate = "cursor:pointer;";
	for (var i = 0; i < iLength; i++){			
		if(!+"\v1"){//Õë¶ÔIEä¯ÀÀÆ÷
			oAskOnline.item(i).setAttribute("onclick", function(){window.open(sUrl,sName,sParam)});
			oAskOnline.item(i).style.setAttribute("cssText", cssDate);
		}else{
			oAskOnline.item(i).setAttribute("onclick", "window.open('"+sUrl+"','"+sName+"','"+sParam+"')");
			oAskOnline.item(i).setAttribute("style", cssDate);
		}
		
	}
}