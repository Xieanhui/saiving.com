// JavaScript Document
var scrollTop = null;
var nav;
var pos = null;
nav = document.getElementById("navigate");
window.onscroll= scroll_ad;
function scroll_ad(){
	scrollTop = document.documentElement.scrollTop || document.body.scrollTop;
	pos = scrollTop - nav.offsetTop;
	if(pos > 0){
		nav.style.position = "fixed";
		nav.style.top = "0px";	
		if (document.body.clientWidth<940)
			nav.style.left = "0px";
		else
			nav.style.left = (document.body.clientWidth/2-nav.clientWidth/2)+"px"; 
	}
	else if (pos <= 0){
		nav.style.position = "relative";
		nav.style.left = "0px";	
	}
}

window.onresize = resize_ad;
function resize_ad() {
	if (document.body.clientWidth>940){
			nav.style.width = "100%";
		}else {
			nav.style.width = "940px";
		}
	}
