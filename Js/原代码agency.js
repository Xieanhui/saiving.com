//各省及直辖市办事处
var arrAgency = [["\u5317\u4eac","beijing"],//北京
["\u6cb3\u5317","beijing"],//河北
["\u5929\u6d25","beijing"],//天津
["\u5185\u8499\u53e4","ruimong"],//内蒙古
["\u65b0\u7586","xinjiang"],//新疆
["\u7518\u8083","ganshu"],//甘肃
["\u9655\u897f","xianxi"],//陕西
["\u9752\u6d77","qinghai"],//青海
["\u8fbd\u5b81","jilin"],//辽宁
["\u5409\u6797","jilin"],//吉林
["\u9ed1\u9f99\u6c5f","jilin"], //黑龙江
["\u5b81\u590f","linxia"], //宁夏
["\u6cb3\u5357","henan"],//河南
["\u5c71\u897f","shanxi"],//山西
["\u5c71\u4e1c","shandong"],//山东
["\u56db\u5ddd","sichung"],//四川
["\u91cd\u5e86","sichung"],//重庆
["\u6c5f\u82cf","jiangshu"],//江苏
["\u6e56\u5317","hubei"],//湖北
["\u5b89\u5fbd","anhui"],//安徽
["\u4e0a\u6d77","shanghai"],//上海
["\u6d59\u6c5f","jiejiang"],//浙江
["\u6e56\u5357","hunan"],//湖南
["\u798f\u5efa","fujian"],//福建
["\u6c5f\u897f","jiangxi"],//江西
["\u5e7f\u897f","guangxi"],//广西
["\u5e7f\u4e1c","guangdong"],//广东
["\u8d35\u5dde","guijou"],//贵州
["\u4e91\u5357","yunnan"],//云南
["\u6d77\u5357","guangdong"],//海南
["\u897f\u85cf","qinghai"],//西藏
["\u53f0\u6e7e","beijing"],//台湾
["\u9999\u6e2f","beijing"],//香港
["\u6fb3\u95e8","beijing"]];//澳门

function setCookie(name,value){
	var expDate = new Date(); 
	document.cookie = name + "="+ escape(value) + "; expires="+ expDate.setTime(expDate.getTime + 24*60*60*1000*1); 
}

function getCookie(name){  
        var arr = document.cookie.match(new RegExp("(^|)"+name+"=([^;]*)(;|$)"));  
        if(arr != null){  
            return unescape(arr[2]);  
        }else{  
            return "";  
        }  
}

function delCookie(name){  
        var exp = new Date();  
        exp.setTime(exp.getTime() - 1);  
        var cval=getCookie(name);  
        if(cval!=null) document.cookie= name + "="+cval+";expires="+exp.toGMTString();  
}

function setAgency(uAgency, oAgency){	//根据省份名称来填充办事处。
	switch(uAgency)
	{
		case "\u5317\u4eac": //北京
			oAgency.innerHTML = "北京总部";	
			break;		
		case "\u6cb3\u5317": //河北
		case "\u5929\u6d25": //天津
		case "\u5185\u8499\u53e4": // 内蒙古
		case "\u65b0\u7586": //新疆
		case "\u7518\u8083": //甘肃
		case "\u9655\u897f": //陕西
		case "\u9752\u6d77": //青海
		case "\u8fbd\u5b81": //吉林
		case "\u5409\u6797": //辽宁
		case "\u9ed1\u9f99\u6c5f": //黑龙江
		case "\u5b81\u590f": //宁夏
		case "\u6cb3\u5357": //河南
		case "\u5c71\u897f": //山西
		case "\u5c71\u4e1c": //山东
		case "\u56db\u5ddd": //四川
		case "\u91cd\u5e86": //重庆
		case "\u6c5f\u82cf": //江苏
		case "\u6e56\u5317": //湖北
		case "\u5b89\u5fbd": //安徽
		case "\u4e0a\u6d77": //上海
		case "\u6d59\u6c5f": //浙江
		case "\u6e56\u5357": //湖南
		case "\u798f\u5efa": //福建
		case "\u6c5f\u897f": //江西
		case "\u5e7f\u897f": //广西
		case "\u5e7f\u4e1c": //广东
		case "\u8d35\u5dde": //贵州
		case "\u4e91\u5357": //云南
		case "\u6d77\u5357": //海南
		case "\u897f\u85cf": //西藏
			oAgency.innerHTML = uAgency + "办事处";
			break;
		case "\u53f0\u6e7e": //台湾
		case "\u9999\u6e2f": //香港
		case "\u6fb3\u95e8": //澳门
			oAgency.innerHTML = "港澳台及海外办事处";
			break;
		default:
			oAgency.innerHTML = "港澳台及海外办事处"; 
			break;
	}		
}

function setAnchorLink(uAgency, arrAgency, oAgency, oContactUs, oOffice){ //根据省份设置链接

	var i;	
	
	for (i=0; i<arrAgency.length; i++){		
		if (arrAgency[i][0] == uAgency){
				oAgency.setAttribute("href","http://www.saiving.com/about/contact/index.html#"+arrAgency[i][1]);
				oContactUs.setAttribute("href","http://www.saiving.com/about/contact/index.html#"+arrAgency[i][1]);
				if(oOffice==null) {}
				else {
					oOffice.setAttribute("href","http://www.saiving.com/about/contact/index.html#"+arrAgency[i][1]);
					}
				break;
		} 								
	}		
}

function showAgency(){
	
	var oAgency = document.getElementById("agency");
	var oContactUs = document.getElementById("contactUs");
	var oOffice = document.getElementById("office") || null;
	var cookieAgency = getCookie("agency");	
		
	if(cookieAgency){	//如果COOKIE中保存有省份
		
		setAgency(cookieAgency, oAgency); //显示省份办事处
		setAnchorLink(cookieAgency, arrAgency, oAgency, oContactUs, oOffice);		//设置锚链接
				
	}else{		//如果COOKIE中没有保存省份
		
			try{					
					$LAB.script("http://int.dpool.sina.com.cn/iplookup/iplookup.php?format=js").wait(function(){	
						
						if(remote_ip_info["province"] == null || remote_ip_info["province"] == "\u4fdd\u7559\u5730\u5740" || remote_ip_info["province"] == "") {
							setAgency("\u5317\u4eac", oAgency);	 //显示省份办事处						
							setAnchorLink("\u5317\u4eac", arrAgency, oAgency, oContactUs, oOffice);		//设置锚链接										
							setCookie("agency", "\u5317\u4eac"); //设置COOKIE
						} else {
							setAgency(remote_ip_info["province"], oAgency);	 //显示省份办事处						
							setAnchorLink(remote_ip_info["province"], arrAgency, oAgency, oContactUs, oOffice);		//设置锚链接										
							setCookie("agency", remote_ip_info["province"]); //设置COOKIE
						}
						
					});																	
				
			}catch(err){	
			
				try {
					
					$LAB.script("http://counter.sina.com.cn/ip/").wait(function(){
						
						if(ILData[1] == "\u4fdd\u7559\u5730\u5740" || ILData[1] == null || ILData[1] == "") {
								setAgency("\u5317\u4eac", oAgency);	 //显示省份办事处						
								setAnchorLink("\u5317\u4eac", arrAgency, oAgency, oContactUs, oOffice);		//设置锚链接										
								setCookie("agency", "\u5317\u4eac"); //设置COOKIE
						} else {
							setAgency(ILData[1], oAgency);	 //显示省份办事处						
							setAnchorLink(ILData[1], arrAgency, oAgency, oContactUs, oOffice);		//设置锚链接										
							setCookie("agency", ILData[1]); //设置COOKIE
						}
					});	
					
				} catch (err) {
					
					$LAB.script("http://pv.sohu.com/cityjson").wait(function(){
						
						setAgency(returnCitySN["cname"], oAgency);	 //显示省份办事处						
						setAnchorLink(returnCitySN["cname"], arrAgency, oAgency, oContactUs, oOffice);		//设置锚链接										
						setCookie("agency", returnCitySN["cname"]); //设置COOKIE					
								
					});	
					
				}				
										
			}
	}	

}

showAgency();