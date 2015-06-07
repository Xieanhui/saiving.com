//��ʡ��ֱϽ�а��´�
var arrAgency = [["\u5317\u4eac","beijing"],//����
["\u6cb3\u5317","beijing"],//�ӱ�
["\u5929\u6d25","beijing"],//���
["\u5185\u8499\u53e4","ruimong"],//���ɹ�
["\u65b0\u7586","xinjiang"],//�½�
["\u7518\u8083","ganshu"],//����
["\u9655\u897f","xianxi"],//����
["\u9752\u6d77","qinghai"],//�ຣ
["\u8fbd\u5b81","jilin"],//����
["\u5409\u6797","jilin"],//����
["\u9ed1\u9f99\u6c5f","jilin"], //������
["\u5b81\u590f","linxia"], //����
["\u6cb3\u5357","henan"],//����
["\u5c71\u897f","shanxi"],//ɽ��
["\u5c71\u4e1c","shandong"],//ɽ��
["\u56db\u5ddd","sichung"],//�Ĵ�
["\u91cd\u5e86","sichung"],//����
["\u6c5f\u82cf","jiangshu"],//����
["\u6e56\u5317","hubei"],//����
["\u5b89\u5fbd","anhui"],//����
["\u4e0a\u6d77","shanghai"],//�Ϻ�
["\u6d59\u6c5f","jiejiang"],//�㽭
["\u6e56\u5357","hunan"],//����
["\u798f\u5efa","fujian"],//����
["\u6c5f\u897f","jiangxi"],//����
["\u5e7f\u897f","guangxi"],//����
["\u5e7f\u4e1c","guangdong"],//�㶫
["\u8d35\u5dde","guijou"],//����
["\u4e91\u5357","yunnan"],//����
["\u6d77\u5357","guangdong"],//����
["\u897f\u85cf","qinghai"],//����
["\u53f0\u6e7e","beijing"],//̨��
["\u9999\u6e2f","beijing"],//���
["\u6fb3\u95e8","beijing"]];//����

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

function setAgency(uAgency, oAgency){	//����ʡ�������������´���
	switch(uAgency)
	{
		case "\u5317\u4eac": //����
			oAgency.innerHTML = "�����ܲ�";	
			break;		
		case "\u6cb3\u5317": //�ӱ�
		case "\u5929\u6d25": //���
		case "\u5185\u8499\u53e4": // ���ɹ�
		case "\u65b0\u7586": //�½�
		case "\u7518\u8083": //����
		case "\u9655\u897f": //����
		case "\u9752\u6d77": //�ຣ
		case "\u8fbd\u5b81": //����
		case "\u5409\u6797": //����
		case "\u9ed1\u9f99\u6c5f": //������
		case "\u5b81\u590f": //����
		case "\u6cb3\u5357": //����
		case "\u5c71\u897f": //ɽ��
		case "\u5c71\u4e1c": //ɽ��
		case "\u56db\u5ddd": //�Ĵ�
		case "\u91cd\u5e86": //����
		case "\u6c5f\u82cf": //����
		case "\u6e56\u5317": //����
		case "\u5b89\u5fbd": //����
		case "\u4e0a\u6d77": //�Ϻ�
		case "\u6d59\u6c5f": //�㽭
		case "\u6e56\u5357": //����
		case "\u798f\u5efa": //����
		case "\u6c5f\u897f": //����
		case "\u5e7f\u897f": //����
		case "\u5e7f\u4e1c": //�㶫
		case "\u8d35\u5dde": //����
		case "\u4e91\u5357": //����
		case "\u6d77\u5357": //����
		case "\u897f\u85cf": //����
			oAgency.innerHTML = uAgency + "���´�";
			break;
		case "\u53f0\u6e7e": //̨��
		case "\u9999\u6e2f": //���
		case "\u6fb3\u95e8": //����
			oAgency.innerHTML = "�۰�̨��������´�";
			break;
		default:
			oAgency.innerHTML = "�۰�̨��������´�"; 
			break;
	}		
}

function setAnchorLink(uAgency, arrAgency, oAgency, oContactUs, oOffice){ //����ʡ����������

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
		
	if(cookieAgency){	//���COOKIE�б�����ʡ��
		
		setAgency(cookieAgency, oAgency); //��ʾʡ�ݰ��´�
		setAnchorLink(cookieAgency, arrAgency, oAgency, oContactUs, oOffice);		//����ê����
				
	}else{		//���COOKIE��û�б���ʡ��
		
			try{					
					$LAB.script("http://int.dpool.sina.com.cn/iplookup/iplookup.php?format=js").wait(function(){	
						
						if(remote_ip_info["province"] == null || remote_ip_info["province"] == "\u4fdd\u7559\u5730\u5740" || remote_ip_info["province"] == "") {
							setAgency("\u5317\u4eac", oAgency);	 //��ʾʡ�ݰ��´�						
							setAnchorLink("\u5317\u4eac", arrAgency, oAgency, oContactUs, oOffice);		//����ê����										
							setCookie("agency", "\u5317\u4eac"); //����COOKIE
						} else {
							setAgency(remote_ip_info["province"], oAgency);	 //��ʾʡ�ݰ��´�						
							setAnchorLink(remote_ip_info["province"], arrAgency, oAgency, oContactUs, oOffice);		//����ê����										
							setCookie("agency", remote_ip_info["province"]); //����COOKIE
						}
						
					});																	
				
			}catch(err){	
			
				try {
					
					$LAB.script("http://counter.sina.com.cn/ip/").wait(function(){
						
						if(ILData[1] == "\u4fdd\u7559\u5730\u5740" || ILData[1] == null || ILData[1] == "") {
								setAgency("\u5317\u4eac", oAgency);	 //��ʾʡ�ݰ��´�						
								setAnchorLink("\u5317\u4eac", arrAgency, oAgency, oContactUs, oOffice);		//����ê����										
								setCookie("agency", "\u5317\u4eac"); //����COOKIE
						} else {
							setAgency(ILData[1], oAgency);	 //��ʾʡ�ݰ��´�						
							setAnchorLink(ILData[1], arrAgency, oAgency, oContactUs, oOffice);		//����ê����										
							setCookie("agency", ILData[1]); //����COOKIE
						}
					});	
					
				} catch (err) {
					
					$LAB.script("http://pv.sohu.com/cityjson").wait(function(){
						
						setAgency(returnCitySN["cname"], oAgency);	 //��ʾʡ�ݰ��´�						
						setAnchorLink(returnCitySN["cname"], arrAgency, oAgency, oContactUs, oOffice);		//����ê����										
						setCookie("agency", returnCitySN["cname"]); //����COOKIE					
								
					});	
					
				}				
										
			}
	}	

}

showAgency();