// JavaScript Document
//�����Ķ���ת��ΪӢ�Ķ��� str:���滻���ַ�������������������������������������������������������������������������������
/*
1.ReplaceDot(str)�����ĵĶ���ת��ΪӢ�ĵ�
2.CheckContentLen(Str,FS_Alert,Len)���ȼ��
3.isNumber(Str,FS_Alert,Msg,isInteger)���ּ��
4.isEmpty(Str,FS_Alert)�Ƿ�Ϊ��
5.isChinese(Str,FS_Alert)�����Ľ����ش���
6.containValue(str,value,FS_Alert)����ָ���ַ������ش���
7.checkMail(str,FS_Alert)�ʼ��Ϸ���
8.valiateDate(str,FS_Alert) ���ںϷ���
9.Do.these()�������
*/
function ReplaceDot(str)
{
	var Obj=document.getElementById(str);
	var oldValue=Obj.value;
	while(oldValue.indexOf("��")!=-1)//Ѱ��ÿһ�����Ķ��ţ����滻
	{
		Obj.value=oldValue.replace("��",",");
		oldValue=Obj.value;
	}
}

//����ַ����ȣ�Str:������ַ���FS_Alert:������Ϣ��ʾ������Len:���Ƴ��ȡ�����������������������������������������������������
function CheckContentLen(Str,FS_Alert,Len)
{
	var Obj=document.getElementById(Str);
	var minLen;
	var maxLen;
	var index=Len.indexOf("-")
	if(index>0)
	{
		minLen=parseInt(Len.substring(0,index))
		maxLen=parseInt(Len.substring(index+1,Len.length))
		if(Obj.value.length<minLen||Obj.value.length>maxLen)
		{
			document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font> ���ȷ�ΧΪ"+Len+"</span>";
			return false;
		}	
	}else if(Obj.value.length>Len)
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>���ȷ�ΧӦС��:"+Len+"</span>";
		return false;
	}
	document.getElementById(FS_Alert).innerHTML=""
	return true;
}
//����ַ��Ƿ�Ϊ���֣�Str:������ַ���FS_Alert:������Ϣ��ʾ������isInteger:�Ƿ�Ϊ����������������������������������������������
function isNumber(Str,FS_Alert,Msg,isInteger)
{
	var Obj=document.getElementById(Str)
	if(Obj.value=='')
	{
		document.getElementById(FS_Alert).innerHTML=""
		return true;
	}
	else if(isNaN(Obj.value)||Obj.value<0)
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>"+Msg+"</span>";
		return false;
	}
	else if(!isNaN(Obj.value)&&Obj.value>=0)
	{
		if(isInteger)
		{
			if(Obj.value.indexOf(".")>=0)//�Ƿ�Ϊ����
			{
				document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>��ʹ������</span>";
				return false;
			}else
			{
				document.getElementById(FS_Alert).innerHTML=""
				return true;
			}
		}
		else
		{
			document.getElementById(FS_Alert).innerHTML=""
			return true;
		}
	}
}
//����ַ��Ƿ�Ϊ�գ�Str:������ַ���FS_Alert:������Ϣ��ʾ��������������������������������������������������
function isEmpty(Str,FS_Alert)
{
	var Obj=document.getElementById(Str);
	var value=Obj.value.replace(/(^\s*)|(\s*$)/g, "");
	if(value=="")
	{

		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>���ݲ���Ϊ��</span>";
		return false;
	}else
	{
		var Str_Len = "";
		var Len_Color = "";
		Str_Len = value.length;
		if (Str_Len <= 50)
		{
			Len_Color = "006600";	
		}
		else if (Str_Len > 50 && Str_Len <= 100)
		{
			Len_Color = "3300FF";	
		}
		else if (Str_Len > 100)
		{
			Len_Color = "FF0000";		
		}
		document.getElementById(FS_Alert).innerHTML="<span class='tx'>������<font style=\"color:#" + Len_Color + ";font-weight:bold;\">" + Str_Len + "</font></span>";
		return true;
	}
}
//����ַ��Ƿ�Ϊ���ģ�Str:������ַ���FS_Alert:������Ϣ��ʾ��������������������������������������������������
function isChinese(Str,FS_Alert)
{ 
	var Number = "0123456789.,abcdefghijklmnopqrstuvwxyz-\/ABCDEFGHIJKLMNOPQRSTUVWXYZ`~!@#$%^&*()_";
	var Obj=document.getElementById(Str);
	for (i = 0; i < Obj.value.length;i++)
	{   
		var c = Obj.value.charAt(i);
		if (Number.indexOf(c) == -1) 
		{
			document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>�벻Ҫʹ�������ַ�</span>";
			return false;
		}
	}
	document.getElementById(FS_Alert).innerHTML="";
	return true
}
//�ж��Ƿ����ָ����ֵ,������������false������ʾ�û��������������������������������������������������������������������
function containValue(str,value,FS_Alert)
{
	var Obj=document.getElementById(str);
	var str=Obj.value;
	var myArray=value.split(',');
	var flag=false;
	for(var i=0;i<myArray.length;i++)
	{
		if(str.indexOf(myArray[i])!=-1)
			flag=true;
	}
	if(flag)
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>�����ʽ�����벻Ҫ����["+value+"]</span>";
		return false;


	}else
	{
		document.getElementById(FS_Alert).innerHTML=""
		return true;
	}
}
//����ʼ��ĺϷ��ԡ���������������������������������������������������������������������������������������������
function checkMail(str,FS_Alert)
{
	var msg="";
	var strr;
	try{
		var mail=document.getElementById(str).value;
		if(mail=="")
		{
			msg="";
		}else{
			re=/(\w+@\w+\.\w+)(\.{0,1}\w*)(\.{0,1}\w*)/i;
			re.exec(mail);
			if (RegExp.$3!=""&&RegExp.$3!="."&&RegExp.$2!=".") strr=RegExp.$1+RegExp.$2+RegExp.$3
			else
			if (RegExp.$2!=""&&RegExp.$2!=".") strr=RegExp.$1+RegExp.$2
			else strr=RegExp.$1
			if (strr!=mail) 
			{
				msg="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>����д��ȷ���ʼ���ַ</span>";
			}
		}
		if (FS_Alert!=""){
			if (msg==""){
				document.getElementById(FS_Alert).innerHTML="";
				return true;
			}else{
				document.getElementById(FS_Alert).innerHTML=msg;
				return false;
			}
		}
		else{
			if (msg==""){
				return true;
			}else{
				return false;
			}
		}
	}
	catch(e){
		return false;
	}
	
}
//������ڵĺϷ��ԡ���������������������������������������������������������������������������������������������������
function valiateDate(str,FS_Alert) 
{
	var valobj=document.getElementById(str);
	var dar=valobj.value.replace(".","-")
	if(dar=="")
	{
		document.getElementById(FS_Alert).innerHTML="";
		return true;
	}
	if (dar=="") return;
	if(dar.split("-")[0].length==2)
	{
		var Current_Date = new Date();
		var Current_year = Current_Date.getYear();
		dar=Current_year.toString().substr(0,2)+dar
	}
	var datePat = /^(\d{4})(-)(\d{1,2})(-)(\d{1,2})$/;

	var matchArray = dar.match(datePat); // is the format ok?
	if (matchArray == null) 
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>����ո�ʽ:��������-����-����!</span>";
		return false;
	}
	month = matchArray[3]; // parse date into variables
	day = matchArray[5];
	year = matchArray[1];
	if (month < 1 || month > 12) 
	{ // check month range
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>�·ݳ���!</span>";
		return false;
	}
	if (day < 1 || day > 31) 
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>���ڳ���!</span>";
		return false;
	}
	if ((month==4 || month==6 || month==9 || month==11) && day==31) 
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>"+month+"��û��31��!</span>";
		return false;
	}
	if (month == 2) { // check for february 29th
	var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
	if (day>29) 
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>2�²��ܳ���29��!</span>";
		return false;
	}
	if (day==29 && !isleap) 
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>"+year + "�겻�����꣬2��û��29��!</span>";
		return false;
	}
	}
	document.getElementById(FS_Alert).innerHTML=""
	return true;
}
//���ʱ��ĺϷ��ԡ���������������������������������������������������������������������������������������������������
function validateTime(str,FS_Alert)
{
	var valobj=document.getElementById(str);
	var timeStr=valobj.value.replace("��",":")
	if (timeStr=="") return false;
	if (timeStr.length==4)
	timeStr=timeStr.substr(0,2)+":"+timeStr.substr(2,2)
	else if (timeStr.length==6)
	timeStr=timeStr.substr(0,2)+":"+timeStr.substr(2,2)+":"+timeStr.substr(4,2)

	valobj.value=timeStr
	var timePat = /^(\d{1,2}):(\d{1,2})(:(\d{1,2}))?$/;

	var matchArray = timeStr.match(timePat);
	if (matchArray == null) 
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>�����ʱ����������ո�ʽ:Сʱ:����!</span>";
		//valobj.value="08:00"
		return false;
	}
	hour = matchArray[1];
	minute = matchArray[2];
	if (timeStr.length==6) var seconde = matchArray[3];
	if (hour < 0  || hour > 23) 
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>Сʱ��������00--23֮��!</span>";
		return false;
	}

	if (minute < 0 || minute > 59) 
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>������������00--59֮��!</span>";
		return false;
	}
	
	if (seconde < 0 || seconde > 59) 
	{
		document.getElementById(FS_Alert).innerHTML="<font style=\"font-family:Webdings;color:red\">x</font><span class='tx'>����������00--59֮��!</span>";
		return false;
	}
	
	document.getElementById(FS_Alert).innerHTML=""
	return true;
}

//�����������ĺϷ��ԡ���������������������������������������������������������������������������������������������������
var Do ={
	these: function() 
	{
		var flag=true;
		for (var i = 1; i < arguments.length; i++) 
		{
			var lambda = arguments[i];
			if(lambda())
				continue;
			flag=false;			
		}
		if(flag)
		{
			document.getElementById(arguments[0]).className="RightInput"
		}else
		{
			document.getElementById(arguments[0]).className="WarnInput"
		}
	}
}