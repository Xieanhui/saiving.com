// JavaScript Document
//�����Ķ���ת��ΪӢ�Ķ��� str:���滻���ַ�
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

//����ַ����ȣ�Str:������ַ���FS_Alert:������Ϣ��ʾ������Len:���Ƴ���
function CheckContentLen(Str,FS_Alert,Len)
{
	var Obj=document.getElementById(Str);
	if(Obj.value.length>Len)
	{
		document.getElementById(FS_Alert).innerHTML="<font color='F43631'>�����벻Ҫ����"+Len+"</font>";
		return false;
	}
	return true;
}
//����ַ��Ƿ�Ϊ���֣�Str:������ַ���FS_Alert:������Ϣ��ʾ������isInteger:�Ƿ�Ϊ����
function isNumber(Str,FS_Alert,Msg,isInteger)
{
	var Obj=document.getElementById(Str)
	if(Obj.value=='')
	{
		document.getElementById(FS_Alert).innerHTML="<font color='F43631'>�ô�����Ϊ��</font>";
		return false;
	}
	else if(isNaN(Obj.value)||Obj.value<0)
	{
		document.getElementById(FS_Alert).innerHTML="<font color='F43631'>"+Msg+"</font>";
		return false;
	}
	else if(!isNaN(Obj.value)&&Obj.value>=0)
	{
		if(isInteger)
		{
			if(Obj.value.indexOf(".")>=0)//�Ƿ�Ϊ����
			{
				document.getElementById(FS_Alert).innerHTML="<font color='F43631'>��ʹ������</font>";
				return false;
			}else
			{
				document.getElementById(FS_Alert).innerHTML="";
				return true;
			}
		}
		else
		{
			document.getElementById(FS_Alert).innerHTML="";
			return true;
		}
	}
}
//����ַ��Ƿ�Ϊ�գ�Str:������ַ���FS_Alert:������Ϣ��ʾ������
function isEmpty(Str,FS_Alert)
{
	var Obj=document.getElementById(Str);
	if(Obj.value=="")
	{
		document.getElementById(FS_Alert).innerHTML="<font color='F43631'>�ô�����Ϊ��</font>";
		return false;
	}else
	{
		document.getElementById(FS_Alert).innerHTML="";
		return true;
	}
}
//����ַ��Ƿ�Ϊ���ģ�Str:������ַ���FS_Alert:������Ϣ��ʾ������
function isChinese(Str,FS_Alert)
{ 
	var Number = "0123456789.,abcdefghijklmnopqrstuvwxyz-\/ABCDEFGHIJKLMNOPQRSTUVWXYZ`~!@#$%^&*()_";
	var Obj=document.getElementById(Str);
	for (i = 0; i < Obj.value.length;i++)
	{   
		var c = Obj.value.charAt(i);
		if (Number.indexOf(c) == -1) 
		{
			document.getElementById(FS_Alert).innerHTML="<font color='F43631'>�벻Ҫʹ�������ַ�</font>";
			return false;
		}
	}
	document.getElementById(FS_Alert).innerHTML="";
	return true
}