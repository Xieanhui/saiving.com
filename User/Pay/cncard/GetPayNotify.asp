<!-- #include file="MD5.asp" -->
<%
'*******************************************
'�ļ�����GetPayNotify.asp
'��Ҫ���ܣ���ʾ��������Ҫ��ɽ�������֧������֧��֪ͨ��Ϣ����֤��Ϣ��Ч�ԣ��ж�֧��������������߼���������֪ͨ����
'�汾��v1.7��Build2008-09-22��
'˵����
'	1.��ҳ���벻Ҫʹ������response.redirect��ҳ��ת������
'	2.��ֱ�ӽ�֪ͨ���������XML����ʽ����ڱ�ҳ������֧��@�����������������
'	3.������ڱ�ҳ����д����̻��Լ���������򣬽��з�����ϵ�в���
'��Ȩ���У����������������缼�����޹�˾
'����֧����ϵ��ʽ��86-010-82888778-8054 ת������
'*******************************************

'---��ȡ����֧���������̻����͵�֧��֪ͨ��Ϣ(���¼��Ϊ֪ͨ��Ϣ)
	Const PayHandleURL,c_pass
	PayHandleURL	= "http://www.yoursitename.com/urlpath/GetPayHandle.asp"	'�̻����û���ʾ������ҳ���URL(��Ӧ�����ļ���GetPayHandle.asp)	 	
	c_pass			= "Test"					'�̻���֧����Կ����¼�̻������̨(https://www.cncard.net/admin/)���ڹ�����ҳ���ҵ���ֵ

	Dim c_mid,c_order,c_orderamount,c_ymd,c_transnum,c_succmark,c_moneytype,c_cause,c_memo1,c_memo2,c_signstr
	c_mid			= request("c_mid")			'�̻���ţ��������̻��ɹ��󼴿ɻ�ã������������̻��ɹ����ʼ��л�ȡ�ñ��
	c_order			= request("c_order")		'�̻��ṩ�Ķ�����
	c_orderamount	= request("c_orderamount")	'�̻��ṩ�Ķ����ܽ���ԪΪ��λ��С���������λ���磺13.05
	c_ymd			= request("c_ymd")			'�̻���������Ķ����������ڣ���ʽΪ"yyyymmdd"����20050102
	c_transnum		= request("c_transnum")		'����֧�������ṩ�ĸñʶ����Ľ�����ˮ�ţ����պ��ѯ���˶�ʹ�ã�
	c_succmark		= request("c_succmark")		'���׳ɹ���־��Y-�ɹ� N-ʧ��			
	c_moneytype		= request("c_moneytype")	'֧�����֣�0Ϊ�����
	c_cause			= request("c_cause")		'�������֧��ʧ�ܣ����ֵ����ʧ��ԭ��		
	c_memo1			= request("c_memo1")		'�̻��ṩ����Ҫ��֧�����֪ͨ��ת�����̻�����һ
	c_memo2			= request("c_memo2")		'�̻��ṩ����Ҫ��֧�����֪ͨ��ת�����̻�������
	c_signstr		= request("c_signstr")		'����֧�����ض�������Ϣ����MD5���ܺ���ַ���

	'---У����Ϣ������---
		IF c_mid="" or c_order="" or c_orderamount="" or c_ymd="" or c_moneytype="" or c_transnum="" or c_succmark="" or c_signstr="" THEN
			'֧����Ϣ����		
			Call WriteNotice()			
		END IF

	'---����õ�֪ͨ��Ϣƴ���ַ�������Ϊ׼������MD5���ܵ�Դ������Ҫע����ǣ���ƴ��ʱ���Ⱥ�˳���ܸı�
		
		
		srcStr = c_mid & c_order & c_orderamount & c_ymd & c_transnum & c_succmark & c_moneytype & c_memo1 & c_memo2 & c_pass

	'---��֧��֪ͨ��Ϣ����MD5����
		r_signstr	= MD5(srcStr)

	'---У���̻���վ��֪ͨ��Ϣ��MD5���ܵĽ��������֧�������ṩ��MD5���ܽ���Ƿ�һ��
		IF r_signstr<>c_signstr Then
			'ǩ����֤ʧ��			
			Call WriteNotice()			
		END IF

	'---У���̻����
		Dim MerchantID	'�̻��Լ��ı��
		IF MerchantID<>c_mid Then			
			'�ύ���̻��������				
			Call WriteNotice()			
		END IF

	'---У���̻�����ϵͳ���Ƿ���֪ͨ��Ϣ���صĶ�����Ϣ
		Dim conn	'�̻�ϵͳ����������
		sql="select top 1 ������ from �̻��Ķ����� where �̻�������="& c_order
		set rs=server.CreateObject("adodb.recordset")
		rs.open sql,conn
		IF rs.eof THEN
			'δ�ҵ��ö�����Ϣ		
			Call WriteNotice()			
		END IF

	'---У���̻�����ϵͳ�м�¼�Ķ�����������֧������֪ͨ��Ϣ�еĽ���Ƿ�һ��
		Dim r_orderamount	'�̻��Լ�ϵͳ��¼�Ķ������
		r_orderamount=rs("�������")	'�̻����Լ�����ϵͳ��ȡ��ֵ
		IF ccur(r_orderamount)<>ccur(c_orderamount) THEN
			'֧���������
			Call WriteNotice()			
		END IF

	'---У���̻�����ϵͳ�м�¼�Ķ����������ں�����֧������֪ͨ��Ϣ�еĶ������������Ƿ�һ��
		Dim r_ymd	'�̻��Լ�ϵͳ��¼�Ķ�����������
		r_ymd=rs("������������")	'�̻����Լ�����ϵͳ��ȡ��ֵ
		IF r_ymd<>c_ymd THEN
			'����ʱ������			
			Call WriteNotice()			
		END IF

	'---У���̻�ϵͳ�м�¼����Ҫ��֧�����֪ͨ��ת���Ĳ���������֧������֪ͨ��Ϣ���ṩ�Ĳ����Ƿ�һ��
		Dim r_memo1	'�̻��Լ�ϵͳ��¼����Ҫ��֧�����֪ͨ��ת���Ĳ���һ
		r_memo1 = rs("ת������һ")
		Dim r_memo2	'�̻��Լ�ϵͳ��¼����Ҫ��֧�����֪ͨ��ת���Ĳζ�
		r_memo2 = rs("ת��������")
		IF r_memo1<>c_memo1 or r_memo2<>c_memo2 THEN
			'�����ύ����
			Call WriteNotice()			
		END IF

	'---У�鷵�ص�֧������ĸ�ʽ�Ƿ���ȷ
		IF c_succmark<>"Y" and c_succmark<>"N" THEN
			'�����ύ����	
			Call WriteNotice()			
		END IF

	'---���ݷ��ص�֧��������̻������Լ��ķ����Ȳ���
		IF c_succmark="Y" Then			' c_succmark="Y"����֧���ɹ�			
			'�����̻��Լ�������򣬽��з�����ϵ�в���.����˹����г��ִ�����Ҫֹͣ����ļ���ִ�У���ֱ�ӵ���WriteNotice()������̣���Ҫ��ҳ�����������Ϣ��			
		Else
			'�����̻��Լ�������򣬽���֧��ʧ��ʱ��ϵ�в���������˹����г��ִ�����Ҫֹͣ����ļ���ִ�У���ֱ�ӵ���WriteNotice()������̣���Ҫ��ҳ�����������Ϣ��			
		END If	
	'---֪ͨ�����Ѿ��յ�֪ͨ���������յ������ص���Ϣ����Զ����������ص���Ϣ��Ϊ�û���ʾ�����趨��PayHandleURL������
		Call WriteNotice()

	'---���֧�����֪ͨ�����������������Ѿ��յ���֪ͨ
		Sub WriteNotice()
			'<result>��ֵ�̶�Ϊ1����ʾ�̻��ѳɹ��յ����ص�֧���ɹ���֪ͨ��
			'<reURL> ���̻���ʾ���û�������ҳ���URL(��Ӧ�����ļ���GetPayHandle.asp)
			Response.Write("<result>1</result><reURL>" & PayHandleURL & "</reURL>")
			Response.End()
		End Sub	
%>