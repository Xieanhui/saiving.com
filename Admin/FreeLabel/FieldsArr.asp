<%
'=======================================
'�������ݱ��е��ֶ����Ͷ�Ӧ������������
'��ѯ�ֶ�
'�����ֶ�
'��ѯ�����ֶ�
'=======================================
Dim GetStrFun
If G_IS_SQL_DB = 1 Then
	GetStrFun = "SubString"
Else
	GetStrFun = "Mid"
End IF

'=======================================
' ���ű�
'=======================================
Dim NSAllFCNArr,NSAllFENArr,NSAllFTypeArr,NSConArr,NSOrderArr
Dim NewAllENameFields,NewAllENameFieldsType
'---�������ű��ֶ�������
NSAllFCNArr = Array("���","���ű��","Ȩ��","������Ŀ���","����ר��","���ű���","������","���ݵ���","������ʾ����","������ɫ","�����Ƿ����","�����Ƿ�б��","��������","�������ŵ�ַ","��������","�Ƿ�ͼƬ����","��ͼ��ַ","Сͼ��ַ","ģ���ַ","���Ȩ��","��Դ","�༭","�ؼ���","����","�����","����·��","�����ļ���","�ļ���չ��","ͼƬͷ��","�Ƿ�����","�Ƿ�ɾ��","���ʱ��","�Ƿ�ݸ�","���л����","�����","���߶�","�������","����ַ","�Ƽ�","����","��������","Ͷ��","Զ�̴�ͼ","����ͷ��","�ȵ�","����","������","����","�õ�")
'---�������ű��ֶ��ֶ���
NewAllENameFields = "ID||NewsID||PopId||ClassID||SpecialEName||NewsTitle||CurtTitle||NewsNaviContent||isShowReview||TitleColor||titleBorder||TitleItalic||IsURL||URLAddress||Content||isPicNews||NewsPicFile||NewsSmallPicFile||Templet||isPop||Source||Editor||Keywords||Author||Hits||SaveNewsPath||FileName||FileExtName||TodayNewsPic||isLock||isRecyle||addtime||isdraft||IsAdPic||" & GetStrFun & "(AdPicWH,1,1)||" & GetStrFun & "(AdPicWH,3,1)||AdPicLink||AdPicAdress||" & GetStrFun & "(NewsProperty,1,1)||" & GetStrFun & "(NewsProperty,3,1)||" & GetStrFun & "(NewsProperty,5,1)||" & GetStrFun & "(NewsProperty,7,1)||" & GetStrFun & "(NewsProperty,9,1)||" & GetStrFun & "(NewsProperty,11,1)||" & GetStrFun & "(NewsProperty,13,1)||" & GetStrFun & "(NewsProperty,15,1)||" & GetStrFun & "(NewsProperty,17,1)||" & GetStrFun & "(NewsProperty,19,1)||" & GetStrFun & "(NewsProperty,21,1)"
NSAllFENArr = Split(NewAllENameFields,"||")
'---�������ű��ֶ�����
NewAllENameFieldsType = "����,�ı���ID,����,�ı���ID,�ı���ID,�ı�,�ı�,��ע,�ж�������,�ı�,�ж�������,�ж�������,�ж�������,�ı�,��ע,�ж�������,�ı�,�ı�,�ı�,�ж�������,�ı�,�ı�,�ı�,�ı�,����,�ı�,�ı�,�ı�,�ж�������,�ж�������,�ж�������,����ʱ����,�ж�������,�ж�������,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�"
NSAllFTypeArr = Split(NewAllENameFieldsType,",")
'---����������ѯ�����ű��ֶ�
NSConArr = Array("���","���ű��","���ű���","������","���ݵ���","������ɫ","�������ŵ�ַ","��������","��ͼ��ַ","Сͼ��ַ","ģ���ַ","��Դ","�༭","�ؼ���","����","�����","���ʱ��","�������","����ַ","����ר��")
'---����������������ű��ֶ�
NSOrderArr = Array("���","Ȩ��","�����","���ʱ��")

'=======================================
' ������Ŀ��
'=======================================
Dim NS_CAllFCNArr,NS_CAllENArr,NS_CAllTypeArr,NS_CConArr,NS_COrderArr
Dim NCAllENameFields,NCAllENameFieldsType
'---������Ŀ�������ֶ�������
NS_CAllFCNArr = Array("���","��Ŀ���","��ĿȨ��","��Ŀ������","��ĿӢ����","�Ƿ��ⲿ��Ŀ","�ⲿ��Ŀ��ַ","����Ŀ���","��Ŀģ��","��Ŀ����ģ��","��������","��Ŀ����ԱID","���Ȩ��","�ļ���չ��","���ʱ��","����Ͷ��","�鵵ʱ��","��Ŀ������ʾ","ˢ����Ϣ����","����˵��","����ͼƬ","�Զ����ֶη���ID","����Ĭ�����","�������ģʽ","����·��","�����ʽ","ɾ������Ͷ��","��Ŀ�ؼ���","��Ŀ����","�Ƿ�ɾ��","���л����","�����","���߶�","�������","����ַ")
'---������Ŀ�������ֶ���
NCAllENameFields = "ID||ClassID||OrderID||ClassName||ClassEName||IsURL||UrlAddress||ParentID||Templet||NewsTemplet||Domain||ClassAdmin||isPop||FileExtName||Addtime||isConstr||Oldtime||isShow||RefreshNumber||ClassNaviContent||ClassNaviPic||DefineID||NewsCheck||AddNewsType||SavePath||FileSaveType||isConstrDel||ClassKeywords||Classdescription||ReycleTF||IsAdPic||" & GetStrFun & "(AdPicWH,1,1)||" & GetStrFun & "(AdPicWH,3,1)||AdPicLink||AdPicAdress"
NS_CAllENArr = Split(NCAllENameFields,"||")
'---������Ŀ�������ֶ�����
NCAllENameFieldsType = "����,�ı���ID,����,�ı�,�ı�,�ж�������,�ı�,�ı���ID,�ı�,�ı�,�ı�,�ı���ID,�ж�������,�ı�,����ʱ����,�ж�������,����,�ж�������,����,��ע,�ı�,����,�ж�������,�ж�������,�ı�,�ж�������,�ж�������,�ı�,�ı�,�ж�������,�ж�������,�ı�,�ı�,�ı�,�ı�"
NS_CAllTypeArr = Split(NCAllENameFieldsType,",")
'---��ѯ��������Ŀ���ֶ�
NS_CConArr = Array("���","��Ŀ���","��Ŀ������","��ĿӢ����","�ⲿ��Ŀ��ַ","����Ŀ���","��Ŀģ��","��Ŀ����ģ��","��������","���ʱ��","����˵��","����ͼƬ","����·��","��Ŀ�ؼ���","��Ŀ����","�����","���߶�","�������","����ַ")
'---��������ű��ֶ�
NS_COrderArr = Array("���","��ĿȨ��","���ʱ��")


'=======================================
'���ر�
'=======================================
Dim DSAllFCNArr,DSAllFENArr,DSAllFTypeArr,DSConArr,DSOrderArr
Dim DownAllENameFields,DownAllENameFieldsType 
'---���ر��������ֶ�������
DSAllFCNArr = Array("���","���ر��","������Ŀ���","���","��Ȩ","���ʱ��","�Ǽ�","�Ƿ����","���Ȩ��","���ش���","�޸�ʱ��","�ṩ��EMAIL","�ļ���չ��","�ļ���","�����С","���԰汾","��������","����ģ��","��ѹ����","ͼƬ��ַ","��������","������","�ṩ�ߵ�ַ","�Ƽ�����","��������","������Ҫ���","ϵͳƽ̨","��������","����汾","��������","���ѵ���","����·��","�������","����ר��")
'---���ر��������ֶ�Ӣ����
DownAllENameFields = "ID||DownLoadID||ClassID||Description||Accredit||AddTime||Appraise||AuditTF||BrowPop||ClickNum||EditTime||EMail||FileExtName||FileName||FileSize||Language||Name||NewsTemplet||PassWord||Pic||Property||Provider||ProviderUrl||RecTF||ReviewTF||ShowReviewTF||SystemType||Types||Version||OverDue||ConsumeNum||SavePath||Hits||SpeicalID"
DSAllFENArr = Split(DownAllENameFields,"||")
'---���ر������ֶ�����
DownAllENameFieldsType = "����,�ı���ID,�ı���ID,��ע,�ж�������,����ʱ����,�ж�������,�ж�������,�ı�,����,����ʱ����,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ж�������,�ı�,�ı�,�ж�������,�ж�������,�ж�������,�ı�,�ж�������,�ı�,����,����,�ı�,����,����,�ı�"
DSAllFTypeArr = Split(DownAllENameFieldsType,",")
'---��ѯ�����ر��ֶ�
DSConArr = Array("���","���ر��","���","��Ȩ","���ʱ��","�Ǽ�","���ش���","�޸�ʱ��","�ṩ��EMAIL","�����С","���԰汾","��������","��ѹ����","ͼƬ��ַ","��������","������","�ṩ�ߵ�ַ","ϵͳƽ̨","��������","����汾","��������","���ѵ���","�������","�Ƽ�����","����ר��")
'---��������ر��ֶ�
DSOrderArr = Array("���","���ʱ��","���ش���","�޸�ʱ��","���ѵ���","�������")

'=======================================
'������Ŀ��
'=======================================
Dim DCAllFCNArr,DCAllFENArr,DCAllFTypeArr,DCConArr,DCOrderArr
Dim D_CAllENameFields,D_CAllENameFieldsType
'---������Ŀ�������ֶ�������
DCAllFCNArr = Array("���","��Ŀ���","��ĿȨ��","��Ŀ������","��ĿӢ����","�Ƿ��ⲿ��Ŀ","�ⲿ��Ŀ��ַ","����Ŀ���","��Ŀģ��","����ģ��","��������","��Ŀ����Ա���","���Ȩ��","�ļ���չ��","���ʱ��","����Ͷ��","������ʾ","ˢ������","��������","����ͼƬ","�Զ����ֶα��","Ĭ�����","����·��","��ҳ��������","Ͷ������ɾ��","��Ŀ�ؼ���","��Ŀ����","�Ƿ�ɾ��")
'---������Ŀ�������ֶ�Ӣ����
D_CAllENameFields = "ID||ClassID||OrderID||ClassName||ClassEName||IsURL||UrlAddress||ParentID||Templet||NewsTemplet||Domain||ClassAdmin||isPop||FileExtName||Addtime||isConstr||isShow||RefreshNumber||ClassNaviContent||ClassNaviPic||DefineID||NewsCheck||SavePath||FileSaveType||isConstrDel||ClassKeywords||Classdescription||ReycleTF"
DCAllFENArr = Split(D_CAllENameFields,"||")
'---������Ŀ�������ֶ�����
D_CAllENameFieldsType = "����,�ı���ID,����,�ı�,�ı�,�ж�������,�ı�,�ı���ID,�ı�,�ı�,�ı�,�ı���ID,�ж�������,�ı�,����ʱ����,�ж�������,�ж�������,����,��ע,�ı�,����,�ж�������,�ı�,�ж�������,�ж�������,�ı�,�ı�,�ж�������"
DCAllFTypeArr = Split(D_CAllENameFieldsType,",")
'---������Ŀ���ѯ�ֶ�
DCConArr = Array("���","��Ŀ���","��Ŀ������","��ĿӢ����","�ⲿ��Ŀ��ַ","��Ŀģ��","����ģ��","��������","���ʱ��","��������","����ͼƬ","��Ŀ�ؼ���","��Ŀ����")
'---������Ŀ�������ֶ�
DCOrderArr = Array("���","��ĿȨ��","���ʱ��")


'=======================================
'��Ʒ��
'=======================================
Dim MSAllFCNArr,MSAllFENArr,MSAllFTypeArr,MSConArr,MSOrderArr
Dim MallAllENameFields,MallAllENameFieldsType
'---��Ʒ�������ֶ�������
MSAllFCNArr = Array("���","��Ʒ����","������","��Ʒ���к�","������Ŀ���","�ؼ���","����ר��ID","���","��澯��","ԭ��","�ּ�","��������","��Ʒ����","��������","��Ʒģ��","������","����","�Ƿ��з�Ʊ","�����","��������","����·��","�ļ���","�ļ���չ��","Сͼ��ַ","��ͼ��ַ","�Ƽ�","�ȵ�","�ؼ�","����","����","�õ�","����","��������","�������","������","���ۿ�ʼ����","���۽�������","�Ƿ�ɾ��","���������","�۳�����","��ƷȨ��","��ʾ����")
'---��Ʒ�������ֶ�Ӣ����
MallAllENameFields = "ID||ProductTitle||Barcode||Serialnumber||ClassID||Keyword||SpecialID||Stockpile||StockpileWarn||OldPrice||NewPrice||" & GetStrFun & "(IsWholesale,0,1)||ProductContent||RepairContent||TempletFile||MakeFactory||ProductsAddress||IsInvoice||Click||MakeTime||SavePath||FileName||FileExtName||smallPic||BigPic||" & GetStrFun & "(StyleFlagBit,1,1)||" & GetStrFun & "(StyleFlagBit,3,1)||" & GetStrFun & "(StyleFlagBit,5,1)||" & GetStrFun & "(StyleFlagBit,7,1)||" & GetStrFun & "(StyleFlagBit,9,1)||" & GetStrFun & "(StyleFlagBit,11,1)||" & GetStrFun & "(StyleFlagBit,13,1)||SaleStyle||AddTime||Discount||DiscountStartDate||DiscountEndDate||ReycleTF||AddMember||saleNumber||popid||isShowReview"
MSAllFENArr = Split(MallAllENameFields,"||")
'---��Ʒ�������ֶ�����
MallAllENameFieldsType = "����,�ı�,�ı�,�ı�,�ı���ID,�ı�,����ID,�ı�,����,����,����,�ж�������,��ע,��ע,�ı�,�ı�,�ı�,�ж�������,����,����ʱ����,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ı�,�ж�������,����ʱ����,����,����ʱ����,����ʱ����,�ж�������,�ı�,����,�ж�������,�ж�������"
MSAllFTypeArr = Split(MallAllENameFieldsType,",")
'---��Ʒ���ѯ�ֶ�
MSConArr = Array("���","��Ʒ����","������","��Ʒ���к�","�ؼ���","���","��澯��","ԭ��","�ּ�","��������","��Ʒ����","��������","������","����","�Ƿ��з�Ʊ","�����","��������","Сͼ��ַ","��ͼ��ַ","��������","�������","������","���ۿ�ʼ����","���۽�������","���������","�۳�����","����ר��ID")
'---��Ʒ�������ֶ�
MSOrderArr = Array("���","ԭ��","�ּ�","�����","��������","�������","������","���ۿ�ʼ����","���۽�������","�۳�����","��ƷȨ��")


'=======================================
'��Ʒ��Ŀ��
'=======================================
Dim MCAllFCNArr,MCAllFENArr,MCAllFTypeArr,MCConArr,MCOrderArr
Dim M_CAllENameFields,M_CAllENameFieldsType
'---��Ʒ��Ŀ�������ֶ�������
MCAllFCNArr = Array("���","��Ŀ���","����Ŀ���","��ĿȨ��","��ĿӢ����","��Ŀ������","��Ŀģ��","��Ʒģ��","�Ƿ��ⲿ��Ŀ","�ⲿ��Ŀ��ַ","�������","��������","���Ȩ��","�Ƿ�̳и���Ŀ����","��Ŀ����Ա","������ʾ","��Ŀ��������","��Ŀ����ͼƬ","�Զ����ֶ�ID","��Ŀ�ؼ���","��Ŀ����","�Ƿ�ɾ��","�ļ���չ��","����·��","�ļ���������")
'---��Ʒ��Ŀ�������ֶ�Ӣ����
M_CAllENameFields = "ID||ClassID||ParentID||OrderID||ClassEName||ClassCName||ClassTemplet||ProductsTemplet||IsUrl||UrlAddress||Addtime||Domain||IsLimited||IsInherit||ClassAdmin||NaviShow||NaviContent||NaviPic||DefineID||Keywords||Description||ReycleTF||FileExtName||SavePath||FileSaveType"
MCAllFENArr = Split(M_CAllENameFields,"||")
'---��Ʒ��Ŀ�������ֶ�����
M_CAllENameFieldsType = "����,�ı���ID,�ı���ID,����,�ı�,�ı�,�ı�,�ı�,�ж�������,�ı�,����ʱ����,�ı�,�ж�������,�ж�������,�ж�������,�ж�������,��ע,�ı�,����ID,��ע,��ע,�ж�������,�ı�,�ı�,�ж�������"
MCAllFTypeArr = Split(M_CAllENameFieldsType,",")
'---��Ʒ��Ŀ���ѯ�ֶ�
MCConArr = Array("���","��Ŀ���","��ĿӢ����","��Ŀ������","��Ŀģ��","��Ʒģ��","�ⲿ��Ŀ��ַ","�������","��������","��Ŀ��������","��Ŀ����ͼƬ","��Ŀ�ؼ���","��Ŀ����")
'---��Ʒ��Ŀ�������ֶ�
MCOrderArr = Array("���","��ĿȨ��","�������")

'=======================================
' �����ֶ����������е�λ��
'=======================================
Function GetInnerFieldsNum(FieldCName,FieldCNameArray)
	Dim i,FiledName
	FiledName = FieldCName & ""
	For i = 0 to UBound(FieldCNameArray)
		if FieldCNameArray(i) & "" = FiledName Then
			GetInnerFieldsNum = i
			Exit For
			Exit Function
		End if
	Next
	GetInnerFieldsNum = GetInnerFieldsNum
End Function
%>







