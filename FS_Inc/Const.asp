<%
'++++++++++----------������Ϣ��ʼ----------++++++++++
'�벻Ҫ�����޸Ĵ��ļ�
'������"'"����ȥ��
'�벻Ҫʹ�ûس�����
'������ȫ�ֱ������˽⣬����ϸ�鿴���������ĳ���ʹ��˵��

'-----ϵͳ������Ŀ¼�����治�ܴ�/�����磺F5��Test/F5
Const G_VIRTUAL_ROOT_DIR		= ""

'-----���ﳵģ��·��{FS:Mall_Cart_Content}
Const G_MALL_CART_TEMPLET       ="/Templets/Mall/Cart.html"

'-----����Ŀ¼�����治�ܴ�/����������Ŀ¼
Const G_ADMIN_DIR				= "Admin"

'-----�û�Ŀ¼�����治�ܴ�/����������Ŀ¼
Const G_USER_DIR				= "User"

'-----�û��ļ�Ŀ¼�����治�ܴ�/����������Ŀ¼�����磺User��Test/User
Const G_USERFILES_DIR			= "UserFiles"

'-----�ϴ��ļ�Ŀ¼�����治�ܴ�/����������Ŀ¼
Const G_UP_FILES_DIR			= "Files"

'-----ģ���ļ�Ŀ¼,���治�ܴ�/,��������Ŀ¼
Const G_TEMPLETS_DIR			= "Templets"

'-----�ؼ���������Ŀ¼��
Const G_DBCONFIGFILE_DIR		= "DBConfigFiles" 

'-----��ϵͳ��Sql���ݿ⣻1:SQL,0:ACCESS
Const G_IS_SQL_DB				= "0"

'-----�����ݿ�,���治�ܴ�/,��������Ŀ¼
Const G_DATABASE_CONN_STR		= "Foosun_Data/FS500.mdb"

'-----��Աϵͳ��Sql���ݿ⣻1:SQL,0:ACCESS
Const G_IS_SQL_User_DB			= "0"

'-----��Ա���ݿ�,���治�ܴ�/,��������Ŀ¼
Const G_User_DATABASE_CONN_STR	= "Foosun_Data/FS_ME.mdb"

'-----�ɼ�ϵͳ��Sql���ݿ⣻1:SQL,0:ACCESS
Const G_IS_SQL_COLLECT_DB		= "0"

'-----�ɼ����ݿ�,���治�ܴ�/,��������Ŀ¼
Const G_COLLECT_DATA_STR		= "Foosun_Data/Collect.mdb"

'-----����һ�ζ�ȡ����������������ɼ���������Сһ��
Const CollectMaxOfOnePage		= "8"

'-----�ɼ������ļ�����λ��
Const G_SAVE_FILE_PATH			= "File"

'-----�ɼ������б���ҳ�������
Const G_NEWS_LIST_PAGES_NUMBER	= "300"

'-----�鵵���ݿ����ͣ�0ΪACCESS���ݿ⣬1ΪSQL���ݿ�
Const G_IS_SQL_Old_News_DB		= "0"

'-----�鵵���ݿ�,���治�ܴ�/,���ܴ�����Ŀ
Const G_Old_News_DATABASE_CONN_STR	= "Foosun_Data/OldNews.mdb"

'-----�鵵ģ���ַ,�޸�ģ�壬�����Ӧ�ı�ǩ�趨��
Const G_OLD_TEMPLET_PATH			= "/Templets/OldNews/Index.htm"

'-----IP��ַ��,��ע�ⲻҪ������Ŀ¼
Const G_IP_DATABASE_CONN_STR	= "Foosun_Data/AddressIp.mdb"

'-----��¼��־��������־��������־����������Ĭ��Ϊ7
Const G_HOLD_LOG_DAY_NUM		= "7"

'-----ǰ̨����ģʽ,0Ϊ��ͨģʽ��1Ϊ��ʱ������
Const G_SEARCH_TYPE				= "1"

'-----�������ݶ�����ͣһ��0Ϊ��ֹͣ
Const G_REFRESH_NUM_TIME		= "0"

'-----�����Զ���ҳ�ַ�����Ϊ0�򲻷�ҳ ��ҳ�ַ�һ��������2�� ��ҳ�ַ���������Html���
Const G_FS_Page_Txtlength		= "0"

'-----ϵͳ�汾��Ϣ FoosunCMS 5.0
Const G_COPYRIGHT				= "5.0"

'��ȫ����������ã���0,3,4,2,0����
'���ã�����ͨ���������ݿ��SQLע��õ��˹���Ա������������Բ��ܽ���ϵͳ
'��1λ	�Ƿ����ð�ȫ���� Ϊ0ʱ������ Ϊ1ʱ����
'��2λ	ȡ��֤���еĵڼ�λ�������㣬ȡ1-4֮�������
'��3λ	ȡ��֤���еĵڼ�λ�������㣬ȡ1-4֮�������
'��4λ	��ȡ�õ���λ��֤����ʲô���㣬1Ϊ�ӷ����㣻2Ϊ�˷�����
'��5λ	���õ��Ľ�����뵽����ĵڼ�λ����
'���簲ȫ���������Ϊ1��1��3��2��5  ��Ϊ���ð�ȫ�룬����֤��ĵ�һλ�͵���λ��˵Ľ�����뵽����ĵ���λ����
'������½ʱ ��������֤��Ϊ3568 ����Ա����ΪTryLogin
'����Ӧ�����������ΪTryLo18gin
'�����������֤�����������ĸ���벻Ҫʹ�ô˹���
Const G_SAFE_PASS_SET_STR		= "0,1,2,2,0"
'Session��ʱ,1Ϊ�������ڣ�0Ϊ10���ӹ���
Const G_SESSION_TIME_OUT		= "1"

'�ű���ʱ
Const G_SERVER_SCRIPT_TIME_OUT	= "6000"

'ϵͳ��½��ȫ�Ƿ���֤�����ݿ���,0����֤�����ݿ��У�1��֤�����ݿ���
Const G_SESSION_GETDATA			= "0"

'����ˢ��Ƶ��
'��һ��������ʾ������ˢ��һ�Σ��ڶ�����������ÿ��ˢ�¶��ٸ���¼
Const G_REFRESH_SPEED			= "1,5"

'���������������ʼ

'Adodb.Connection��
Const G_FS_CONN		= "Adodb.Connection"

'Scripting.FileSystemObject��
Const G_FS_FSO		= "Scripting.FileSystemObject"

'Adodb.RecordSet��
Const G_FS_RS		= "Adodb.RecordSet"

'Adodb.Stream��
Const G_FS_STREAM	= "Adodb.Stream"

'Scripting.Dictionary��
Const G_FS_DICT		= "Scripting.Dictionary"

'Microsoft.XMLHTTP��
Const G_FS_HTTP		= "Microsoft.XMLHTTP"

'MSXML2.XMLHTTP
Const G_FS_XMLHTTP	= "MSXML2.XMLHTTP"

'MSXML2.ServerXMLHTTP
Const G_MSXML2_SERVERXMLHTTP	= "MSXML2.ServerXMLHTTP"

'msxml2.FreeThreadedDOMDocument
Const G_MSXML2_DOCUMENT	= "msxml2.FreeThreadedDOMDocument"

'Excel.Application
Const G_EXCEL_APPLICATION	= "Excel.Application"

'CDONTS.NewMail
Const G_CDONTS_NEWMAIL	= "CDONTS.NewMail"

'CreatePreviewImage.cGvbox
Const G_CREATEPREVIEW_CGVBOX	= "CreatePreviewImage.cGvbox"

'SoftArtisans.ImageGen
Const G_SOFTARTISANS_IMAGEGEN	= "SoftArtisans.ImageGen"

'Persits.Jpeg
Const G_PERSITS_JPEG	= "Persits.Jpeg"

'wsImage.Resize
Const G_WSIMAGE_RESIZE	= "wsImage.Resize"

'MSWC.BrowserType
Const G_MSWC_BROWSERTYPE	= "MSWC.BrowserType"

'JMail.Message
Const G_JMAIL_MESSAGE		= "JMail.Message"

'WScript.Shell
Const G_WSCRIPT_SHELL		= "WScript.Shell"

'JRO.JetEngine
Const G_JRO_JETENGINE		= "JRO.JetEngine"

Const G_Badwords	= "�Ҳ�|xx,���|yy,fuck|ff,NND|nn"
'���ű������±�־.��ʽ������(1)���ǹر�(0)|��ʽ(1)����ͼƬ(0)|CSS��ʽ|�������ڵ�������ʾ��־��
'�����ʹ�õ�CSS��ʽ����Ҫ��CSS���a���ж��� �磺news a{color:Red;}
'�����ͼƬ������д������ͼƬ��ַ������·���������·�����磺http://www.foosun.net/1.gif��/1.gif
Const G_newNews		= "1|1|newNews|10"

'���뻭�л���ʾ�����������ұ߻�����ߡ�ֻ��Ҫ2��ֵ��right����left
Const G_CodeContentAlign = "left"

'���ŷ�ҳ���õ���ʽ����ѡ1��2��3��4��5,Ĭ��Ϊ4,��ҳ��ʽ�ο�������ǩ���ռ������б����ࡣ
Const G_NEWSPAGESTYLE = "5"

'�����ı�ǩ���ݿ����ӡ���������Ŀ¼
Const G_LABEL_DATA_STR		= "Foosun_Data/label.mdb"

'�����ղ�ͼƬ��ַ,ע��ͼƬ��ַǰ�����ܼ�"/"
Const AddFavorite = "sys_images/icon_star_2.gif"

'���ͺ���ͼƬ��ַ,ע��ͼƬ��ַǰ�����ܼ�"/"

Const SendFriend = "sys_images/sendmail.gif"

'++++++++++----------������Ϣ����----------++++++++++
%>