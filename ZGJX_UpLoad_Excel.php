<? require("connection.php"); ?>
<?
//����Excel�ļ�
function uploadFile($file,$filetempname)
{

//�Լ����õ��ϴ��ļ����·��
$filePath = 'upFile/';
$str = "";

//�����·��������PHPExcel��·�����޸�
set_include_path('.'. PATH_SEPARATOR .
'D:\xampp\htdocs\XM1\PHPExcel' . PATH_SEPARATOR .
get_include_path());

require_once 'PHPExcel.php';
require_once 'PHPExcel\IOFactory.php';
require_once 'PHPExcel\Reader\Excel5.php';

$filename=explode(".",$file);//���ϴ����ļ����ԡ�.����Ϊ׼��һ�����顣
$time=date("y-m-d-H-i-s");//ȥ��ǰ�ϴ���ʱ��
$filename[0]=$time;//ȡ�ļ���t�滻
$name=implode(".",$filename); //�ϴ�����ļ���
$uploadfile=$filePath.$name;//�ϴ�����ļ�����ַ


//move_uploaded_file() �������ϴ����ļ��ƶ�����λ�á����ɹ����򷵻� true�����򷵻� false��
$result=move_uploaded_file($filetempname,$uploadfile);//�����ϴ�����ǰĿ¼��
if($result) //����ϴ��ļ��ɹ�����ִ�е���excel����
{
$objReader = PHPExcel_IOFactory::createReader('Excel5');//use excel2007 for 2007 format
$objPHPExcel = $objReader->load($uploadfile);
$sheet = $objPHPExcel->getSheet(0);
$highestRow = $sheet->getHighestRow(); // ȡ��������
$highestColumn = $sheet->getHighestColumn(); // ȡ��������

//ѭ����ȡexcel�ļ�,��ȡһ��,����һ��
for($j=2;$j<=$highestRow;$j++)
{
for($k='A';$k<=$highestColumn;$k++)
{
$str .= iconv('utf-8','gbk',$objPHPExcel->getActiveSheet()->getCell("$k$j")->getValue()).'\\';//��ȡ��Ԫ��
}
//explode:�������ַ����ָ�Ϊ���顣
$strs = explode("\\",$str);
$sql = "INSERT INTO ������ (id,����,����) VALUES('".
$strs[0]."','". //��������
$strs[1]."','". //������
$strs[2]."','". //����Ŀ��
')"; //����Ŀ���

mysql_query("set names GBK");//�����ָ�����ݿ��ַ�����һ������������ݿ�����ϵ��
if(!mysql_query($sql))
{
return false;
}
$str = "";
}

unlink($uploadfile); //ɾ���ϴ���excel�ļ�
$msg = "����ɹ���";
}
else
{
$msg = "����ʧ�ܣ�";
}

return $msg;
}
?>