<? require("connection.php"); ?>

<? require("ZGJX_UpLoad_Excel.php"); ?>

<?

if($leadExcel == "true")
{

//获取上传的文件名
$filename = $HTTP_POST_FILES['inputExcel']['name'];

//上传到服务器上的临时文件名
$tmp_name = $_FILES['inputExcel']['tmp_name'];
$msg = uploadFile($filename,$tmp_name);
}

?>

<form name="form2" method="post" enctype="multipart/form-data">
<input type="hidden" name="leadExcel" value="true">
<table align="center" width="90%" border="0">
<tr>
<td>
<input type="file" name="inputExcel"><input type="submit" value="导入数据">
</td>
</tr>
</table>
</form>