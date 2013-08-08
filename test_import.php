<html>
</head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
</head>

</html>
<?php
	ini_set('memory_limit','64M');
	require_once 'connect.php';
	require_once 'PHPExcel.php';
	require_once 'PHPExcel\IOFactory.php';
	require_once 'PHPExcel\Reader\Excel5.php';
	//设置缓存模式，处理大数据的excel表
//	$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_in_memory_gzip; 
//	$cacheSettings = array(); 
//	PHPExcel_Settings::setCacheStorageMethod($cacheMethod,$cacheSettings); 
	
	//$objReader=PHPExcel_IOFactory::createReader('Excel5');
	$filePath="test1.xls";
	$PHPExcel = new PHPExcel();
	/**默认用excel2007读取excel，若格式不对，则用之前的版本进行读取*/
	$PHPReader = new PHPExcel_Reader_Excel2007();
	if(!$PHPReader->canRead($filePath)){
	   $PHPReader = new PHPExcel_Reader_Excel5();
	   if(!$PHPReader->canRead($filePath)){
		 echo 'no Excel';
		 return ;
	   }
	}
	$objPHPExcel=$PHPReader->load("test1.xls");
	$sheet=$objPHPExcel->getSheet(0);
	
	$highestRow=$sheet->getHighestRow();
	$highestColumn=$sheet->getHighestColumn();
	echo $highestRow."<br>";
	echo $highestColumn.'<br>';
	 for($j = 3; $j <= $highestRow; $j ++){  
	    for($k = 'A'; $k <= $highestColumn; $k ++) {  
	       $array[$j][$k] = (string)$objPHPExcel->getActiveSheet ()->getCell ( "$k$j" )->getValue ();  
		
	    } 
	
	echo $sql="INSERT INTO `test`(`id`, `names`, `销量`) VALUES ('{$array[$j]['A']}','{$array[$j]['B']}','{$array[$j]['C']}')";
	mysql_query($sql);
	mysql_query("set names 'gbk'");
	} 
	 //for($k = 'A'; $k <= $highestColumn; $k ++) { 
	//	  for($j = 3; $j <= $highestRow; $j ++){ 
	//		  $array[$k][$j] = (string)$objPHPExcel->getActiveSheet ()->getCell ( "$k$j" )->getValue();  
	//	  }
	// }
	 //for
?>