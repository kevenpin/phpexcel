<?php
	require_once('connect.php');
	require_once 'PHPExcel.php';
	require_once 'PHPExcel\IOFactory.php';
	require_once 'PHPExcel\Reader\Excel5.php';
	$objPHPExcel=new PHPExcel();
	$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue('A1','序号')
		->setCellValue('B1','产品名')
		->setCellValue('C1','销量');
	
	$sql = "SELECT * FROM `test`";
	$query=mysql_query($sql);
	$i=1;
	while($rs=mysql_fetch_array($query)){
		$i++;
		 $objPHPExcel->setActiveSheetIndex(0)
			 ->setCellValue('A'.$i,$rs['id'])
			 ->setCellValue('B'.$i,$rs['names'])
			 ->setCellValue('c'.$i,$rs[2]);
		 
		
	}
	$objPHPExcel->getActiveSheet()->setTitle("销量表");
	$objPHPExcel->setActiveSheetIndex(0);
	$objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');
	$objWriter->save("result.xls");
	?>