<?
	/*********************************************************************************************\
	***********************************************************************************************
	**                                                                                           **
	**  eSchool                                                                                  **
	**  Version 2.0                                                                              **
	**                                                                                           **
	**  http://www.eschool.com.pk                                                                **
	**                                                                                           **
	**  Copyright 2005-14 (C) SW3 Solutions                                                      **
	**  http://www.sw3solutions.com                                                              **
	**                                                                                           **
	**  ***************************************************************************************  **
	**                                                                                           **
	**  Project Manager:                                                                         **
	**                                                                                           **
	**      Name  :  Muhammad Tahir Shahzad                                                      **
	**      Email :  mtshahzad@sw3solutions.com                                                  **
	**      Phone :  +92 333 456 0482                                                            **
	**      URL   :  http://www.mtshahzad.com                                                    **
	**                                                                                           **
	***********************************************************************************************
	\*********************************************************************************************/

	$objPhpExcel = new PHPExcel( );

	$objPhpExcel->getProperties()->setCreator($sSchool)
								 ->setLastModifiedBy($_SESSION["AdminName"])
								 ->setTitle("Income and Expenditure Report")
								 ->setSubject("Reports")
								 ->setDescription("Income and Expenditure Report")
								 ->setKeywords("")
								 ->setCategory("Reports");

	$objPhpExcel->setActiveSheetIndex(0);


	$objPhpExcel->getActiveSheet()->setCellValue("A1", "Educational Development Network");
	$objPhpExcel->getActiveSheet()->getStyle("A1")->getFont()->setSize(24);
	$objPhpExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true);
	$objPhpExcel->getActiveSheet()->mergeCells("A1:C1");
	$objPhpExcel->getActiveSheet()->getStyle("A1")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

	$objPhpExcel->getActiveSheet()->setCellValue("A2", ((count($iSchools) == 1) ? $sSchool : $_SESSION['SiteTitle']));
	$objPhpExcel->getActiveSheet()->getStyle("A2")->getFont()->setSize(20);
	$objPhpExcel->getActiveSheet()->getStyle("A2")->getFont()->setBold(true);
	$objPhpExcel->getActiveSheet()->mergeCells("A2:C2");
	$objPhpExcel->getActiveSheet()->getStyle("A2")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
	
	$objPhpExcel->getActiveSheet()->setCellValue("A3", "Income and Expenditure Report");
	$objPhpExcel->getActiveSheet()->getStyle("A3")->getFont()->setSize(16);
	$objPhpExcel->getActiveSheet()->getStyle("A3")->getFont()->setBold(true);
	$objPhpExcel->getActiveSheet()->mergeCells("A3:C3");
	$objPhpExcel->getActiveSheet()->getStyle("A3")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

	$sCurrentYear 		= date('Y', strtotime($sStartDate)); 
	$sCurrentMonth 		= date('m', strtotime($sStartDate)); 
	$iEndMonth 		    = date('m', strtotime($sEndDate)); 
	$sPreviousMonth 	= date('m', strtotime('-1 months', strtotime($sStartDate))); 
	
	if($sCurrentMonth < $sPreviousMonth)
	{
		$sCurrentYear   = $sCurrentYear;
		$sPreviousYear  = $sCurrentYear - 1;
	}
	else
	{
		$sCurrentYear   = $sCurrentYear;
		$sPreviousYear  = $sCurrentYear;
	}
	
	$sArearYear	= (($sCurrentMonth < $sArrearMonth)? $sCurrentYear - 1 : $sCurrentYear);
		
	$iMonthDays   		= date("t", strtotime("{$sPreviousYear}-{$sPreviousMonth}-01"));
	$sPreviousMonthYear = date('Y-m', strtotime('-1 months', strtotime($sEndDate))); 
	$sPreviousStartDate = date('Y-m', strtotime('-1 months', strtotime($sStartDate)))."-01"; 
	$sPreviousEndDate 	= date('Y-m', strtotime('-1 months', strtotime($sStartDate)))."-{$iMonthDays}"; 

	$sCurrentMonth  = $sMonths[str_replace("0", "", $sCurrentMonth)];
	$sPreviousMonth = $sMonths[str_replace("0", "", $sPreviousMonth)];
	$sEndMonth      = $sMonths[str_replace("0", "", $iEndMonth)];

	
	$iRow = 4;

	if ($sStartDate != "" && $sEndDate != "")
	{
		$objPhpExcel->getActiveSheet()->setCellValue("A{$iRow}", (date("M Y", strtotime($sStartDate))." to ".date("M Y", strtotime($sEndDate))));
		$objPhpExcel->getActiveSheet()->getStyle("A{$iRow}")->getFont()->setSize(12);
		$objPhpExcel->getActiveSheet()->getStyle("A{$iRow}")->getFont()->setBold(true);
		$objPhpExcel->getActiveSheet()->mergeCells("A{$iRow}:J{$iRow}");
		$objPhpExcel->getActiveSheet()->getStyle("A{$iRow}")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

		$iRow ++;
	}

	$objPhpExcel->getActiveSheet()->setCellValue("A{$iRow}", ("As on ".date($_SESSION["DateFormat"])));
	$objPhpExcel->getActiveSheet()->getStyle("A{$iRow}")->getFont()->setSize(11);
	$objPhpExcel->getActiveSheet()->mergeCells("A{$iRow}:J{$iRow}");
	$objPhpExcel->getActiveSheet()->getStyle("A{$iRow}")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

 

	$sHeadingStyle = array('font' => array('bold' => true, 'size' => 11),
						   'alignment' => array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT),
						   'borders' => array('top' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
											  'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
											  'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
											  'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN)),
						   'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'DDDDDD')) );


	$sBorderStyle = array('alignment' => array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT),
						  'borders' => array('top' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
											 'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
											 'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
											 'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN)));



	$sBlockStyle = array('borders' => array('top' => array('style' => PHPExcel_Style_Border::BORDER_THICK),
											'right' => array('style' => PHPExcel_Style_Border::BORDER_THICK),
											'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THICK),
											'left' => array('style' => PHPExcel_Style_Border::BORDER_THICK)));
	
	$sCellStyle = array('font' => array('underline' => true));
										
	$iRow += 2;

	$objPhpExcel->getActiveSheet()->setCellValue("A{$iRow}", "Transaction Description");
	$objPhpExcel->getActiveSheet()->setCellValue("B{$iRow}", "Amount");
	$objPhpExcel->getActiveSheet()->setCellValue("C{$iRow}", "Amount");


	for ($i = 0; $i < 3; $i ++)
		$objPhpExcel->getActiveSheet()->duplicateStyleArray($sHeadingStyle, ((getExcelCol($i))."{$iRow}:".(getExcelCol($i)).$iRow));


	$iRow ++;

	$sConditions = "school_id='$iSchool'";
	//$sConditions = "school_id='$iSchools[0]'";
	
	if ($sStartDate != "" && $sEndDate != "")
		$sConditions .= " AND (date BETWEEN '$sStartDate' AND '$sEndDate')";

	$iSchool = $iSchools[0];
	
	$iGrandPaybale = 0;
	$iGrandIncome  = 0;
	$iGrandTotal   = 0;
	$iGrandFee     = 0;
	
	$iBankBalance 			= getDbValue("SUM(amount)", "tbl_accounts_payments", "$sConditions AND type='R' ");
	$iPreviousArrears 		= getDbValue("SUM(amount)", "tbl_student_fee", "school_id='$iSchool' AND `due_date` BETWEEN  '$sPreviousStartDate' AND '$sPreviousEndDate' AND paid='N'");
	$iPreviousFee 			= getDbValue("SUM(sfd.amount-sfd.concession)", "tbl_student_fee sf, tbl_student_fee_details sfd", "sf.student_session_id IN (SELECT id FROM tbl_student_sessions WHERE school_id='$iSchool') AND sf.paid='Y' AND (sf.paid_date BETWEEN '$sPreviousStartDate' AND '$sPreviousEndDate') AND sf.id=sfd.student_fee_id AND fee_category='Tuition Fee'");
	$iCurrentFee 			= getDbValue("SUM(sfd.amount-sfd.concession)", "tbl_student_fee sf, tbl_student_fee_details sfd", "sf.student_session_id IN (SELECT id FROM tbl_student_sessions WHERE school_id='$iSchool') AND sf.paid='Y' AND (sf.paid_date BETWEEN '$sStartDate' AND '$sEndDate') AND sf.id=sfd.student_fee_id AND fee_category='Tuition Fee'");
	$iStationaryReceivable 	= getDbValue("SUM(sfd.amount)", "tbl_student_fee sf, tbl_student_fee_details sfd", "sf.student_session_id IN (SELECT id FROM tbl_student_sessions WHERE school_id='$iSchool') AND sf.paid='Y' AND (sf.paid_date BETWEEN '$sStartDate' AND '$sEndDate') AND sf.id=sfd.student_fee_id AND fee_category='Stationery Charges'");
	$iCashAvailable 		= getDbValue("SUM(amount)", "tbl_accounts_payments", "$sConditions AND type='R'");

	
	
	$iColumn = 0;

	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn, $iRow, "Balance at Bank");
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn+2, $iRow++, formatNumber($iBankBalance, false));
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn, $iRow, "Previous Arrears till {$sPreviousMonth} {$sPreviousYear}");
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn+2, $iRow++, formatNumber($iPreviousArrears, false));
	
	while (strtotime($sStartDate) <= strtotime($sEndDate))
	{
        $sMonthYear    = date("Y-m", strtotime($sStartDate));
        $sFeeStartDate = date("Y-m", strtotime("{$sStartDate}"))."-01";
		$iMonthDays    = date("t", strtotime("{$sFeeStartDate}-01"));
		$sFeeEndDate   = date("Y-m", strtotime("{$sStartDate}"))."-$iMonthDays";
		
		$iFee = getDbValue("SUM(sfd.amount-sfd.concession)", "tbl_student_fee sf, tbl_student_fee_details sfd", "sf.student_session_id IN (SELECT id FROM tbl_student_sessions WHERE school_id='$iSchool') AND sf.paid='Y' AND (sf.paid_date BETWEEN '$sFeeStartDate' AND '$sFeeEndDate') AND sf.id=sfd.student_fee_id AND fee_category='Tuition Fee'");
		
		$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn, $iRow, "Tution Fee ".$sMonths[str_replace("0", "", date('m', strtotime($sMonthYear)))]." ".date('Y', strtotime($sMonthYear)));
		$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn+2, $iRow++, formatNumber($iFee, false));
		
		$iGrandFee += $iFee;
		
        $sStartDate = date ("Y-m-d", strtotime("+1 month", strtotime($sStartDate)));
    }
	
	$iGrandIncome = $iBankBalance + $iPreviousArrears + $iGrandFee + $iStationaryReceivable + $iCashAvailable;
	
	
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn, $iRow, "Stationary Charges Receivable");
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn+2, $iRow++, formatNumber($iStationaryReceivable, false));
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn, $iRow, "Cash Available for use");
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn+2, $iRow++, formatNumber($iGrandIncome, false));
	
	
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn, $iRow, "Current Payables");
	$objPhpExcel->getActiveSheet()->getStyle(getExcelCol($iColumn).$iRow.":".getExcelCol($iColumn).$iRow)->applyFromArray($sCellStyle);
	
	$iRow++;
	
	for ($i = 8; $i <= $iRow; $i ++)
	{
		for ($j = 0; $j < 3; $j ++)
			$objPhpExcel->getActiveSheet()->duplicateStyleArray($sBorderStyle, (getExcelCol($j).$i.":".getExcelCol($j).$i));
	}
	
	$iRowExpense = $iRow;
	
	$iSalary = getDbValue("SUM(payable_salary)", "tbl_payroll", "school_id='$iSchool' AND DATE_FORMAT(from_date, '%Y-%m')='$sPreviousMonthYear'");
	
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn, $iRow, "Salary for the m/o {$sCurrentMonth} - {$sEndMonth} {$sCurrentYear}");
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow(($iColumn+1), $iRow++, formatNumber($iSalary, false));
	
	$iGrandPaybale += $iSalary;
	
	$sCurrentPayablesList = getList("tbl_current_payables cp, tbl_current_payables_details cpd", "cpd.id", "CONCAT(transaction_detail, ',', amount)", "cp.school_id='$iSchool' AND DATE_FORMAT(`date`, '%Y-%m')='".date("Y-m", strtotime($sEndDate))."' AND cp.id=cpd.payable_id AND cpd.approved='Y'");

	
	$iRowExpense = $iRow;
	
	foreach($sCurrentPayablesList as $iCurrentPayable => $sCurrentPayable)
	{
		$iColumn = 0;
		
		@list($sTransactionDetail, $iAmount) = @explode(",", $sCurrentPayable);
		
		
		$iGrandPaybale += $iAmount;
		
		$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn++, $iRowExpense, $sTransactionDetail);
		$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn, $iRowExpense, formatNumber($iAmount, false));
		
		for ($i = 0; $i < 3; $i ++)
			$objPhpExcel->getActiveSheet()->duplicateStyleArray($sBorderStyle, (getExcelCol($i).$iRowExpense.":".getExcelCol($i).$iRowExpense));
		
		$iRowExpense++;
	}
	
	$iRowExpense--; 
	
	$objPhpExcel->getActiveSheet()->duplicateStyleArray($sBlockStyle,(getExcelCol($iColumn).($iRow-1).":".getExcelCol($iColumn).($iRowExpense)));
	
	$iRow = $iRowExpense + 1;
	
	$iColumn = 0;
	
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn, $iRow, "Total Current Payables");
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn+1, $iRow, "(".formatNumber($iGrandPaybale, false). ")");
	$objPhpExcel->getActiveSheet()->getStyle(getExcelCol($iColumn).$iRow.":".getExcelCol($iColumn).$iRow)->applyFromArray($sCellStyle);
	
	for ($i = 0; $i < 3; $i ++)
		$objPhpExcel->getActiveSheet()->duplicateStyleArray($sBorderStyle, (getExcelCol($i).$iRow.":".getExcelCol($i).$iRow));
	
	$objPhpExcel->getActiveSheet()->duplicateStyleArray($sHeadingStyle, ((getExcelCol(1))."{$iRow}:".(getExcelCol(1)).$iRow));
	
	$iRow ++;
	
	$iGrandTotal = $iGrandIncome - $iGrandPaybale;
	
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn, $iRow, "Surplus/Defict");
	$objPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($iColumn+2, $iRow, formatNumber($iGrandTotal, false));
	
	for ($i = 0; $i < 3; $i ++)
		$objPhpExcel->getActiveSheet()->duplicateStyleArray($sBorderStyle, (getExcelCol($i).$iRow.":".getExcelCol($i).$iRow));
		
		
	$objPhpExcel->getActiveSheet()->getColumnDimension("A")->setWidth(70);
	$objPhpExcel->getActiveSheet()->getColumnDimension("B")->setWidth(20);
	$objPhpExcel->getActiveSheet()->getColumnDimension("C")->setWidth(20);


	$objPhpExcel->getActiveSheet()->getHeaderFooter()->setOddHeader('');
	$objPhpExcel->getActiveSheet()->getHeaderFooter()->setOddFooter("&L&B Income and Expenditure Report &R Generated on ".date("d-M-Y"));

	$objPhpExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
	$objPhpExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);

	$objPhpExcel->getActiveSheet()->getPageMargins()->setTop(0.4);
	$objPhpExcel->getActiveSheet()->getPageMargins()->setRight(0.2);
	$objPhpExcel->getActiveSheet()->getPageMargins()->setLeft(0.4);
	$objPhpExcel->getActiveSheet()->getPageMargins()->setBottom(0);

	$objPhpExcel->getActiveSheet()->getPageSetup()->setFitToWidth(1);

	$objPhpExcel->getActiveSheet()->setTitle("Income and Expenditure Report");



	$sExcelFile = "Income and Expenditure Report-".((count($iSchools) == 1) ? $sSchool : $_SESSION['SiteTitle'])."-".date($_SESSION["DateFormat"]).".xlsx";

	header("Content-Type: application/vnd.ms-excel");
	header("Content-Disposition: attachment;filename=\"{$sExcelFile}\"");
	header("Cache-Control: max-age=0");

	$objWriter = PHPExcel_IOFactory::createWriter($objPhpExcel, 'Excel2007');
	$objWriter->save("php://output");
?>