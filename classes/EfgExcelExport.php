<?php if (!defined('TL_ROOT')) die('You can not access this file directly!');

/**
 * Contao Open Source CMS
 *
 * Copyright (c) 2005-2015 Leo Feyer
 *
 * @license LGPL-3.0+
 */


/**
 * Class EfgExcelExport
 *
 * Excel 2003/2007 export for EFG using PHPExcel library
 * @copyright  Helmut Schottmüller 2008-2015
 * @author     Helmut Schottmüller <hschottm@gmail.com>
 * @package    Controller
 */
class EfgExcelExport extends Backend
{
	/**
	* Calculate the Excel cell address (A,...,Z,AA,AB,...) from a numeric index
	*
	*/
	private function getCellTitle($index)
	{
		$alphabet = array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z");
		if ($index < 26) return $alphabet[$index];
		return $alphabet[floor($index / 26)-1] . $alphabet[$index-(floor($index / 26)*26)];
	}
	
	/**
	 * Export the form data to Microsoft Excel
	 *
	 */
	public function export($headers, $data)
	{
		if (!is_array($data) || count($data) == 0) return;
                $formtitle_natural = $data[1]["form"];

		// Create new PHPExcel object
		$objPHPExcel = new \PHPExcel();
		// Set properties
		$objPHPExcel->getProperties()->setCreator("Contao");
		$objPHPExcel->getProperties()->setLastModifiedBy("Contao");
		$objPHPExcel->getProperties()->setTitle($formtitle_natural);
		$objPHPExcel->getProperties()->setSubject($formtitle_natural);
		$objPHPExcel->getProperties()->setDescription($formtitle_natural);
		$objPHPExcel->getProperties()->setKeywords("office 2007 Contao");
		$objPHPExcel->getProperties()->setCategory("form input data");
		
		$objPHPExcel->setActiveSheetIndex(0);
		$intRowCounter = 0;
		$intColCounter = 0;
		// List records
		foreach ($headers as $header)
		{
			$cell = $this->getCellTitle($intColCounter) . ($intRowCounter+1);
			$intColCounter++;
			$objPHPExcel->getActiveSheet()->setCellValue((string)$cell, utf8_encode($header));
		}

		foreach ($data as $row)
		{
			$intRowCounter++;
			$intColCounter = 0;
			foreach ($row as $value)
			{
				$objPHPExcel->getActiveSheet()->setCellValue($this->getCellTitle($intColCounter) . ($intRowCounter+1), utf8_encode($value));
				$intColCounter++;
			}
			// autosize columns
			for ($i = 0; $i < count($headers); $i++)
			{
				$objPHPExcel->getActiveSheet()->getColumnDimension($this->getCellTitle($i))->setAutoSize(true);
			}
		}

		$objPHPExcel->getActiveSheet()->duplicateStyleArray(
				array(
					'font'    => array(
						'bold'      => true,
						'color'     => array(
							'argb' => 'FF000080'
						)
					),
					'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
					),
					'borders' => array(
						'top'     => array(
		 					'style' => PHPExcel_Style_Border::BORDER_THIN
		 				),
						'bottom'     => array(
		 					'style' => PHPExcel_Style_Border::BORDER_THIN
		 				)
		 			),
				),
				'A1:' . $this->getCellTitle(count($headers)-1) . '1'
		);
		
		// Rename sheet
		$filename = escapeshellcmd(str_replace(" ", "_", $formtitle_natural));
		$formtitle_natural== utf8_substr($formtitle_natural, 0, 31);
		$objPHPExcel->getActiveSheet()->setTitle($formtitle_natural);
		// Set active sheet index to the first sheet, so Excel opens this as the first sheet
		$objPHPExcel->setActiveSheetIndex(0);
				
		// Save Excel 2007 file
		$objWriter = new \PHPExcel_Writer_Excel2007($objPHPExcel);
		$objWriter->save(TL_ROOT . "/system/tmp/export_" . $this->strFormKey . "_" . date("Ymd") .".xlsx");
		header('Content-Type: appplication/excel');
		header('Content-Transfer-Encoding: binary');
		header('Content-Disposition: attachment; filename="export_' . $filename . '_' . date("Ymd") .'.xlsx"');
		header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
		header('Pragma: public');
		header('Expires: 0');
		$f = new \File("system/tmp/export_" . $this->strFormKey . "_" . date("Ymd") .".xlsx");
		print $f->getContent();
		$f->delete();
		exit;
	}
}

?>
