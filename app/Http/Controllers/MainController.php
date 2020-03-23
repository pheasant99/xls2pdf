<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use \App\Lib\excel2pdf;
use \PhpOffice\PhpSpreadsheet\IOFactory;

class MainController extends Controller
{
	/*
	 *	テスト１
	 */
	public function index()
	{
		$ex	=new excel2pdf();
	//	$fname	= '..\storage\app\public/001_納品書_タテ型.xlsx';
		$fname	= '..\storage\app\public/test.xlsx';
	//	$fname	= '..\storage\app\public/test__.xlsx';
	//	$fname	= '..\storage\app\public/test.xls';
/*		ブックを読み込んでセット	*  /
		$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
	//	$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xls();
		$book	= null;
		$book	= $reader->load($fname);
		$ex->setBook($book);
*/
		$ex->setExcelFilename($fname);
		$book	= $ex->getBook();
		
		//カラム幅をセット・・・取得出来ないと思ったら出来たので不要？
		$cw	= array();
/*		$cw[1]	= 118 * 1.1;
		$cw[2]	= 206 * 1.1;
		$cw[3]	=  73 * 1.1;
		$cw[4]	=  58 * 1.1;
		$cw[5]	=  18 * 1.1;
		$cw[6]	=  43 * 1.1;
		$cw[7]	=  28 * 1.1;
		$cw[8]	= 175 * 1.1;
*/
//		$ex->setColumnWidthsPx($cw);
//		$ex->setColumnWidthsmm($cw);
//		$ex->setColumnWidthsmm( 11.0, 20 );
		
		$sheet	=null;
		if($book != null) {
			// シートが1枚の場合
		//	$sheet = $book->getSheet(0);
			$sheet = $ex->getSheet();
			//セルに値をセットしてみる（自動計算されるか）
			$sheet->setCellValue('A18', 'まつもと');
			$sheet->setCellValue('D18', 15);		//数量
			$sheet->setCellValue('E18', 1200);		//単価
			//
			$sheet->setCellValue('A19', 'あいうえお');
			$sheet->setCellValue('D19', 4);			//数量
			$sheet->setCellValue('E19', 100);		//単価
		}
		
		$sheet	= $ex->getSheet();
		if($sheet != null) {
			$ex->setSheet($sheet);
			
			$ex->writePDF();		//PDF出力・・・この後 exit()?
			
		/*	セルの情報をダンプ	*/
		//	$ex->debugCell(0,0);		//カラム幅、行高さの出力
		//--------
	/*		$ex->debugExcelCell( 1, 1);
			$ex->debugExcelCell( 2, 1);
			$ex->debugExcelCell( 3, 1);
			$ex->debugExcelCell( 1, 2);
			$ex->debugExcelCell( 2, 2);
			$ex->debugExcelCell( 3, 2);
	*/
	/*	* /
			$ex->debugExcelCell( 1,30);
			$ex->debugExcelCell( 9,30);
			$ex->debugExcelCell( 1,34);
			$ex->debugExcelCell( 9,34);
	/*	*/	//
	//		$ex->debugExcelCell( 1, 2);
	//		$ex->debugExcelCell( 1, 3);
			$ex->debugCell( 2, 2);
			$ex->debugCell( 2, 3);
			$ex->debugCell( 2, 4);
			$ex->debugCell( 2, 5);
			$ex->debugCell( 2, 6);
			$ex->debugCell( 2, 7);
			$ex->debugCell( 2, 8);

			$ex->debugCell( 4, 2);
			$ex->debugCell( 4, 3);
			$ex->debugCell( 4, 4);
			$ex->debugCell( 4, 5);
			$ex->debugCell( 4, 6);
			$ex->debugCell( 4, 7);
			$ex->debugCell( 4, 8);
	//		$ex->debugCell( 1, 4);
		/*	デバッグ用	* /
			for($r=30;$r<35;$r++) {
				for($c=1;$c<10;$c++) {
					$ex->debugCell($c,$r);
				}
			}
		/*	*/
		}
	}
	
	//-----------------------------------
	//
	public function index2()
	{
		$ex	=new excel2pdf();
	//	$fname	= '..\storage\app\public/001_納品書_タテ型.xlsx';
		$fname	= '..\storage\app\public/test.xlsx';
/*		ブックを読み込んでセット	*  /
		$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
	//	$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xls();
		$book	= null;
		$book	= $reader->load($fname);
		$ex->setBook($book);
*/
		$ex->setExcelFilename($fname);
		$book	= $ex->getBook();
		
		//カラム幅をセット・・・取得出来ないと思ったら出来たので不要？
		$cw	= array();
/*		$cw[1]	= 118 * 1.1;
		$cw[2]	= 206 * 1.1;
		$cw[3]	=  73 * 1.1;
		$cw[4]	=  58 * 1.1;
		$cw[5]	=  18 * 1.1;
		$cw[6]	=  43 * 1.1;
		$cw[7]	=  28 * 1.1;
		$cw[8]	= 175 * 1.1;
*/
		for($i=1;$i<20;$i++) {
//			$cw[$i]	= 175;				//px
//			$cw[$i]	= 19.5;//25.4;		//mm
			$cw[$i]	= 11.0;				//mm
		}
//		$ex->setColumnWidthsPx($cw);
//		$ex->setColumnWidthsmm($cw);
//		$ex->setColumnWidthsmm( 11.0, 20 );
		
		$sheet	=null;
		if($book != null) {
			// シートが1枚の場合
		//	$sheet = $book->getSheet(0);
			$sheet = $ex->getSheet();
			//セルに値をセットしてみる（自動計算されるか）
			$sheet->setCellValue('A17', 'まつもと');
			$sheet->setCellValue('D17', 15);		//数量
			$sheet->setCellValue('E17', 1200);		//単価
			//
			$sheet->setCellValue('A18', 'あいうえお');
			$sheet->setCellValue('D18', 4);			//数量
			$sheet->setCellValue('E18', 100);		//単価
		}
		
	//	$sheet	= $ex->getSheet();
		if($sheet != null) {
			$ex->setSheet($sheet);
			
			$ex->writePDF();		//PDF出力・・・この後 exit()?
			
		/*	セルの情報をダンプ	*  /
			$ex->debugCell(0,0);		//カラム幅、行高さの出力
		/*	デバッグ用	*  /
			for($r=1;$r<27;$r++) {
				for($c=1;$c<18;$c++) {
					$ex->debugCell($c,$r);
				}
			}
		/*	*/
		}
	}
	
	//-----------------------------------
	//
	public function index3()
	{
		
		$ex	=new excel2pdf();
		
//		$ret	= $ex->area2index('A1:d5');
//		var_dump($ret);
//		echo "<br><br><br><br>";

//		try {
			$fname	= '..\storage\app\public/001_納品書_タテ型.xlsx';
		//	$fname	= '..\storage\app\public/test.xlsx';
			$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
		//	$book	= $reader->load($fname);
			$book	= $reader->load($fname);
			
			if ($book != null) {
				// シートが1枚の場合
				$sheet = $book->getSheet(0);
				//おまじない？
				$sheet = $sheet->calculateColumnWidths();
				$sheet = $sheet->refreshColumnDimensions();
				
//				$h		= $sheet->getRowDimension(3)->getRowHeight();
//				$w		= $sheet->getColumnDimension('A')->getWidth();

				$rdim	= $sheet->getColumnDimensions();
				$c	= count($rdim);
				echo "カラム 数: ${c}<br><br>";
				foreach ($rdim as $ky => $ob) {
					$v	= $ob->getWidth();
					echo "(clm=${ky}):${v}  ";
				}
			//	for($r=1;$r<$c;$r++) {
			//		if(isset($rdim[$r])) {
			//			$v	= $rdim[$r]->getWidth();
			//			echo "(clm=${r}):${v}  ";
			//		}
			//	}
				echo "<br><br>";
				
				$rdim	= $sheet->getRowDimensions();
				$c	= count($rdim);
				echo "行 : ${c}<br><br>";
				for($r=1;$r<$c;$r++) {
					if(isset($rdim[$r])) {
						$v	= $rdim[$r]->getRowHeight();
						echo "(${r}):${v} ";
					}
				}
				echo "<br><br>";
				
				
/*	*/			for($r=1;$r<36;$r++) {
					for($c=1;$c<10;$c++) {
						$cell	= $sheet->getCellByColumnAndRow($c,$r,false);
						if($cell!=null) {
						//	$val	= $cell->getValue();
							$val	= $cell->getFormattedValue();
							
							$mg		= $cell->isMergeRangeValueCell();	//代表セル
							$marge	= $cell->getMergeRange();
							$style	= $cell->getStyle();
							$brders	= $style->getBorders();
							$bdr	= $brders->getBottom();
							$bline	= $bdr->getBorderStyle();
							
						//	echo $val."(${bline})(${mg} ${marge})\t";
							echo $val."(${mg} ${marge})\t";
						}
						else {
							echo "(cell:null)\t";
						}
					}
					echo "<br>";
				}
/*	*/			
				$area	= $sheet->getPageSetup()->getPrintArea();
				$scale	= $sheet->getPageSetup()->getScale();
				echo '印刷範囲 '.$area.'<br>';
				echo 'スケール '.$scale.'<br>';
				// tinkerによるデバッグ
			//	eval(\Psy\sh());
				exit();
			}
			else {
				throw new \Exception('error.');
			}
//		}
//		catch (\Exception $e) {
//			Log::error($e->getMessage());
//		}
		// 
		// ビュー
		//return view('welcome');
	}
	
	//-----------------------------------
	//		エクセルダウンロード
	public function indexDL()
	{
		$fname		= '..\storage\app\public/001_納品書_タテ型.xlsx';
		$filename	= 'template.xlsx';

//		$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
//		$book	= $reader->load($fname);

		$ex	=new excel2pdf();
		$ex->setExcelFilename($fname);
		$book	= $ex->getBook();

		if($book != null) {
			// シートが1枚の場合
			$sheet = $book->getSheet(0);
			//セルに値をセットしてみる（自動計算されるか）
			$sheet->setCellValue('B18', 'まつもと');
			$sheet->setCellValue('J18', 15);		//数量
			$sheet->setCellValue('L18', 1200);		//単価
			//
			$sheet->setCellValue('B19', 'あいうえお');
			$sheet->setCellValue('J19', 4);			//数量
			$sheet->setCellValue('L19', 100);		//単価
		}

		$writer = IOFactory::createWriter($book, 'Xlsx');

		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="'.$filename.'"');
		header('Cache-Control: max-age=0');
		header('Cache-Control: max-age=1');
		header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
		header('Cache-Control: cache, must-revalidate');
		header('Pragma: public');

		$writer->save('php://output');
		exit();
	}

	/**
	 *
	 */
	public function indexTest()
	{
echo "<br>BORDER_NONE           ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE;
echo "<br>BORDER_DASHDOT        ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOT;
echo "<br>BORDER_DASHDOTDOT     ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOTDOT;
echo "<br>BORDER_DASHED         ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHED;
echo "<br>BORDER_DOTTED         ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOTTED;
echo "<br>BORDER_DOUBLE         ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE;
echo "<br>BORDER_HAIR           ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR;
echo "<br>BORDER_MEDIUM         ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM;
echo "<br>BORDER_MEDIUMDASHDOT     ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOT;
echo "<br>BORDER_MEDIUMDASHDOTDOT  ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOTDOT;
echo "<br>BORDER_MEDIUMDASHED      ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHED;
echo "<br>BORDER_SLANTDASHDOT      ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_SLANTDASHDOT;
echo "<br>BORDER_THICK             ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK;
echo "<br>BORDER_THIN              ". \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
echo "<br><br><br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_GENERAL 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER_CONTINUOUS 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_JUSTIFY 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_FILL 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_DISTRIBUTED 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_JUSTIFY 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_DISTRIBUTED 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::READORDER_CONTEXT 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::READORDER_LTR 	 	."<br>";
echo \PhpOffice\PhpSpreadsheet\Style\Alignment::READORDER_RTL 	."<br>";
echo "<br><br><br>";

echo "--------------------<br>";

	}

	/**
	 *	オブジェクトの取得
	 */
	public	function indexDOBJ()
	{
		$fname	= '..\storage\app\public/001_納品書_タテ型.xlsx';
	//	$fname	= '..\storage\app\public/test.xlsx';

		$ex	=new excel2pdf();
		$ex->setExcelFilename($fname);
		$sheet	= $ex->getSheet();

		$garr	= $sheet->getDrawingCollection();
		$c	= count($garr);
		
		echo("count=${c}<br>");
		
		foreach($garr as $obj){
			$nm	= $obj->getName();
			$cn	= $obj->getCoordinates();
			$x	= $obj->getOffsetX();
			$y	= $obj->getOffsetY();
			$str	= "<br> position ${x} ${y} name:${nm} coordinates:${cn}";
			echo($str);
		}
		exit("<br>end<br");
	}
	
}
