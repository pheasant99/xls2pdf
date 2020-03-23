<?php
/*-----------------------------------------------------------------------------*
	
 *-----------------------------------------------------------------------------*/
namespace App\Lib;

use \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;

//=======================================
/**
 *	セル情報のクラス
 */
class	ExcelCell
{
	public	$posx;
	public	$posy;
	public	$width;
	public	$height;
	
	public	$strVal;
	public	$bgcolor;			//背景色
	public	$bgcolor2;			//背景色
	public	$filltype;			//bool ::FILL_SOLIDの時true
	
	public	$Font;				//フォント名
	public	$FontSize;			//フォントサイズ(mm)
	public	$bold;				//Bool	ボールド
	public	$italic;			//Bool	イタリック
	public	$superscript;		//Bool	上付き
	public	$subscript;			//Bool	添え字
	public	$underline;			//Str	下線
 	public	$strikethrough;		//Bool	取消線
	public	$color;				//Str
	
	public	$HAlignment;		//str	
	public	$VAlignment;		//str	
	public	$wrapText;			//Bool	折り返し
	public	$shrinkToFit;		//Bool	縮小
	public	$indent;			//Int インデント
	//罫線　[0]str:borderStyle  [1]str:color
	public	$bdrtop;			//
	public	$bdrbottom;
	public	$bdrleft;
	public	$bdrright;
	
	public	$dwork;
}
//=======================================
/**
 *	エクセルをPDFへ出力をサポートするクラス
 *
 */
class	excel2pdf
{
	public	$errMessage		= '';
	public	static	$COEFF	= 25.4;			//長さ(mm)の係数 point:0.3528 インチ:0.3937/25.4
	public	static	$POINT	= 0.3528;
	//PDF
	private	static $fontTbl		= array();		//使用するフォント達
	private	static $fontcorrTbl	= array(		//フォントの対応表
					'游ゴシック'		=> 'kozgopromedium',
					'Yu Gothic'			=> 'kozgopromedium',
					'ＭＳ Ｐゴシック'	=> 'kozgopromedium',
					'ＭＳ ゴシック'		=> 'kozgopromedium',
					'ＭＳ Ｐ明朝'		=> 'kozminproregular',
					'ＭＳ 明朝'			=> 'kozminproregular');
	private	$pdfFileName	= 'sheet.pdf';		//出力PDFファイル名
	private	$tcpdf			= null;
	
	//Excel
	private	$excelFileName	= '';				//エクセルファイル名
	private	$book			= null;				//ブックオブジェクト
	private	$sheet			= null;				//シートオブジェクト
	private	$defontsiz		= 12.0;				//
	
	//PDF
		//単位（mm）
	//Excel
	private	$orient	= 'L';			//方向
	private	$pgsize	= 'A4';			//用紙
	
	private	$area	= '';			//印刷範囲
	private	$spclm	= 0;
	private	$sprow	= 0;
	private	$epclm	= 0;
	private	$eprow	= 0;
	private	$csize	= array();		//カラムサイズ
	private	$rsize	= array();		//行サイズ
	private	$cpos	= array();		//カラム位置（X）
	private	$rpos	= array();		//行位置（Y）
	private	$recio	= 1.0;			//印刷倍率
	private	$margentop;				//ページ余白
	private	$margenleft;			//ページ余白
	
	private	$cells	= array();		//印刷範囲のセル情報
	private	$draws	= array();		//図形群
	public	$clmWidthPt= array();	//カラム幅の初期値（エクセルファイルから取得できない時用）

	//-------------------------------------------------
	//	コンストラクタ
	function	__construct ()
	{
	//	$w	= \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToPoints(72);	//デフォルトの幅
		$w	= 56;
		for($i=1;$i<50;$i++) {
			$this->clmWidthPt[$i]	= $w;
		}
	}
	
	//-------------------------------------------------
	/**
	 *	PDF出力で使用するフォント群を登録する
	 *	@param	$ftbl	フォントファイルの配列（順番が番号となる）
	 *	@return	TRUE/FLSE
	 */
	public static function addUseFonts($ftbl)
	{
		excel2pdf::$fontTbl	= $ftbl;		//
	}
	
	//-------------------------------------------------
	/**
	 *	エクセルのフォントに対応するPDF用フォントを指定する。
	 *	@param	$wfont	エクセルのフォント名
	 *	@param	$pfont	対応するPDFフォント名（gothic/mincho）
	 *	@return	なし
	 */
	public static function setPdfFont($wfont,$pfont)
	{
		$pf	= mb_strtolower($pfont);
		switch($pf) {
		case 'ゴシック':
		case 'gothic':
			$pfont	= 'kozgopromedium';
			break;
		case '明朝':
		case 'mincho':
			$pfont	= 'kozminproregular';
			break;
		}
		excel2pdf::$fontcorrTbl[$wfont]	= $pfont;
	}
	
	//-------------------------------------------------
	/**
	 *	出力するPDFファイル名を設定する
	 *	@param	$fn	ファイル名
	 */
	public function setPdfFilename($fn)
	{
		$this->pdfFileName	= $fn;
	}
	
	//-------------------------------------------------
	/**
		出力するPDFファイル名を取得する
		@return	ファイル名
	*/
	public function getPdfFilename()
	{
		return	$this->pdfFileName;
	}
	
	//-------------------------------------------------
	/**
	 *	PDFファイル出力する
	 *	@return	なし
	 */
	public function writePDF()
	{
		//TCPDF初期化
		$this->tcpdf = new \TCPDF($this->orient, "mm", $this->pgsize, true, "UTF-8" );
		
		// 出力するPDFの初期設定
		//使用するフォント
		$this->tcpdf->setPrintHeader( false );    
		$this->tcpdf->setPrintFooter( false );
		$this->tcpdf->AddPage();
		$this->tcpdf->SetFont('kozgopromedium','',14);		//小塚ゴシック
//		$this->tcpdf->SetFont('kozminproregular','',14);	//小塚明朝
//		$this->tcpdf->SetFont('msungstdlight','',14);
//		$this->tcpdf->SetFont('stsongstdlight','',14);
		
		for($r=$this->sprow;$r<=$this->eprow; $r++) {
			for($c=$this->spclm; $c<=$this->epclm; $c++) {
				//セルの取得
				if(!empty($this->cells[$r][$c])) {
					$cell	= $this->cells[$r][$c];
					$this->writeCell($cell);
				}
			}
		}
		
		$this->tcpdf->Output($this->pdfFileName, "I");
		return;
	}
	
	//-------------------------------------------------
	/**
	 *	PDFにセルを出力する
	 *	@param	$cell	セル情報
	 *	@return	なし
	 */
	public function writeCell($cell)
	{
		$border		= '';			//枠線（後で作り変える？線種対応）
		$align		= 'L';			//横方向の位置
		$fill		= $cell->filltype;
		$stretch	= 0;			//テキストの伸縮モード
		$valign		= 'M';			//縦方向の位置
		if($cell->bdrtop[0]   !='none')	$border .= 'T';
		if($cell->bdrbottom[0]!='none')	$border .= 'B';
		if($cell->bdrleft[0]  !='none')	$border .= 'L';
		if($cell->bdrright[0] !='none')	$border .= 'R';
		
		if($cell->HAlignment=='left')	$align = 'L';
		if($cell->HAlignment=='right')	$align = 'R';
		if($cell->HAlignment=='center')	$align = 'C';

		if($cell->VAlignment=='top')	$valign = 'T';
		if($cell->VAlignment=='center')	$valign = 'M';
		if($cell->VAlignment=='bottom')	$valign = 'B';
		
	//	$fill		= false;
		if((!empty($border))||($cell->strVal!='')||($fill)) {
			//フォントの指定
			if(strlen($cell->color)==6) {
				$col	= '#'.$cell->color;
				$cr	= 255-hexdec(substr($col, 1, 2));
				$cg	= 255-hexdec(substr($col, 3, 2));
				$cb	= 255-hexdec(substr($col, 5, 2));
				$this->tcpdf->SetTextColor($cr,$cg,$cb,0);
			}
		//	else {
		//		$this->tcpdf->SetTextColor(255,255,255,0);
		//	}
			$fntnm	= $this->getFontName($cell->Font);		//フォント
			$style	= $this->getFontStyle($cell);
			$fntsiz	= $cell->FontSize * $this->recio;
			$posx	= $cell->posx   * $this->recio;
			$posy	= $cell->posy   * $this->recio;
			$width	= $cell->width  * $this->recio;
			$height	= $cell->height * $this->recio;
			if($fill) {		//セルの色
				$col	= '#'.$cell->bgcolor;
				$cr	= 255-hexdec(substr($col, 1, 2));
				$cg	= 255-hexdec(substr($col, 3, 2));
				$cb	= 255-hexdec(substr($col, 5, 2));
				$this->tcpdf->SetFillColor($cr,$cg,$cb,0);
			}
			$stretch	= 0;
			if($cell->shrinkToFit) {
				$stretch	= 1;
			}
			
			$this->tcpdf->SetXY( $posx, $posy, true);
			$this->tcpdf->SetFont($fntnm,$style,$fntsiz);
			if(empty($cell->wrapText) ) {
				//ボックス（セル）の描画
				$strVal = $cell->strVal;
				if($align == 'L') $strVal = str_repeat('  ', $cell->indent) . $cell->strVal;
				if($align == 'R') $strVal = $cell->strVal . str_repeat('  ', $cell->indent);
				$this->tcpdf->Cell( $width, $height, $strVal,		//サイズ、文字列
										'',  $stretch, $align, $fill, '', $stretch, true, 'T', $valign );
					//				$border, $stretch, $align, $fill, '', $stretch, true, 'T', $valign );
			}
			else {
				//マルチライン
				$this->tcpdf->MultiCell( $width, $height, $cell->strVal,		//サイズ、文字列
										'',	 $align, $fill, 0, $posx, $posy,
					//				$border, $align, $fill, 0, $posx, $posy,
									true, $stretch, false, true, 0, $valign, false );
			}

			//罫線を別に描画する
			if(!empty($border)) {
				$this->drawBorder($cell, $cell->bdrtop[0], $cell->bdrtop[1],			//上
									$posx, $posy, $posx+$width, $posy);
				$this->drawBorder($cell, $cell->bdrbottom[0], $cell->bdrbottom[1],		//下
									$posx, $posy+$height, $posx+$width, $posy+$height);
				$this->drawBorder($cell, $cell->bdrleft[0], $cell->bdrleft[1],			//左
									$posx, $posy, $posx, $posy+$height);
				$this->drawBorder($cell, $cell->bdrright[0], $cell->bdrright[1],		//右
									$posx+$width, $posy, $posx+$width, $posy+$height);
				$border	= '';
			}
		}
	}
	
	//-------------------------------------------------
	/**
	 *	フォント名の変換
	 *	@param	$wf		Windowsのフォント名
	 *	@return	pdfのフォント名
	 */
	public function getFontName($wf)
	{
		$ret	= '';	// 'kozgopromedium';
		if(isset(excel2pdf::$fontcorrTbl[$wf])) {
			$ret	= excel2pdf::$fontcorrTbl[$wf];
		}
		else {
			if(strpos($wf,'ゴシック')!==false) {
				$ret	= 'kozgopromedium';
			}
			if(strpos($wf,'明朝')!==false) {
				$ret	= 'kozminproregular';
			}
		}
		return	$ret;
	}
	
	//-------------------------------------------------
	/**
	 *	フォントスタイル
	 *	@param	$cell		セル情報
	 *	@return	フォントのスタイル文字列
	 */
	public function getFontStyle($cell)
	{
		$ret	= '';
		if($cell->bold)			$ret .= 'B';
		if($cell->italic)		$ret .= 'I';
		if($cell->strikethrough)$ret .= 'D';
		if($cell->underline!=\PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE) {
								$ret .= 'U';
		}
		return	$ret;
	}
	
	//-------------------------------------------------
	/**
	 *	罫線の描画
	 *	@param	$cell	セル情報
	 *	@param	$sx...	座標
	 *	@return	フォントのスタイル文字列
	 */
	public function drawBorder($cell,$style, $clr, $sx, $sy, $ex, $ey)
	{
		if($style !='none') {
			$line['width']	= 0.4;			//
			$line['cap']	= 'butt';		//末端部：butt, round, square
			$line['join']	= 'miter';		//結合部：miter, round, bevel
			$line['dash']	= '';			//on,off
			$line['phase']	= 0;			//破線の開始位置
			$line['color']	= array('R'=>0,'G'=>0,'B'=>0);
			if(strlen($clr)==6) {
				$line['color']['R']	= hexdec(substr($clr, 1, 2));
				$line['color']['G']	= hexdec(substr($clr, 3, 2));
				$line['color']['B']	= hexdec(substr($clr, 5, 2));
			}
			switch($style) {
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOT:		//一点鎖線
				$line['dash']	= '4.0,2.0,1.0,2.0';		//on,off
				$line['width']	= 0.2;						//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOTDOT:		//二点鎖線
				$line['dash']	= '4.0,2.0,1.0,2.0,1.0,2.0';	//on,off
				$line['width']	= 0.2;						//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHED:			//
				$line['dash']	= '2.0,2.0';				//on,off
				$line['width']	= 0.2;						//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOTTED:			//
				$line['dash']	= '0.5,0.5';				//on,off
				$line['width']	= 0.2;						//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE:			//二重線
				$line['width']	= 0.10;						//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR:			//
				$line['width']	= 0.1;						//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM:			//
				$line['width']	= 0.4;						//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOT:	//
				$line['dash']	= '4.0,2.0,1.5,2.0';		//on,off
				$line['width']	= 0.4;						//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOTDOT:	//
				$line['dash']	= '4.0,3.0,2.0,3.0,2.0,3.0';	//on,off
				$line['width']	= 0.4;						//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHED:		//
				$line['dash']	= '3.0,1.5';				//on,off
				$line['width']	= 0.4;						//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_SLANTDASHDOT:	//
				$line['dash']	= '4.0,1.0,2.0,1.0,2.0,1.0';	//on,off
				$line['width']	= 0.4;						//太さ 
				//@@@@@@
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK:			//
				$line['width']	= 0.6;							//太さ 
				break;
			case \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN:			//直線
				$line['width']	= 0.2;							//太さ 
				break;
			}
			if($style!=\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE) {
				//ラインを引く
				$this->tcpdf->Line($sx, $sy, $ex, $ey,$line);
			}
			else {
				if($sx==$ex) {		//縦線
					$this->tcpdf->Line($sx-0.15, $sy, $ex-0.15, $ey,$line);
					$this->tcpdf->Line($sx+0.15, $sy, $ex+0.15, $ey,$line);
				}
				else {				//横線
					$this->tcpdf->Line($sx, $sy-0.15, $ex, $ey-0.15,$line);
					$this->tcpdf->Line($sx, $sy+0.15, $ex, $ey+0.15,$line);
				}
			}
		}
	}
	
	//-------------------------------------------------
	/**
	 *	★セルの確認（デバッグ用）
	 *	@param	$c,$r	セル位置
	 *	@return	なし
	 */
	public function debugExcelCell($c,$r)
	{
		//セルの取得
		$cell	= $this->sheet->getCellByColumnAndRow($c,$r,false);
		if($cell!=null) {
			$mg		= $cell->isMergeRangeValueCell();	//代表セル
			$marge	= $cell->getMergeRange();			//結合範囲　　結合してなければ空
			if($mg) {
				$cell	= $this->getExcelCell($r,$c,$cell,$marge);
			}
			else {
				$cell	= $this->getExcelCell($r,$c,$cell,'');
			}
			echo "<br>行：${r} カラム:${c}  marge=${mg} ${marge}";
			echo "<br>座標　X:{$cell->posx}  Y:{$cell->posy}";
			echo "<br>サイズX:{$cell->width}  Y:{$cell->height}";
			echo "<br>データ　:{$cell->strVal}";
			echo "<br>フォント:{$cell->Font}　　size:{$cell->FontSize} wrap:{$cell->wrapText} 色:{$cell->color}";
			echo "<br>罫線 上下:{$cell->bdrtop[0]} {$cell->bdrbottom[0]}";
			echo "<br>罫線 左右:{$cell->bdrleft[0]} {$cell->bdrright[0]}";
			echo "<br>セル 背景:{$cell->bgcolor} {$cell->bgcolor2} {$cell->filltype}";
			echo "<br>結合:{$cell->dwork}";
			echo "<br>";
		}
		else {
			echo "<br>行：${r} カラム:${c}  NULL";
			echo "<br>";
		}
	}

	//-------------------------------------------------
	/**
	 *	★セルの確認（デバッグ用）
	 *	@param	$c,$r	セル位置
	 *	@return	なし
	 */
	public function debugCell($c,$r)
	{
		if($c==0) {
			echo "<br>カラムサイズ:<br>";
			var_dump($this->csize);
		}
		if($r==0) {
			echo "<br>行高さ:<br>";
			var_dump($this->rsize);
		}
		if(($r>0)&&($c>0)) {
			$cell	= null;
			if(isset($this->cells[$r][$c])) {
				$cell	= $this->cells[$r][$c];
			}
			if(!empty($cell)) {
				echo "<br>行：${r} カラム:${c}";
				echo "<br>座標　X:{$cell->posx}  Y:{$cell->posy}";
				echo "<br>サイズX:{$cell->width}  Y:{$cell->height}";
				echo "<br>データ　:{$cell->strVal}";
				echo "<br>フォント:{$cell->Font}　　size:{$cell->FontSize} wrap:{$cell->wrapText} 色:{$cell->color}";
				echo "<br>罫線 上下:{$cell->bdrtop[0]} {$cell->bdrbottom[0]}";
				echo "<br>罫線 左右:{$cell->bdrleft[0]} {$cell->bdrright[0]}";
				echo "<br>セル 背景:{$cell->bgcolor} {$cell->bgcolor2} {$cell->filltype}";
				echo "<br>結合:{$cell->dwork}";
				echo "<br>";
			}
			else {
				echo "<br>行：${r} カラム:${c} セル無し<br>";
			}
		}
		else {
			echo "<br>余白　　top:{$this->margentop}　　left:{$this->margenleft}";
			echo "<br>倍率　　{$this->recio}<br>";
		}
	}
	
	//-------------------------------------------------
	/**
	 *	カラム幅をピクセル値の配列で設定する
	 *	@param	$wa	カラム幅(px)の配列（1～）
	 */
	public function setColumnWidthsPx($wa, $cot=10)
	{
		if(!is_array($wa)) {
			if(!is_numeric($wa)) {
				$wa	= 172;
			}
			$wa	= array_fill( 1, $cot, $wa );
		}
		$this->clmWidthPt[0]	= \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToPoints(172);	//デフォルトの幅
		foreach($wa as $i => $v){
			$v	= \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToPoints($v);
			$this->clmWidthPt[$i]	= $v;
		}
	}
	
	//-------------------------------------------------
	/**
	 *	カラム幅をmm値の配列で設定する
	 *	@param	$wa	カラム幅(mm)の配列（1～）
	 */
	public function setColumnWidthsmm($wa, $cot=10)
	{
		if(!is_array($wa)) {
			if(!is_numeric($wa)) {
				$wa	= 11.0;
			}
			$wa	= array_fill( 1, $cot, $wa );
		}
		foreach($wa as $i => $v){
			$this->clmWidthPt[$i]	= $v / excel2pdf::$POINT;
		}
	}
	
	//-------------------------------------------------
	/**
	 *	エクセルファイル名を設定する
	 *	@param	$fn	ファイル名
	 */
	public function setExcelFilename($fn)
	{
		//ファイルがあるか？
		if(file_exists($fn)) {
			$ext = pathinfo($fn, PATHINFO_EXTENSION);
			$this->excelFileName	= $fn;
			//エクセルファイルを読込む
			if($ext=='xlsx') {
				$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
			}
			else {
				$reader	= new \PhpOffice\PhpSpreadsheet\Reader\Xls();
			}
			$reader->setIncludeCharts(true);
			$book	= $reader->load($fn);
			$this->setBook($book);
		}
		else {
			$errMessage	= 'File not Exist!! filename='.$fn;
		}
	}
	
	//-------------------------------------------------
	/**
	 *	エクセルファイル名を取得する
	 *	@return	ファイル名
	 */
	public function getExcelFilename()
	{
		return	$this->excelFileName;
	}
	
	//-------------------------------------------------
	/**
	 *	エクセルブックを設定する
	 *	@param	$book	ファイル名
	 */
	public function setBook($book)
	{
		$this->book	= $book;
		//デフォルトのスタイル（フォント）
		$style	= $book->getDefaultStyle();
		$f	= $style->getFont();
		$this->defontsiz	= $f->getSize();
		if(empty($this->sheet)) {
			$sheet	= $book->getActiveSheet();
			//シートを設定
			$this->setSheet($sheet);
		}
	}
	
	//-------------------------------------------------
	/**
	 *	エクセルブックを取得する
	 *	@return	エクセルブック
	 */
	public function getBook()
	{
		return	$this->book;
	}
	
	//-------------------------------------------------
	/**
	 *	エクセルシートを設定する
	 *	@param	$sheet	エクセルシート
	 */
	public function setSheet($sheet)
	{
		if(empty($this->book)) {
			$book	= $sheet->getParent();
			$this->setBook($book);
		}
		//シートを保持
		$this->sheet	= $sheet;
		//印刷エリアの取得
		$page	= $sheet->getPageSetup();
		$this->area	= $page->getPrintArea();
		
		//方向
		$ori		= $page->getOrientation();
		// PageSetup::ORIENTATION_LANDSCAPE
		// PageSetup::ORIENTATION_PORTRAIT
		if($ori==PageSetup::ORIENTATION_PORTRAIT) {
			$this->orient	= 'P';
		}
		$siz		= $page->getPaperSize();
		
		switch($siz) {
		case PageSetup::PAPERSIZE_LEGAL:
			$this->pgsize	= 'LEGAL';
			break;
		case PageSetup::PAPERSIZE_LETTER:
			$this->pgsize	= 'LETTER';
			break;
		case PageSetup::PAPERSIZE_A3:
			$this->pgsize	= 'A3';
			break;
		case PageSetup::PAPERSIZE_A4:
			$this->pgsize	= 'A4';
			break;
		case PageSetup::PAPERSIZE_A5:
			$this->pgsize	= 'A5';
			break;
		case PageSetup::PAPERSIZE_B4:
			$this->pgsize	= 'B4';
			break;
		case PageSetup::PAPERSIZE_B5:
			$this->pgsize	= 'B5';
			break;
		}
		//印刷倍率
		$scale	= $page->getScale();
		if($scale<=0.1) $scale = 100.0;			//念のため
		$this->recio	= $scale / 100.0;
		
		//マージン（単位は？）
		$this->margentop	= ($sheet->getPageMargins()->getTop()  * excel2pdf::$COEFF) / $this->recio;
		$this->margenleft	= ($sheet->getPageMargins()->getLeft() * excel2pdf::$COEFF) / $this->recio;
		
		//セル群を読込む
		$this->loadCells();
	}
	
	//-------------------------------------------------
	/**
	 *	エクセルシートを取得する
	 *	@return	エクセルシート
	 */
	public function getSheet()
	{
		return	$this->sheet;
	}
	
	//-------------------------------------------------
	/**
	 *	エクセルシートからセルの情報を読取る
	 */
	public function loadCells()
	{
		//印刷範囲を内部に設定
		$r	= $this->area2index($this->area);
		$s	= $r['sp'];
		$e	= $r['ep'];
		$this->sprow	= $s[0];	//開始行
		$this->spclm	= $s[1];	//終了行
		$this->eprow	= $e[0];	//開始カラム
		$this->epclm	= $e[1];	//終了カラム
		
		//カラムサイズ
	//	$this->sheet->calculateColumnWidths();			/* ???  */
		$def	= $this->sheet->getDefaultColumnDimension();
		$defw	= $def->getWidth() * $this->defontsiz / 2.0;			//11ポイント？
/*	* /	if($defw<=0) {
			if(isset($this->clmWidthPt[0])) {
				$defw	= $this->clmWidthPt[0];
			}
			else {
	//			$defw	= \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToPoints(172);	//デフォルトの幅 pt
			}
		}
*/
		$defw	/= $this->recio;
		$w		= 0.0;
		$w		= $this->margenleft;		//余白

	//	for($i=$this->spclm; $i<=$this->epclm; $i++) {
	//		if( isset($dims[$i]) ) {
	//			$dims[$i]->setAutoSize(false);
	//		}
	//	}
	//	$this->sheet->calculateColumnWidths();			/* ???  */

		//デフォルトのカラム幅
		$v	= $defw * excel2pdf::$POINT;		//mm へ変換
//		$v	= $defw;	// * excel2pdf::$POINT;		//mm へ変換
// echo "defw =${v}<br>";
		for($i=$this->spclm; $i<=$this->epclm; $i++) {
			$this->csize[$i]	= $v;
		}
//echo "カラム幅mm<br>";
//var_dump($this->csize);
//echo "<br><br><br>";
/* */
		$dims	= $this->sheet->getColumnDimensions();
		foreach ($dims as $ky => $ob) {
			$i	= \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($ky);
			$v	= $ob->getWidth() * $this->defontsiz / 2.0 ;	//半角文字数からポイントへ
			$v	= $v * excel2pdf::$POINT;		//mm へ変換
			if($v>0.0) {
				$this->csize[$i]	= $v;
			}
		}
//echo "カラム幅mm<br>";
//@var_dump($this->csize);
//@echo "<br><br><br>";
/* */
		for($i=$this->spclm; $i<=$this->epclm; $i++) {
			$v	= $this->csize[$i];
			$this->cpos[$i]	= $w;
			$w	+= $v;
		}
/* * /
var_dump($this->cpos);
echo "<br><br><br>";
 exit();
/*	* /	$dims	= $this->sheet->getColumnDimensions();
		for($i=$this->spclm; $i<=$this->epclm; $i++) {
			$vv	= -1;
			if( isset($dims[$i]) ) {
				$vv	= $dims[$i]->getWidth();
			}
			if( $vv<=0 ) {
				if( isset($this->clmWidthPt[$i])) {
					$vv	= $this->clmWidthPt[$i] / $this->recio;
				}
				else {
					$vv	= $defw;
				}
			}
			$v	= $vv * excel2pdf::$POINT;		//mm へ変換
			$this->csize[$i]	= $v;
			$this->cpos[$i]		= $w;
			$w	+= $v;
		}
/* * /
var_dump($this->csize);
echo "<br><br><br>";
var_dump($this->cpos);
echo "<br><br><br>";
exit();
/* */
		//行サイズ
		$def	= $this->sheet->getDefaultRowDimension();
		$defw	= $def->getRowHeight();
		if($defw<=0) $defw = 18.75;
		$dims	= $this->sheet->getRowDimensions();
		$w		= 0.0;
		$w		= $this->margentop;
		for($i=$this->sprow; $i<=$this->eprow; $i++) {
			if( isset($dims[$i]) ) {
				$v	= $dims[$i]->getRowHeight();
			}
			else {
				$v	= $defw;
			}
			$v	= $v * excel2pdf::$POINT;		//mm へ変換
			$this->rsize[$i]	= $v;
			$this->rpos[$i]		= $w;
			$w	+= $v;
		}
		//印刷範囲のセルの情報を取得、nullは結合セルなど
		$this->cells	= array();				//印刷範囲のセル情報
		for($r=$this->sprow;$r<=$this->eprow; $r++) {
			for($c=$this->spclm; $c<=$this->epclm; $c++) {
				//結合済みのセル
				if(isset($this->cells[$r][$c])) {
					$this->cells[$r][$c]	= null;
				}
				else {
					//セルの取得
					$cell	= $this->sheet->getCellByColumnAndRow($c,$r,false);
					if($cell!=null) {
						$mg		= $cell->isMergeRangeValueCell();	//代表セル
						$marge	= $cell->getMergeRange();			//結合範囲　　結合してなければ空
						if($mg==1) {				//生きているセル
							$ec	= $this->getExcelCell($r,$c,$cell,$marge);
							$this->cells[$r][$c]	= $ec;
						}
						else {
							$ec	= $this->getExcelCell($r,$c,$cell,'');
							$this->cells[$r][$c]	= $ec;
						}
					}
					else {
						$this->cells[$r][$c]	= null;
					}
				}
			}
		}
		//図形
		$this->draws	= array();
		$garr	= $this->sheet->getDrawingCollection();		//図形が取得出来ないよ
		foreach($garr as $obj){
			$nm	= $obj->getName();
			$cn	= $obj->getCoordinates();
			$x	= $obj->getOffsetX();
			$y	= $obj->getOffsetY();
		}
	}
	
	//-------------------------------------------------
	/**
	 *	セルの情報を取得する
	 *	@param	$r		行番号
	 *	@param	$c		カラム番号
	 *	@param	$cell	PhpSpreadsheetのセルオブジェクト
	 *	@param	$㎎		結合セルの範囲
	 *	@return	ExcelCellクラスのインスタンス
	 */
	public function getExcelCell($r,$c,$cell,$mg)
	{
		$bb	= null;
		$ec	= new ExcelCell();
	//	if($cell->isFormula()) {
	//		$pv	= $cell->getCalculatedValue();
	//		$cell	= $cell->setCalculatedValue($pv);
	//	}
		
		$val	= $cell->getFormattedValue();		//表示文字列
	//	if($cell->isFormula()) {
	//		$val	= $cell->getCalculatedValue();
	//	}
		$ec->strVal	= $val;
		//左上座標
		$ec->posx	= $this->cpos[$c];
		$ec->posy	= $this->rpos[$r];
		$ec->dwork	= $mg;
		if(empty($mg)) {
			$ec->width	= $this->csize[$c];
			$ec->height	= $this->rsize[$r];
		}
		else {
			$ar	= excel2pdf::area2index($mg);		//結合範囲の合計
			//結合部分のセルを予め情報をセットしないように。。。
			for($ci=$ar['sp'][1]; $ci<=$ar['ep'][1] ;$ci++ ) {
				for($ri=$ar['sp'][0]; $ri<=$ar['ep'][0] ;$ri++ ) {
					$this->cells[$ri][$ci]	= 1;	//null;
				}
			}
			//念のため
			if(($c!=$ar['sp'][1])||($r!=$ar['sp'][0])) {
				return	null;
			}
			
			$w	= 0.0;
			for($i=$ar['sp'][1]; $i<=$ar['ep'][1] ;$i++ ) {
				$w	+= $this->csize[$i];
			}
			$ec->width	= $w;
			
			$w	= 0.0;
			for($i=$ar['sp'][0]; $i<=$ar['ep'][0] ;$i++ ) {
				$w	+= $this->rsize[$i];
			}
			$ec->height	= $w;
			
			//右下のセル
			$celle	= $this->sheet->getCellByColumnAndRow($ar['ep'][1],$ar['ep'][0],false);
			$stylee	= $celle->getStyle();
			$bb	= $stylee->getBorders();
			//★罫線
			$k	= $bb->getBottom();			//下
			$ec->bdrbottom[0]	= $k->getBorderStyle();
			$ec->bdrbottom[1]	= $k->getColor()->getRGB();
			$k	= $bb->getRight();			//右
			$ec->bdrright[0]	= $k->getBorderStyle();
			$ec->bdrright[1]	= $k->getColor()->getRGB();

			$cell	= $this->sheet->getCellByColumnAndRow($c,$r,false);
		}
		
		$style	= $cell->getStyle();
		//セル
		$Fill	= $style->getFill();
		$bgclr	= $Fill->getStartColor();
		$ec->bgcolor		= $bgclr->getRGB();
		$bgclr	= $Fill->getEndColor();
		$ec->bgcolor2		= $bgclr->getRGB();
		$ec->filltype		= $Fill->getFillType()==\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID;
		//フォント情報
		$f	= $style->getFont();
		$ec->Font			= $f->getName();
		$ec->FontSize		= $f->getSize();	// * excel2pdf::$POINT;
		$ec->bold			= $f->getBold();
		$ec->italic			= $f->getItalic();
		$ec->superscript	= $f->getSuperscript();
		$ec->subscript		= $f->getSubscript();
		$ec->underline		= $f->getUnderline();
		$ec->strikethrough	= $f->getStrikethrough();
		$ec->color			= $f->getColor()->getRGB();
		
		//アライメント等
		$a	= $style->getAlignment();
		$ec->HAlignment		= $a->getHorizontal();
		$ec->VAlignment		= $a->getVertical();
		$ec->wrapText		= $a->getWrapText();
		$ec->shrinkToFit	= $a->getShrinkToFit();
		$ec->indent			= $a->getIndent();
		
		//罫線
		$b	= $style->getBorders();
		
		$k	= $b->getTop();				//上
		$ec->bdrtop[0]		= $k->getBorderStyle();
		$ec->bdrtop[1]		= $k->getColor()->getRGB();
		$k	= $b->getLeft();			//左
		$ec->bdrleft[0]		= $k->getBorderStyle();
		$ec->bdrleft[1]		= $k->getColor()->getRGB();
		
		if(empty($bb)) {
			$k	= $b->getBottom();			//下
			$ec->bdrbottom[0]	= $k->getBorderStyle();
			$ec->bdrbottom[1]	= $k->getColor()->getRGB();
			$k	= $b->getRight();			//右
			$ec->bdrright[0]	= $k->getBorderStyle();
			$ec->bdrright[1]	= $k->getColor()->getRGB();
		}
		return	$ec;
	}
	
	//-------------------------------------------------
	/**
	 *	エリアを表す文字列からカラム・ロウの
	 *	インデックス値(1～)へ変換する。
	 *
	 *	@param	$area	セル範囲を表す文字列 ex)'B1:G12'
	 *	@return	セル範囲を表す連想配列 $ret= ['sp'=>(行、列)、'ep'=>(行、列)]
	 */
	public static function area2index($area)
	{
		$ret	= array('sp'=>array(0,0),'ep'=>array(0,0));
		if(!empty($area)) {
			$p	= explode(':',$area);
			$sp	= excel2pdf::name2index($p[0]);
			$ep	= excel2pdf::name2index($p[1]);
			$ret	= array('sp'=>$sp,'ep'=>$ep);
		}
		return $ret;
	}

	//-------------------------------------------------
	/**
	 *	セル位置を表す文字列から行・列番号へ変換する
	 *
	 *	@param	$cn	セル位置の文字列 ex)'C4'
	 *	@return	セル位置のインデックスを表す配列 $ret=（0:行、1:列）
	 */
	public static function name2index($cn)
	{
		$cidx = $ridx = 0;
		for($i=0;$i< strlen($cn); $i++) {
			if(ctype_digit($cn[$i])==TRUE) {
				$clm	= substr( $cn, 0, $i );
				$row	= substr( $cn, $i );
				$cidx	= \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($clm);
				$ridx	= (int)$row;
				break;
			}
		}
		return array($ridx,$cidx);
	}
}
//----------------------------------- eof -----------------------------------
