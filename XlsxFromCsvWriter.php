<?php

/*
 * Original library: https://github.com/mk-j/PHP_XLSXWriter
 */

namespace Berg\Xlsx;

use SplFileObject;

class XlsxFromCsvWriter extends XlsxWriter
{
	public function writeRawFiles($xlsxRawRootPath)
	{
        if (file_exists($xlsxRawRootPath)) {
            throw new \Exception(sprintf('Directory %s already exists', $xlsxRawRootPath));
        }

        mkdir($xlsxRawRootPath);

        mkdir($xlsxRawRootPath . '/docProps');
		file_put_contents($xlsxRawRootPath . '/docProps/app.xml' , self::buildAppXML() );
		file_put_contents($xlsxRawRootPath . '/docProps/core.xml', self::buildCoreXML());

		mkdir($xlsxRawRootPath . '/_rels');
		file_put_contents($xlsxRawRootPath . '/_rels/.rels', self::buildRelationshipsXML());

        mkdir($xlsxRawRootPath . '/xl');
        mkdir($xlsxRawRootPath . '/xl/worksheets');
		foreach($this->sheets_meta as $sheet_meta) {
			rename($sheet_meta['filename'], $xlsxRawRootPath . '/xl/worksheets/' . $sheet_meta['xmlname']);
		}
		if (!empty($this->shared_strings)) {
			rename($this->writeSharedStringsXML(), $xlsxRawRootPath . 'xl/sharedStrings.xml' );
		}
		file_put_contents($xlsxRawRootPath . '/xl/workbook.xml', self::buildWorkbookXML() );
		rename($this->writeStylesXML(), $xlsxRawRootPath . '/xl/styles.xml' );

		file_put_contents($xlsxRawRootPath . '/[Content_Types].xml', self::buildContentTypesXML() );

		mkdir($xlsxRawRootPath . '/xl/_rels');
		file_put_contents($xlsxRawRootPath . '/xl/_rels/workbook.xml.rels', self::buildWorkbookRelsXML() );
	}

	
	public function writeSheetFromCsv($csvFileName, array $csvOptions = array(), $sheet_name='', array $header_types=array() )
	{
        $defaultOptions = array(
            'delimiter' => ",",
            'enclosure' => "\"",
            'escape' => "\\",
            'skip_rows' => 0
        );

        $diffKeys = array_diff_key($csvOptions, $defaultOptions);
        if (count($diffKeys) > 0) {
            throw new \Exception(sprintf('Undefined $csvOptions key(-s): %s', implode(',', $diffKeys)));
        }

        $csvOptions = array_replace($defaultOptions, $csvOptions);

        $file = new SplFileObject($csvFileName);
        $file->openFile('r');
        $file->setFlags(SplFileObject::DROP_NEW_LINE | SplFileObject::SKIP_EMPTY);
        $file->setCsvControl($csvOptions['delimiter'], $csvOptions['enclosure'], $csvOptions['escape']);

        $rowsQuantity = 0;
        $columnsQuantity = 0;
        // lets calculate rows/columns qty
        // todo: good place for optimization: we can calculate and read/write data during only one pass
        while(!$file->eof()) {
            $tmp = $file->fgetcsv();
            if ($tmp === null) {
                continue;
            }

            $columnsQuantity = max($columnsQuantity, count($tmp));

            $rowsQuantity++;
        }

        // slightly modified dirty code from https://github.com/mk-j/PHP_XLSXWriter
        // todo: use XMLWriter instead of this shit

        // removed
		// $data = empty($data) ? array( array('') ) : $data;
		
		$sheet_filename = $this->tempFilename();
		$sheet_default = 'Sheet'.(count($this->sheets_meta)+1);
		$sheet_name = !empty($sheet_name) ? $sheet_name : $sheet_default;
		$this->sheets_meta[] = array('filename'=>$sheet_filename, 'sheetname'=>$sheet_name ,'xmlname'=>strtolower($sheet_default).".xml" );

		$header_offset = empty($header_types) ? 0 : 1;

        // modified
        // $row_count = count($data) + $header_offset;
		// $column_count = count($data[self::array_first_key($data)]);
        $row_count = $rowsQuantity + $header_offset;
        $column_count = $columnsQuantity;


        $max_cell = self::xlsCell( $row_count-1, $column_count-1 );

		$tabselected = count($this->sheets_meta)==1 ? 'true' : 'false';//only first sheet is selected
		$cell_formats_arr = empty($header_types) ? array_fill(0, $column_count, 'string') : array_values($header_types);
		$header_row = empty($header_types) ? array() : array_keys($header_types);

		$fd = fopen($sheet_filename, "w+");
		if ($fd===false) { self::log("write failed in ".__CLASS__."::".__FUNCTION__."."); return; }
		
		fwrite($fd,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
		fwrite($fd,'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
		fwrite($fd,    '<sheetPr filterMode="false">');
		fwrite($fd,        '<pageSetUpPr fitToPage="false"/>');
		fwrite($fd,    '</sheetPr>');
		fwrite($fd,    '<dimension ref="A1:'.$max_cell.'"/>');
		fwrite($fd,    '<sheetViews>');
		fwrite($fd,        '<sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="'.$tabselected.'" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');
		fwrite($fd,            '<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
		fwrite($fd,        '</sheetView>');
		fwrite($fd,    '</sheetViews>');
		fwrite($fd,    '<cols>');
		fwrite($fd,        '<col collapsed="false" hidden="false" max="1025" min="1" style="0" width="11.5"/>');
		fwrite($fd,    '</cols>');
		fwrite($fd,    '<sheetData>');
		if (!empty($header_row))
		{
			fwrite($fd, '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.(1).'">');
			foreach($header_row as $k=>$v)
			{
				$this->writeCell($fd, 0, $k, $v, 'string');
			}
			fwrite($fd, '</row>');
		}

        // modified
        //
		// foreach($data as $i=>$row)
		// {
		// 	fwrite($fd, '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.($i+$header_offset+1).'">');
		//	foreach($row as $k=>$v)
		//	{
		//		$this->writeCell($fd, $i+$header_offset, $k, $v, $cell_formats_arr[$k]);
		//	}
		//	fwrite($fd, '</row>');
		//}
        $file->rewind();
        $i = 0;
        while(!$file->eof()) {
            $row = $file->fgetcsv();
            if ($row === null) {
                continue;
            }

            fwrite($fd, '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.($i+$header_offset+1).'">');

            $j = 0;
            foreach($row as $v) {
                $cell = self::xlsCell($i+$header_offset, $j);
                fwrite($fd,'<c r="'.$cell.'" t="inlineStr"><is><t>'.self::xmlspecialchars($v).'</t></is></c>');
                $j++;
            }

            fwrite($fd, '</row>');

            $i++;
        }

		fwrite($fd,    '</sheetData>');
		fwrite($fd,    '<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
		fwrite($fd,    '<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
		fwrite($fd,    '<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
		fwrite($fd,    '<headerFooter differentFirst="false" differentOddEven="false">');
		fwrite($fd,        '<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
		fwrite($fd,        '<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
		fwrite($fd,    '</headerFooter>');
		fwrite($fd,'</worksheet>');
		fclose($fd);
	}
}
