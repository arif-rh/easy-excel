<?php

namespace Arifrh\EasyExcel;

class EasyExcel
{
    public $_xls = NULL;

    public function __construct($title = 'Sheet 1', $sheet_index = 0, $spreadsheet = null) 
    {
        if (is_object($spreadsheet))
        {
            $this->_xls = $spreadsheet;
        }
        else
        {
            $this->_xls = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        }
        $this->setSheetTitle($title, $sheet_index);
    }
    
    public function setSheetTitle($title = '', $sheet_index = 0)
    {
        $this->_xls->setActiveSheetIndex($sheet_index);
        $this->_xls->getActiveSheet()->setTitle($title);

        return $this;
    }

    public function setActiveSheet($index_or_name = null)
    {
        if (is_string($index_or_name))
            $this->_xls->setActiveSheetIndexByName($index_or_name);

        if (is_int($index_or_name))
            $this->_xls->setActiveSheetIndex($index_or_name);

        return $this;
    }

    public function cloneSheet($sheet_source = '', $destination = '')
    {
        $destination_sheet = empty($destination) ? $sheet_source." (2)" : $destination;
        $clonedWorksheet = clone $this->_xls->getSheetByName($sheet_source);
        $clonedWorksheet->setTitle($destination_sheet);
        $this->_xls->addSheet($clonedWorksheet);

        return $this;
    }

    public function removeSheet($index_or_name = null)
    {
        if (is_string($index_or_name))
            $index_or_name = $spreadsheet->getIndex(
                $this->_xls->getSheetByName($index_or_name)
            );

        $this->_xls->removeSheetByIndex($index_or_name);
    
        return $this;
    }

    public function setColumnsWidth($column_widths)
	{
		foreach($column_widths as $col => $width)
			$this->_xls->getActiveSheet()->getColumnDimension($col)->setWidth($width);

        return $this;
	}

    public function setRowHeight($row, $height)
    {
        $this->_xls->getActiveSheet()->getRowDimension($row)->setRowHeight($height);
        return $this;
    }
    public function setAutoWidthColumn($col)
    {
        $this->_xls->getActiveSheet()->getColumnDimension($col)->setAutoSize(true);
        return $this;
    }

    public function setColumnHeader($colStart, $colEnd, $value, $color = 'f2c5d9f1', $alignments = ['HL','VT'])
	{
		$this
            ->setMergeCellsValue($colStart, $colEnd, $value, $alignments)
            ->setBackgroundColor($colStart, $color);

		return $this;
	}

    public function setLabel($col, $value, $color = 'A1C1FB', $alignments = ['HL','VT'])
	{
		$this
            ->setCellValue($col, $value, $alignments)
            ->setBackgroundColor($col, $color);

		return $this;
    }
    
    public function setLabels($cellValues = [], $color = 'A1C1FB', $alignments = ['HL','VT'])
	{
		if (!empty($cellValues) && is_array($cellValues))
        {
            foreach($cellValues as $col => $value)
            {
                $this->setLabel($col, $value, $color, $alignments);
            }
        }

		return $this;
	}

    public function setMergeCellsValue($colStart, $colEnd, $value, $alignments = ['HL','VT'])
	{
		$this
			->setCellValue($colStart, $value, $alignments)
			->mergeCells($colStart, $colEnd);

        return $this;
    }
    
    public function setMergeLinkValue($colStart, $colEnd, $value, $alignments = ['HL','VT'])
	{
        $this
			->setLinkValue($colStart, $value, $alignments)
			->mergeCells($colStart, $colEnd);

        return $this;
    }
    
    public function setLinkValue($col, $value, $tooltip = 'Click here to visit site', $alignments = ['HL','VT'])
	{
        $this->setCellValue($col, $value, $alignments);

        $this->_xls->getActiveSheet()
            ->getCell($col)
            ->getHyperlink()
            ->setUrl($value)
            ->setTooltip($tooltip);

		return $this;
    }
    
    public function setImageValue($col, $image_path, $name = 'Image', $height = 120, $offsetX = 10, $offsetY = 10)
    {
        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
        $drawing->setName($name);
        $drawing->setDescription($name);
        $drawing->setPath($image_path);
        $drawing->setHeight($height);
        $drawing->setCoordinates($col);
        $drawing->setOffsetX($offsetX);
        $drawing->setOffsetY($offsetY);

        $drawing->setWorksheet($this->_xls->getActiveSheet());
        return $this;
    }

    // add image from image gd object

    public function addImageValue($col, $imageGd, $name = 'Image', $height = 70, $offsetX = 20)
    {
        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
        $drawing->setName($name);
        $drawing->setDescription($name);
        $drawing->setImageResource($imageGd);
        $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_JPEG);
        $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
        $drawing->setHeight($height);
        $drawing->setCoordinates($col);
        $drawing->setOffsetX($offsetX);

        $drawing->setWorksheet($this->_xls->getActiveSheet());
        return $this;
    }

    public function setCellValue($col, $value, $alignments = ['HL','VT'])
	{
        $hasNewLine = stripos($value, '<br>');

        if ($hasNewLine) $value = strtr($value, ['<br>' => "\n"]);

        $this->_xls->getActiveSheet()->setCellValue($col, $value);

        if ($hasNewLine) $this->_xls->getActiveSheet()->getStyle($col)->getAlignment()->setWrapText(true);

        $this->setAlignment($col, $alignments);

		return $this;
    }
    
    // accept array parameter of cell => value keypairs

    public function setCellValues($cellValues = [], $alignments = ['HL','VT'])
	{
        if (!empty($cellValues) && is_array($cellValues))
        {
            foreach($cellValues as $col => $value)
            {
                $this->setCellValue($col, $value, $alignments);
            }
        }

		return $this;
	}

	public function mergeCells($colStart, $colEnd)
	{
		$this->_xls->getActiveSheet()->mergeCells("{$colStart}:{$colEnd}");

		return $this;
	}

	public function setAlignment($col, $alignments = ['HL','VT'])
	{
        foreach ($alignments as $alignment)
        {
            switch($alignment)
            {
                case 'HL' :
                    $this->_xls->getActiveSheet()->getStyle($col)
                        ->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
                    break;

                case 'HC' :
                    $this->_xls->getActiveSheet()->getStyle($col)
                        ->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                    break;

                case 'HR' :
                    $this->_xls->getActiveSheet()->getStyle($col)
                        ->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
                    break;

                case 'VT' :
                    $this->_xls->getActiveSheet()->getStyle($col)
                        ->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
                    break;

                case 'VC' :
                    $this->_xls->getActiveSheet()->getStyle($col)
                        ->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                    break;

                case 'VB' :
                    $this->_xls->getActiveSheet()->getStyle($col)
                        ->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM);
                    break;
            }
        }
		return $this;
	}

	public function setBackgroundColor($col, $color = 'F2FDE9D9')
	{
		$this->_xls->getActiveSheet()->getStyle($col)->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB($color);

		return $this;
    }
    
    public function getColumnLabel($index, $loop = -1)
    {
        $char = chr(65+$index);

        if (preg_match("/^[A-Z]+$/", $char))
            return ($loop >= 0 ? chr(65+$loop) : "").$char;

        $index -= 26;
        $loop++;
        return $this->getColumnLabel($index, $loop);
    }

    public function insertRows($rows = 1, $before = 1)
    {
        $this->_xls->getActiveSheet()->insertNewRowBefore($before, $rows);

        return $this;
    }

    public function forceDownload($filename, $format = 'xlsx')
    {
        $ext = '.'.$format;
        $hasExt = preg_match("/(\.)([a-z]+)$/i", $filename, $files);

        if ($hasExt)
        {
            $filename = strtr($filename, [$files[0] => $ext]);
            $ext = '';
        }

        $contentType = 'pdf';
        switch($format)
        {
            case 'pdf' :
                $contentType = 'pdf';
                break;
            case 'xls' :
            case 'xlsx' :
                $contentType = 'vnd.ms-excel';
                break;
            case 'html' :
                $contentType = 'text/html';
                break;
        }

        header('Content-Type: application/'.$contentType);
        header('Content-Disposition: attachment;filename="'.$filename.$ext.'"');
        header('Cache-Control: max-age=0');

        $writer = NULL;

        if ($format == 'xlsx')
        {
            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->_xls);
        }
        elseif ($format == 'pdf')
        {
            $writer = $this->writerPDF();
            $writer->writeAllSheets();
        }
        elseif ($format == 'html')
        {
            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Html($this->_xls);
        }

        if (is_object($writer))
            $writer->save('php://output');
    }

    public function getPDFWriter()
    {
        $this->_xls->getActiveSheet()->setShowGridlines(true);
        return new \PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf($this->_xls);
    }

    public function writerPDF()
    {
        $this->_xls->getActiveSheet()->setShowGridlines(false);
        return new EasyPDF($this->_xls);
    }

    public function saveAs($filename, $path = './')
    {
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->_xls);
        $writer->save($path.$filename);

        return file_exists($path.$filename);
    }
}
