<?php

namespace Arifrh\EasyExcel;

class Reader
{
    protected $valid_filetypes = ['Xls', 'Xlsx', 'Csv', 'Ods', 'Xml'];

    public $filetype = 'Xlsx';
    public $reader = NULL;
    public $_xls = NULL;

    protected $worksheet = NULL;
    protected $_tmp = [];

    public function __construct($filename = FALSE, $sheets = [], $autoload = FALSE, $filetype = FALSE) 
    {
        if (!$filetype)
            $filetype = ucfirst(strtolower(pathinfo($filename, PATHINFO_EXTENSION)));
        
        if (in_array($filetype, $this->valid_filetypes))
            $this->filetype = $filetype;

        if (is_string($filename) && file_exists($filename))
        {
            $readerType = "\PhpOffice\PhpSpreadsheet\Reader\\".$this->filetype;
            $this->reader = new $readerType();

            if (!empty($sheets))
                $this->reader->setLoadSheetsOnly($sheets);

            $canRead = FALSE; 
            $isExcelType = in_array($this->filetype, ['Xls', 'Xlsx']);

            if ($this->reader->canRead($filename))
            {
                $canRead = TRUE;
                $this->_xls = $this->reader->load($filename);
                $this->_tmp = [];
            }
            elseif ($isExcelType)
            {
                // try to switch type (possibilty file extension is renamed)
                $filetype = $this->filetype == 'Xlsx' ? 'Xls' : 'Xlsx';

                $readerType = "\PhpOffice\PhpSpreadsheet\Reader\\".$filetype;
                $this->reader = new $readerType();

                if ($this->reader->canRead($filename))
                {
                    $canRead = TRUE;
                    $this->_xls = $this->reader->load($filename);
                    $this->_tmp = [];
                }
            }

            if ($canRead)
            {
                if ($autoload) $this->getData();
            }
            else
            {
                if ($isExcelType)
                    die($filename.' can not be read. File type is '.$this->filetype.' but this seem not as valid extension. Try re-create the file from Microsoft Excel');
                else 
                    die($filename.' can not be read. File maybe invalid or Corrupted');
            }
        }
        else 
        {
            die($filename.' can not be read. If this has japanese text, please rename the file and try again');
        }
    }

    public function getSpreadsheet()
    {
        return $this->_xls;
    }

    public function getData($sheetIndexOrAll = true)
    {
        if ($sheetIndexOrAll === true)
        {
            for($sheetIndex=0; $sheetIndex<$this->_xls->getSheetCount(); $sheetIndex++)
                $this->getSheetValues($sheetIndex);
        }
        else 
            $this->getSheetValues((int)$sheetIndexOrAll);
        
        return $this->_tmp;
    }

    protected function getSheetValues($sheetIndex = 0)
    {
        $this->worksheet = $this->_xls->setActiveSheetIndex($sheetIndex);
        foreach ($this->worksheet->getRowIterator() as $i => $row) 
        {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(FALSE); 
            foreach ($cellIterator as $cell)
            {
                $cellValue = $cell->getValue();
                $this->_tmp[$sheetIndex][$i][] = $cellValue;
            }
        }
    }

    protected function getCellValues($excel_data = [], $has_header = true)
    {
        $fields = $values = [];
        foreach($excel_data as $row => $row_values)
        {
            if ($has_header && ($row == 1))
            {
                foreach($row_values as $index => $val)
                    $fields[$index] = $val;
            }
            else
            {
                $tmp_value = [];
                foreach($row_values as $index => $val)
                {
                    $tmp_value = array_merge($tmp_value, [$fields[$index] => trim($val)]);   
                }
                $values[] = $tmp_value;
            }
        }
        return array_filter($values);
    }

    public function toArray($has_header = true, $sheetIndexOrAll = true)
    {
        $arr = [];
        if ($sheetIndexOrAll === true)
        {
            for($sheetIndex=0; $sheetIndex<$this->_xls->getSheetCount(); $sheetIndex++)
            {
                $excel_data = $this->_tmp[$sheetIndex];
                $arr[$sheetIndex] = $this->getCellValues($excel_data);
            }
        }
        else 
        {
            $sheetIndex = (int) $sheetIndexOrAll;
            $excel_data = $this->_tmp[$sheetIndex];
            $arr[$sheetIndex] = $this->getCellValues($excel_data);
        }
        return $arr;
    }

    public function toTable($data = FALSE)
    {
        $_data = is_array($data) ? $data : $this->toArray();

        foreach ($_data as $sheetIndex) 
        {
            echo '<table border="1">' . PHP_EOL;
            $row = 1;
            foreach($sheetIndex as $rows)
            {
                if ($row == 1)
                {
                    echo '<tr>' . PHP_EOL;
                    foreach ($rows as $col => $cell) {
                        echo '<td>' . $col . '</td>' . PHP_EOL;
                    }
                    echo '</tr>' . PHP_EOL;
                }

                echo '<tr>' . PHP_EOL;
                foreach ($rows as $col => $cell) {
                    echo '<td>' . $cell . '</td>' . PHP_EOL;
                }
                $row++;
                echo '</tr>' . PHP_EOL;
            }
            echo '</table><br>' . PHP_EOL;
        }
    }
}
