<?php

namespace Arifrh\EasyExcel;

class EasyPDF extends \PhpOffice\PhpSpreadsheet\Writer\Pdf\Tcpdf
{
    protected $pdf;
    protected $font = 'cid0jp';
    protected $paperSize;

    protected function createExternalWriterInstance($orientation, $unit, $paperSize)
    {
        $this->pdf = new \TCPDF($orientation, $unit, $paperSize);
        $this->pdf->setFontSubsetting(false);
        return $this->pdf;
    }
}