<?php

class ExcelExport
{
    public function export()
    {
        ini_set('memory_limit', '1024M');
        set_time_limit(0);

        // format data
        $header1 = [
            [
                'Week',
                'Average of Bounce 30',
                'Average of Bounce 30 Goal',
                'Average of Bounce 90',
                'Average of Bounce 90 Goal',
            ],
        ];
        $header2 = [
            [
                'Week',
                'MASC Code',
                'MASC Name',
                'Model Code',
                'Model Description',
                'Bounce 30',
                'Bounce 30 Goal',
                'Bounce 90',
                'Bounce 90 Goal',
            ],
        ];
        $summary = '';
        $summary = array_merge($header1, $summary);// first sheet data
        $records = '';
        $records = array_merge($header2, $records);// second sheet data

        // init PHPExcel
        include('/PHPExcel.php');
        $PHPExcel = new \PHPExcel();

        // first sheet
        $currentSheet = $PHPExcel->getActiveSheet();
        $sheet_title   = 'First sheet';
        $currentSheet->setTitle($sheet_title);
        // set format
        $currentSheet->getStyle('B')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet->getStyle('C')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet->getStyle('D')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet->getStyle('E')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet->fromArray($summary);
        $j = $currentSheet->getHighestRow();
        // set chart
        $labels  = [
            new \PHPExcel_Chart_DataSeriesValues('String', '!$B$1', null, 1),
            new \PHPExcel_Chart_DataSeriesValues('String', '!$D$1', null, 1),
        ];
        $labels2 = [
            new \PHPExcel_Chart_DataSeriesValues('String', '!$C$1', null, 1),
            new \PHPExcel_Chart_DataSeriesValues('String', '!$E$1', null, 1),
        ];
        $xLabels = [
            new \PHPExcel_Chart_DataSeriesValues('String', '!$A$2:$A$' . $j, null, $j - 1),
        ];
        $datas   = [
            new \PHPExcel_Chart_DataSeriesValues('Number', '!$B$2:$B$' . $j, null, $j - 1),
            new \PHPExcel_Chart_DataSeriesValues('Number', '!$D$2:$D$' . $j, null, $j - 1),
        ];
        $datas2  = [
            new \PHPExcel_Chart_DataSeriesValues('Number', '!$C$2:$C$' . $j, null, $j - 1),
            new \PHPExcel_Chart_DataSeriesValues('Number', '!$E$2:$E$' . $j, null, $j - 1),
        ];
        $series  = [
            new \PHPExcel_Chart_DataSeries(
                \PHPExcel_Chart_DataSeries::TYPE_BARCHART,
                \PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
                range(0, count($labels) - 1),
                $labels,
                $xLabels,
                $datas
            ),
            new \PHPExcel_Chart_DataSeries(
                \PHPExcel_Chart_DataSeries::TYPE_LINECHART,
                \PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
                range(0, count($labels2) - 1),
                $labels2,
                null,
                $datas2
            )
        ];
        // set direction for bar chart
        $series[0]->setPlotDirection(\PHPExcel_Chart_DataSeries::DIRECTION_COL);
        $layout = new \PHPExcel_Chart_Layout();
        $layout->setShowPercent(true);
        $areas = new \PHPExcel_Chart_PlotArea($layout, $series);
        $legend = new \PHPExcel_Chart_Legend(\PHPExcel_Chart_Legend::POSITION_BOTTOM, $layout, false);
        $title = new \PHPExcel_Chart_Title('');
        $ytitle = new \PHPExcel_Chart_Title('');
        $chart = new \PHPExcel_Chart('line_chart', $title, $legend, $areas, true, false, $title, $ytitle);
        $chart->setTopLeftPosition("H2")->setBottomRightPosition("Q20");
        // add chart to the first sheet
        $currentSheet->addChart($chart);

        // second sheet
        $sheetname = new \PHPExcel_Worksheet();
        $sheetname->setTitle('Second sheet');
        $PHPExcel->addSheet($sheetname);
        $currentSheet2 = $PHPExcel->getSheet(1);
        $currentSheet2->getStyle('F')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet2->getStyle('G')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet2->getStyle('H')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet2->getStyle('I')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet2->fromArray($records);

        // save and send
        $PHPWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, 'Excel2007');
        $PHPWriter->setIncludeCharts(true);
        $PHPWriter->save('export.xlsx');
    }
}

