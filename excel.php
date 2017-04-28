<?php

class ExcelExport
{
    public function export()
    {
        // 设置两个worksheet的第一行
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
                'Bounce 30',
                'Bounce 30 Goal',
                'Bounce 90',
                'Bounce 90 Goal',
            ],
        ];
        $summary = '';// 第一页数据
        $summary = array_merge($header1, $summary);// 第一页数据和第一行合并
        $records = '';// 第二页数据
        $records = array_merge($header2, $records);// 第二页数据和第二行合并

        // 初始化 PHPExcel
        include('/PHPExcel.php');
        $PHPExcel = new \PHPExcel();

        // 获取第一个worksheet
        $currentSheet = $PHPExcel->getActiveSheet();
        // 设置worksheet名称
        $sheet_title   = 'First sheet';
        $currentSheet->setTitle($sheet_title);
        // 设置某列的格式（实例中位百分比显示）
        $currentSheet->getStyle('B')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet->getStyle('C')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet->getStyle('D')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet->getStyle('E')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        // 载入数据
        $currentSheet->fromArray($summary);
        $j = $currentSheet->getHighestRow();

        // 设置图表

        // 设置需要处理的数据标签
        $labels  = [
            new \PHPExcel_Chart_DataSeriesValues('String', '!$B$1', null, 1),// 第二个参数如果 ！ 前不加worksheet，默认位当前worksheet
            new \PHPExcel_Chart_DataSeriesValues('String', '!$D$1', null, 1),
        ];
        $labels2 = [
            new \PHPExcel_Chart_DataSeriesValues('String', '!$C$1', null, 1),
            new \PHPExcel_Chart_DataSeriesValues('String', '!$E$1', null, 1),
        ];
        // 设置X轴的刻度（Y轴一般自动生成）
        $xLabels = [
            new \PHPExcel_Chart_DataSeriesValues('String', '!$A$2:$A$' . $j, null, $j - 1),
        ];
        // 设置每个数据标签的数据
        $datas   = [
            new \PHPExcel_Chart_DataSeriesValues('Number', '!$B$2:$B$' . $j, null, $j - 1),
            new \PHPExcel_Chart_DataSeriesValues('Number', '!$D$2:$D$' . $j, null, $j - 1),
        ];
        $datas2  = [
            new \PHPExcel_Chart_DataSeriesValues('Number', '!$C$2:$C$' . $j, null, $j - 1),
            new \PHPExcel_Chart_DataSeriesValues('Number', '!$E$2:$E$' . $j, null, $j - 1),
        ];
        // 封装数据
        $series  = [
            new \PHPExcel_Chart_DataSeries(
                \PHPExcel_Chart_DataSeries::TYPE_BARCHART,// 第一组封装成bar chart
                \PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
                range(0, count($labels) - 1),
                $labels,
                $xLabels,
                $datas
            ),
            new \PHPExcel_Chart_DataSeries(
                \PHPExcel_Chart_DataSeries::TYPE_LINECHART,// 第一组封装成line chart
                \PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
                range(0, count($labels2) - 1),
                $labels2,
                null,
                $datas2
            )
        ];

        $series[0]->setPlotDirection(\PHPExcel_Chart_DataSeries::DIRECTION_COL); // 对于 bar chart 必须设置的一项

        $layout = new \PHPExcel_Chart_Layout();
        $layout->setShowPercent(true);
        $areas = new \PHPExcel_Chart_PlotArea($layout, $series);
        $legend = new \PHPExcel_Chart_Legend(\PHPExcel_Chart_Legend::POSITION_BOTTOM, $layout, false);
        $title = new \PHPExcel_Chart_Title('');
        $ytitle = new \PHPExcel_Chart_Title('');
        $chart = new \PHPExcel_Chart('line_chart', $title, $legend, $areas, true, false, $title, $ytitle);
        // 设置图表位置（左上/右下）
        $chart->setTopLeftPosition("H2")->setBottomRightPosition("Q20");
        // 把chart加入第一个worksheet
        $currentSheet->addChart($chart);

        // 第二个worksheet
        $sheetname = new \PHPExcel_Worksheet();
        $sheetname->setTitle('Second sheet');
        $PHPExcel->addSheet($sheetname);
        $currentSheet2 = $PHPExcel->getSheet(1);
        $currentSheet2->getStyle('F')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet2->getStyle('G')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet2->getStyle('H')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet2->getStyle('I')->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
        $currentSheet2->fromArray($records);

        // 生成文件
        $PHPWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, 'Excel2007');
        $PHPWriter->setIncludeCharts(true);// 图表必须
        $PHPWriter->save('export.xlsx');
    }
}

