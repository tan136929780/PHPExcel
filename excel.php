<?php

namespace app\models;

use Yii;

class MascReport extends BaseModel
{
    /**
     * @param  $logId
     * @param  $summary
     * @return
     */
    public function export($logId, $summary)
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
        $summary = json_decode($summary);
        $sum = [];
        $i = 0;
        foreach ($summary as $key => $record) {
            $sum[$i]['Week']      = $record->week;
            $sum[$i]['bounce_30']      = round($record->bounce_30 / 10000, 4);
            $sum[$i]['bounce_30_goal'] = round($record->bounce_30_goal / 10000, 4);
            $sum[$i]['bounce_90']      = round($record->bounce_90 / 10000, 4);
            $sum[$i]['bounce_90_goal'] = round($record->bounce_90_goal / 10000, 4);
            $i++;
        }
        $summary = array_merge($header1, $sum);
        $query  = (new \yii\db\Query())->select(['week', 'masc_code', 'masc_name', 'apc_code', 'apc_description', 'bounce_30', 'bounce_30_goal', 'bounce_90', 'bounce_90_goal', 'brand'])->where(['log_id' => $logId])->from('masc_report');
        $query->orderBy(['week' => 'desc', 'masc_code' => 'desc', 'apc_code' => 'desc', 'brand' => 'desc']);
        $records = $query->all();

        foreach ($records as $key => &$record) {
            if ($record['week'] == 'WEEK TO DATE') {
                unset($records[$key]);
                continue;
            }
            $record['bounce_30']      = round($record['bounce_30'] / 10000, 4);
            $record['bounce_30_goal'] = round($record['bounce_30_goal'] / 10000, 4);
            $record['bounce_90']      = round($record['bounce_90'] / 10000, 4);
            $record['bounce_90_goal'] = round($record['bounce_90_goal'] / 10000, 4);
        }
        $records = array_merge($header2, $records);

        // init PHPExcel
        $phpExcelPath = Yii::getAlias('@app') . '/libs/Excel';
        include($phpExcelPath . '/PHPExcel.php');
        $fileName = 'MASC_Report_Bounce_30-90' . date('_Ymd_His_') . rand(10000, 99999) . '.xlsx';
        $filePath = Yii::getAlias('@app') . '/runtime/';
        $PHPExcel = new \PHPExcel();

        // first sheet
        $currentSheet = $PHPExcel->getActiveSheet();
        $sheet_title   = 'Bounce 30-90 graph';
        $currentSheet->setTitle($sheet_title);
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
        $series[0]->setPlotDirection(\PHPExcel_Chart_DataSeries::DIRECTION_COL);
        $layout = new \PHPExcel_Chart_Layout();
        $layout->setShowPercent(true);
        $areas = new \PHPExcel_Chart_PlotArea($layout, $series);
        $legend = new \PHPExcel_Chart_Legend(\PHPExcel_Chart_Legend::POSITION_BOTTOM, $layout, false);
        $title = new \PHPExcel_Chart_Title('');
        $ytitle = new \PHPExcel_Chart_Title('');
        $chart = new \PHPExcel_Chart('line_chart', $title, $legend, $areas, true, false, $title, $ytitle);
        $chart->setTopLeftPosition("H2")->setBottomRightPosition("Q20");
        $currentSheet->addChart($chart);

        // second sheet
        $sheetname = new \PHPExcel_Worksheet();
        $sheetname->setTitle('Bounce 30-90 database');
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
        $PHPWriter->save($filePath . $fileName);
        $fileUrl  = Helper::zipFileAndUploadAws($filePath . $fileName, false);
        $response = Yii::$app->getResponse();
        $response->setDownloadHeaders(basename($fileUrl), 'application/zip');
        $response->sendFile($fileUrl);
    }
}

