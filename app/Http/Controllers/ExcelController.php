<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Input;
use App\Http\Controllers\ImportController;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use Riskihajar\Terbilang;
use App\Http\Controllers\Traits\APIAccess;
use App\Http\Controllers\Traits\CurlNow;
use Carbon\Carbon;
use Mpdf\Mpdf;
use Barryvdh\DomPDF\Facade as PDF;

class ExcelController extends Controller
{
    use CurlNow, APIAccess;
    private $ControllerKey = 'wfp';

    public function purchaseRequest(Request $request, $year = null, $key, $SD, $flash) {
        if($key !== 'export_key_K2hKr7D0H9u')
            return 'NO_EXPORT_KEY';
        if ($year == null)
            $year = date('Y');
        $input = array(
            'link' => "/api/export/pr/".$year,
            'apiKey' =>  $key,
            'flash' => $flash,
            'flash2' => $SD,
            'user' => 'user_d',
        );
        // $input2 = array(
        //     'link' => "/api/export/rpr/".$year,
        //     'apiKey' =>  $key,
        //     'flash' => $flash,
        //     'flash2' => $SD,
        //     'user' => 'user_d',
        // );
        // $requestPR = $this->getcurl($input2);
        // dd($requestPR);
        $received = $this->getcurl($input);
        // dd($received);
        // $whatIWant = substr($received['user'][4], strpos($data, "||") + 1);
        // $whatIWant = strtok( $received['user'][4], '||' );
        // $whatIWant = ltrim(strstr($received['user'][4], '||'), '||');
        // dd($received);
        // $dateNow = Carbon::now();
        // dd(Carbon::today()->toFormattedDateString());

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load("xls/PR.xlsx");

        // START INSERT VALUES
        // START INSERT VALUES
        // START INSERT VALUES

        $data = [
            'header' => ['entity_name' => 'DOH-CHD-CAR','fund_cluster' => $received['allppmp'][0]['source_abbr'].' - '.$received['allppmp'][0]['parent_type_abbr'].' - '.$received['allppmp'][0]['type_abbr'],'section' => $received['user'][6],'PR_number' => $received['allppmp'][0]['prNumber'],'RCC' => 'RCC','date' => Carbon::today()->toFormattedDateString()
            ,'requested_by_name' => strtok( $received['user'][4], '||' ),'requested_by_designation' => ltrim(strstr($received['user'][4], '||'), '||')
            ,'approved_by_name' => strtok( $received['user'][2], '||' ),'approved_by_designation' => ltrim(strstr($received['user'][2], '||'), '||')
            ],
            'data' => []
            // 'data' => [
            //     ['unit' => 'unit 1','item_description' => 'item_description 1','quantity' => 'quantity 1','unit_cost' => 'unit_cost 1','total_cost' => 'total_cost 1'],
            //     ['unit' => 'unit 1','item_description' => 'item_description 1','quantity' => 'quantity 1','unit_cost' => 'unit_cost 1','total_cost' => 'total_cost 1'],
            //     ['unit' => 'unit 1','item_description' => 'item_description 1','quantity' => 'quantity 1','unit_cost' => 'unit_cost 1','total_cost' => 'total_cost 1'],
            // ],
        ];

        foreach ($received['allppmp'] as $key => $k) {
            $data['data'][$key] = ['unit' => $k['unit'],'item_description' => $k['item_name'].'.                                                           Charge to: '.$k['parent_type_abbr'].' - '.$k['source_name'].' - '.$k['type_abbr'],'quantity' => $k['qty'],'unit_cost' => $k['abc'],'total_cost' => $k['estimated_budget']];
        }

        foreach($spreadsheet->getActiveSheet()->getRowDimensions() as $rd) {
            $rd->setRowHeight(-1);
        }
        // Table Header
        // Table Header
        $spreadsheet->getActiveSheet()
            ->setCellValue('A6', 'Entity Name: '. $data['header']['entity_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D6', 'Fund Cluster: '. $data['header']['fund_cluster']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A7', 'Office/Section: '. $data['header']['section']);
        $spreadsheet->getActiveSheet()
            // ->setCellValue('C7', 'PR Number: '. $data['header']['PR_number']);
            ->setCellValue('C7', 'PR Number: ');
        $spreadsheet->getActiveSheet()
            ->setCellValue('C8', 'Responsibility Center Code: '. $data['header']['RCC']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('E7', 'Date: '. $data['header']['date']);

        // Table Footer
        // Table Footer
        $spreadsheet->getActiveSheet()
            ->setCellValue('B18',$data['header']['requested_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('B19',$data['header']['requested_by_designation']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D18',$data['header']['approved_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D19',$data['header']['approved_by_designation']);

        // Table insert
        // $sheetCol = ['A','B','C','D','E','F'];
        $startRow = 11;
        // foreach ($sheetCol as $keyi => $i) {
            foreach ($data['data'] as $keyj => $j) {

                $spreadsheet->getActiveSheet()->insertNewRowBefore($startRow, 1);
                // Stock/Property No.
                    // $spreadsheet->getActiveSheet()
                    //     ->setCellValue($i.$startRow, $j);
                // Unit
                $spreadsheet->getActiveSheet()
                    ->setCellValue('B'.$startRow, $j['unit']);//A11
                // Item Description
                $spreadsheet->getActiveSheet()
                    ->setCellValue('C'.$startRow, $j['item_description']);
                // Quantity
                $spreadsheet->getActiveSheet()
                    ->setCellValue('D'.$startRow, $j['quantity']);
                // Unit Cost
                $spreadsheet->getActiveSheet()
                    ->setCellValue('E'.$startRow, $j['unit_cost']);
                // Total Cost
                $spreadsheet->getActiveSheet()
                    ->setCellValue('F'.$startRow, $j['total_cost']);
                $spreadsheet->getActiveSheet()->getRowDimension(1)->setRowHeight(-1);
                $startRow++;
            }
        // }
        // END INSERT VALUES
        // END INSERT VALUES
        // END INSERT VALUES

        $response = response()->streamDownload(function() use ($spreadsheet) {
            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
        });
        $response->setStatusCode(200);
        $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $response->headers->set('Content-Disposition', 'attachment; filename="Purchase Request '.$received['user']['6'].'.xlsx"');
        $response->send();
     }


     public function purhcaseOrder(Request $request, $year = null, $key, $SD, $flash, $bidder_id, $POno)
     {
        $mySD = explode("xx",$SD);
        // $user = ['division_id' => $mySD[0], 'section_id' => $mySD[1]];
        if($key !== 'export_key_K2hKr7D0H9u')
            return 'NO_EXPORT_KEY';
        if ($year == null)
            $year = date('Y');
        $input = array(
            'link' => "/api/export/po/".$year,
            'apiKey' =>  $key,
            'flash' => $flash,
            'flash2' => $SD,
            'user' => $SD,
            'flash3' => $bidder_id,
        );

        $received = $this->getcurl($input);
        // dd($received);
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load("xls/PO.xlsx");

        // START INSERT VALUES
        // START INSERT VALUES
        // START INSERT VALUES

        $data = [
            'header' => ['entity_name' => 'DOH-CHD-CAR','section' => $mySD[0],'date' => Carbon::today()->toFormattedDateString()
            ,'prepared_by_name' => strtok( $mySD[4], '||' ),'prepared_by_designation' => ltrim(strstr($mySD[4], '||'), '||')
            ,'recommended_by_name' => strtok( $mySD[2], '||' ),'recommended_by_designation' => ltrim(strstr($mySD[2], '||'), '||')
            ,'dateOfDelivery' => '30 Working Days'
            ],
            'data' => []
        ];


        foreach ($received['app'] as $key => $k) {
            $data['data'][$key] = ['unit' => $k['unit'],'description' => $k['item_name'],'qty' => $k['qty']
            ,'cost' => $k['bid_amount']
            ,'amount' => $k['bid_amount'] * $k['qty']
            ,'supplier' => $k['bidder_name']
            ,'bidder_TIN' => $k['bidder_TIN']
            ,'bidder_address' => $k['bidder_address']
            ,'qtyxbidder_price' => $k['bid_amount'] * $k['qty']
            ,'MOP' => $k['MOP_mode']];
        }

        // Table Header
        // Table Header
        $spreadsheet->getActiveSheet()
            ->setCellValue('A4', $data['header']['entity_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A7', 'Supplier: '. $data['data'][0]['supplier']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A8', 'Address: '. $data['data'][0]['bidder_address']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A9', 'TIN: '. $data['data'][0]['bidder_TIN']);

        $spreadsheet->getActiveSheet()
            ->setCellValue('D7', 'P.O. Number: '. $POno);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D8', 'Date: '.$data['header']['date']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D9', 'Mode of Procurement: '. $data['data'][0]['MOP']);

        $spreadsheet->getActiveSheet()
            ->setCellValue('A13', 'Place of Delivery: '.$data['header']['entity_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D13', 'Delivery Term: ');

        $spreadsheet->getActiveSheet()
            ->setCellValue('A14', 'Date of Delivery: '.$data['header']['dateOfDelivery']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D14', 'Payment Term: ');


        // Table Footer
        // Table Footer

        // $spreadsheet->getActiveSheet()
        //     ->setCellValue('A19',$data['header']['prepared_by_name']);
        // $spreadsheet->getActiveSheet()
        //     ->setCellValue('A20',$data['header']['prepared_by_designation']);
        $totalAmount = 0;
        foreach ($data['data'] as $key => $value) {
            $totalAmount = $totalAmount +$value['qtyxbidder_price'];
        }
        $spreadsheet->getActiveSheet()
            ->setCellValue('C17', $totalAmount);
        $spreadsheet->getActiveSheet()
            ->setCellValue('F17', $totalAmount);

        $spreadsheet->getActiveSheet()
            ->setCellValue('D25',$data['header']['recommended_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D26',$data['header']['recommended_by_designation']);

        // Table insert

        $startRow = 16;
        foreach ($data['data'] as $keyj => $j) {
                $spreadsheet->getActiveSheet()
                    ->setCellValue('B'.$startRow, $j['unit']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('C'.$startRow, $j['description']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('D'.$startRow, $j['qty']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('E'.$startRow, $j['cost']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('F'.$startRow, $j['amount']);
                $spreadsheet->getActiveSheet()
                    ->insertNewRowBefore($startRow+1, 1);
                $startRow++;
            if(count($data['data']) >= $keyj){
                $spreadsheet->getActiveSheet()
                    ->insertNewRowBefore($startRow+1, 4);
            }
        }
        // END INSERT VALUES
        // END INSERT VALUES
        // END INSERT VALUES
        // return $StratSRow.'|'.$CoreSRow.'|'.$SupportSRow;
        $response = response()->streamDownload(function() use ($spreadsheet) {
            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
        });
        $response->setStatusCode(200);
        $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $response->headers->set('Content-Disposition', 'attachment; filename="Purchase Order '.'.xlsx"');
        $response->send();
     }

     public function wfp(Request $request, $year = null, $key, $SD, $flash)
     {

        $mySD = explode("xx",$SD);
        // $user = ['division_id' => $mySD[0], 'section_id' => $mySD[1]];

        if($key !== 'export_key_K2hKr7D0H9u')
            return 'NO_EXPORT_KEY';
        if ($year == null)
            $year = date('Y');
        $input = array(
            'link' => "/api/export/wfp/",
            'apiKey' =>  $key,
            'flash' => $flash,
            'flash2' => $SD,
            'user' => $SD,
        );


        $received = $this->getcurl($input);
        // dd($received);
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load("xls/WFP.xlsx");

        // START INSERT VALUES
        // START INSERT VALUES
        // START INSERT VALUES

        $data = [
            'header' => ['entity_name' => 'DOH-CHD-CAR','section' => $mySD[0],'date' => Carbon::today()->toFormattedDateString()
            ,'prepared_by_name' => strtok( $mySD[4], '||' ),'prepared_by_designation' => ltrim(strstr($mySD[4], '||'), '||')
            ,'recommended_by_name' => strtok( $mySD[2], '||' ),'recommended_by_designation' => ltrim(strstr($mySD[2], '||'), '||')
            ],
            'data' => []
        ];


        foreach ($received['allwfp'] as $key => $k) {
            if ($k['cost'] == 0) {
                $k['fund_source'] = 'SRDS';
            }
            $data['data'][$key] = ['activities' => $k['activities'],'timeframe' => $k['timeframe'],'function' => $k['function']
            ,'q1' => $k['q1']
            ,'q2' => $k['q2']
            ,'q3' => $k['q3']
            ,'q4' => $k['q4']
            ,'cost' => $k['cost']
            ,'function_type' => $k['function_type']
            ,'name_family' => $k['name_family']
            ,'name' => $k['name']
            ,'fund_source' => $k['parent_type']. '|' . $k['source_name']. '|' . $k['type_abbr']];
        }

        // dd($received['allwfp']);
        // Table Header
        // Table Header
        $spreadsheet->getActiveSheet()
            ->setCellValue('A4', 'Name of DOH Unit:'.$data['header']['entity_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A5', 'Budget Line Item:'.' ');

        // Table Footer
        // Table Footer
        $spreadsheet->getActiveSheet()
            ->setCellValue('A19',$data['header']['prepared_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A20',$data['header']['prepared_by_designation']);

        $spreadsheet->getActiveSheet()
            ->setCellValue('D19',$data['header']['recommended_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D20',$data['header']['recommended_by_designation']);

        $spreadsheet->getActiveSheet()
            ->setCellValue('C19',$data['header']['recommended_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('C20',$data['header']['recommended_by_designation']);

        // Table insert
        $StratSRow = 10;
        $CoreSRow = 12;
        $SupportSRow = 14;


        //quick fix//
        //quick fix//
                                            // foreach ($data['data'] as $key => $j) {
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->setCellValue('A'.$StratSRow, $j['function']);
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->setCellValue('B'.$StratSRow, $j['activities']);
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->setCellValue('C'.$StratSRow, $j['timeframe']);
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->setCellValue('D'.$StratSRow, $j['q1']);
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->setCellValue('E'.$StratSRow, $j['q2']);
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->setCellValue('F'.$StratSRow, $j['q3']);
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->setCellValue('G'.$StratSRow, $j['q4']);
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->setCellValue('H'.$StratSRow, $j['cost']);
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->setCellValue('I'.$StratSRow, $j['fund_source']);
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->setCellValue('J'.$StratSRow, $j['name_family'].' ' .$j['name']);
                                            //     $spreadsheet->getActiveSheet()
                                            //         ->insertNewRowBefore($StratSRow+1, 1);

                                            //     $StratSRow++;
                                            // }
        //quick fix//
        //quick fix//

        $a = 0; $b = 0; $c = 0;
        foreach ($received['allwfp'] as $key => $value) {
            if ($value['function_type'] == 'A. Strategic Functions') {
                $a = $a + 1;
            }
            if ($value['function_type'] == 'B. Core Functions') {
                $b = $b + 1;
            }
            if ($value['function_type'] == 'C. Support Functions') {
                $c = $c + 1;
            }
        }
        // return $a.$b.$c;
        if ($a == 0 ) $a = 1;
        if ($b == 0 ) $b = 1;
        if ($c == 0 ) $c = 1;

        $spreadsheet->getActiveSheet()
            ->insertNewRowBefore($StratSRow+1,$a);

        $spreadsheet->getActiveSheet()
            ->insertNewRowBefore($CoreSRow+1+$a, $b);

        $spreadsheet->getActiveSheet()
            ->insertNewRowBefore($SupportSRow+1+$a+$b, $c);

        // foreach ($data['data'] as $key => $j) {
        //     if($j['function_type'] === 'A. Strategic Functions'){
        //         // dd('start');
        //     }
        //     if($j['function_type'] === 'B. Core Functions'){

        //         // dd('core');
        //     }
        //     if($j['function_type'] === 'C. Support Functions'){
        //         // dd('sup');
        //     }
        // }
        // dd( $data['data']);
        foreach ($data['data'] as $keyj => $j) {
            if ($j['function_type'] === 'A. Strategic Functions') {
                $spreadsheet->getActiveSheet()
                    ->setCellValue('A'.$StratSRow, $j['function']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('B'.$StratSRow, $j['activities']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('C'.$StratSRow, $j['timeframe']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('D'.$StratSRow, $j['q1']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('E'.$StratSRow, $j['q2']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('F'.$StratSRow, $j['q3']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('G'.$StratSRow, $j['q4']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('H'.$StratSRow, $j['cost']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('I'.$StratSRow, $j['fund_source']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('J'.$StratSRow, $j['name_family'].' ' .$j['name']);
                $StratSRow++;
            }
            if($j['function_type'] === 'B. Core Functions') {
                $spreadsheet->getActiveSheet()
                    ->setCellValue('A'.($CoreSRow+$a), $j['function']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('B'.($CoreSRow+$a), $j['activities']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('C'.($CoreSRow+$a), $j['timeframe']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('D'.($CoreSRow+$a), $j['q1']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('E'.($CoreSRow+$a), $j['q2']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('F'.($CoreSRow+$a), $j['q3']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('G'.($CoreSRow+$a), $j['q4']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('H'.($CoreSRow+$a), $j['cost']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('I'.($CoreSRow+$a), $j['fund_source']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('J'.($CoreSRow+$a), $j['name_family'].' ' .$j['name']);
                $CoreSRow++;
            }

            if($j['function_type'] === 'C. Support Functions') {
                $spreadsheet->getActiveSheet()
                    ->setCellValue('A'.($SupportSRow + $a + $b), $j['function']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('B'.($SupportSRow + $a + $b), $j['activities']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('C'.($SupportSRow + $a + $b), $j['timeframe']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('D'.($SupportSRow + $a + $b), $j['q1']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('E'.($SupportSRow + $a + $b), $j['q2']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('F'.($SupportSRow + $a + $b), $j['q3']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('G'.($SupportSRow + $a + $b), $j['q4']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('H'.($SupportSRow + $a + $b), $j['cost']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('I'.($SupportSRow + $a + $b), $j['fund_source']);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('J'.($SupportSRow + $a + $b), $j['name_family'].' ' .$j['name']);
                $SupportSRow++;
            }
        }

        // END INSERT VALUES
        // END INSERT VALUES
        // END INSERT VALUES
        // return $StratSRow.'|'.$CoreSRow.'|'.$SupportSRow;
        $response = response()->streamDownload(function() use ($spreadsheet) {
            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
        });
        $response->setStatusCode(200);
        $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $response->headers->set('Content-Disposition', 'attachment; filename="WFP '.'.xlsx"');
        $response->send();
     }

     public function ppmp(Request $request, $year = null, $key, $SD, $flash)
     {
        $mySD = explode("xx",$SD);
        // $user = ['division_id' => $mySD[0], 'section_id' => $mySD[1]];

        if($key !== 'export_key_K2hKr7D0H9u')
            return 'NO_EXPORT_KEY';
        if ($year == null)
            $year = date('Y');
        $input = array(
            'link' => "/api/export/ppmp/",
            'apiKey' =>  $key,
            'flash' => $flash,
            'flash2' => $SD,
            'user' => $SD,
        );

        $received = $this->getcurl($input);
        // dd($received);
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load("xls/PPMPnew.xlsx");

        // START INSERT VALUES
        // START INSERT VALUES
        // START INSERT VALUES

        $data = [
            'header' => ['entity_name' => 'DOH-CHD-CAR','section' => $mySD[0],'date' => Carbon::today()->toFormattedDateString()
            ,'prepared_by_name' => strtok( $mySD[4], '||' ),'prepared_by_designation' => ltrim(strstr($mySD[4], '||'), '||')
            ,'recommended_by_name' => strtok( $mySD[2], '||' ),'recommended_by_designation' => ltrim(strstr($mySD[2], '||'), '||')
            ],
            'data' => []
        ];
        foreach($received['allPPMP'] as $key => $v){
            $data['mops'][$key] = $v['parent_type_abbr'].'-'.$v['source_abbr'].'-'.$v['type_abbr'];
        }
        foreach ($received['allPPMP'] as $key => $k) {
            $fundsss = $k['parent_type'] .' | ' . $k['type'] .' | ' . $k['source_name'];
            $data['data'][$key] = ['general_description' => $k['item_name'],'qty' => $k['qty'],'estimated_budget' => $k['estimated_budget']
            ,'fundsrc' => $fundsss
            ,'mode' => $k['mode']
            ,'m1' => $k['milestones1'],'m2' => $k['milestones2'],'m3' => $k['milestones3'],'m4' => $k['milestones4'],'m5' => $k['milestones5'],'m6' => $k['milestones6']
            ,'m7' => $k['milestones7'],'m8' => $k['milestones8'],'m9' => $k['milestones9'],'m10' => $k['milestones10'],'m11' => $k['milestones11'],'m12' => $k['milestones12']];
        }

        // Table Header
        // Table Header
        $spreadsheet->getActiveSheet()
            ->setCellValue('A5', 'END-USER/UNIT: '.$data['header']['section']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A6', 'Charged To:'. $data['mops'][0]);

        // Table Footer
        // Table Footer
        $spreadsheet->getActiveSheet()
            ->setCellValue('A18',$data['header']['prepared_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A19',$data['header']['prepared_by_designation']);

        $spreadsheet->getActiveSheet()
            ->setCellValue('D18',$data['header']['recommended_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D19',$data['header']['recommended_by_designation']);

        $spreadsheet->getActiveSheet()
            ->setCellValue('K18',$data['header']['recommended_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('K19',$data['header']['recommended_by_designation']);

        // Table insert
        $start = 11;


        //quick fix//
        //quick fix//
        foreach ($data['data'] as $key => $j) {
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('A'.$start, $j['function']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('B'.$start, $j['general_description']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('C'.$start, $j['qty']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('D'.$start, $j['estimated_budget']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('E'.$start, $j['mode']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('F'.$start, $j['m1']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('G'.$start, $j['m2']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('H'.$start, $j['m3']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('I'.$start, $j['m4']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('J'.$start, $j['m5']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('K'.$start, $j['m6']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('L'.$start, $j['m7']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('M'.$start, $j['m8']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('N'.$start, $j['m9']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('O'.$start, $j['m10']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('P'.$start, $j['m11']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('Q'.$start, $j['m12']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('R'.$start, $j['fundsrc']);
            $spreadsheet->getActiveSheet()
                ->insertNewRowBefore($start+1, 1);
            $start++;
        }
        //quick fix//
        //quick fix//

        // END INSERT VALUES
        // END INSERT VALUES
        // END INSERT VALUES
        // return $StratSRow.'|'.$CoreSRow.'|'.$SupportSRow;
        $response = response()->streamDownload(function() use ($spreadsheet) {
            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
        });
        $response->setStatusCode(200);
        $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $response->headers->set('Content-Disposition', 'attachment; filename="PPMP '.'.xlsx"');
        $response->send();
     }


     public function APPOffice(Request $request, $year = null, $key, $SD, $flash)
     {
        $mySD = explode("xx",$SD);
        // $user = ['division_id' => $mySD[0], 'section_id' => $mySD[1]];

        if($key !== 'export_key_K2hKr7D0H9u')
            return 'NO_EXPORT_KEY';
        if ($year == null)
            $year = date('Y');
        $input = array(
            'link' => "/api/export/APPOffice/",
            'apiKey' =>  $key,
            'flash' => $flash,
            'flash2' => $SD,
            'user' => $SD,
        );

        $received = $this->getcurl($input);
        // dd($received['allPPMP']);
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load("xls/PPMPnew.xlsx");

        // START INSERT VALUES
        // START INSERT VALUES
        // START INSERT VALUES

        $data = [
            'header' => ['entity_name' => 'DOH-CHD-CAR','section' => $mySD[0],'date' => Carbon::today()->toFormattedDateString()
            ,'prepared_by_name' => strtok( $mySD[4], '||' ),'prepared_by_designation' => ltrim(strstr($mySD[4], '||'), '||')
            ,'recommended_by_name' => strtok( $mySD[2], '||' ),'recommended_by_designation' => ltrim(strstr($mySD[2], '||'), '||')
            ],
            'data' => []
        ];
        foreach($received['app'] as $key => $v){
            $data['mops'][$key] = $v['parent_type_abbr'].'-'.$v['source_abbr'].'-'.$v['type_abbr'];
        }
        foreach ($received['app'] as $key => $k) {
            $data['data'][$key] = ['general_description' => $k['item_name'],'qty' => $k['qtyTotal'],'estimated_budget' => $k['qtyxabc']
            ,'mode' => $k['mode']
            ,'m1' => $k['milestones1'],'m2' => $k['milestones2'],'m3' => $k['milestones3'],'m4' => $k['milestones4'],'m5' => $k['milestones5'],'m6' => $k['milestones6']
            ,'m7' => $k['milestones7'],'m8' => $k['milestones8'],'m9' => $k['milestones9'],'m10' => $k['milestones10'],'m11' => $k['milestones11'],'m12' => $k['milestones12']];
        }

        // Table Header
        // Table Header
        $spreadsheet->getActiveSheet()
            ->setCellValue('A5', 'END-USER/UNIT: '.$data['header']['section']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A6', 'Charged To:'. $data['mops'][0]);

        // Table Footer
        // Table Footer
        $spreadsheet->getActiveSheet()
            ->setCellValue('A18',$data['header']['prepared_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A19',$data['header']['prepared_by_designation']);

        $spreadsheet->getActiveSheet()
            ->setCellValue('D18',$data['header']['recommended_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D19',$data['header']['recommended_by_designation']);

        $spreadsheet->getActiveSheet()
            ->setCellValue('K18',$data['header']['recommended_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('K19',$data['header']['recommended_by_designation']);

        // Table insert
        $start = 11;


        //quick fix//
        //quick fix//
        foreach ($data['data'] as $key => $j) {
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('A'.$start, $j['function']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('B'.$start, $j['general_description']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('C'.$start, $j['qty']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('D'.$start, $j['estimated_budget']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('E'.$start, $j['mode']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('F'.$start, $j['m1']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('G'.$start, $j['m2']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('H'.$start, $j['m3']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('I'.$start, $j['m4']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('J'.$start, $j['m5']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('K'.$start, $j['m6']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('L'.$start, $j['m7']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('M'.$start, $j['m8']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('N'.$start, $j['m9']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('O'.$start, $j['m10']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('P'.$start, $j['m11']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('Q'.$start, $j['m12']);
            $spreadsheet->getActiveSheet()
                ->insertNewRowBefore($start+1, 1);
            $start++;
        }
        //quick fix//
        //quick fix//

        // END INSERT VALUES
        // END INSERT VALUES
        // END INSERT VALUES
        // return $StratSRow.'|'.$CoreSRow.'|'.$SupportSRow;
        $response = response()->streamDownload(function() use ($spreadsheet) {
            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
        });
        $response->setStatusCode(200);
        $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $response->headers->set('Content-Disposition', 'attachment; filename="PPMP '.'.xlsx"');
        $response->send();
     }

     public function APPOfficeCategory(Request $request, $year = null, $key, $SD, $flash)
     {
        ini_set('max_exection_time', 0);
        $mySD = explode("xx",$SD);
        // $user = ['division_id' => $mySD[0], 'section_id' => $mySD[1]];

        if($key !== 'export_key_K2hKr7D0H9u')
            return 'NO_EXPORT_KEY';
        if ($year == null)
            $year = date('Y');
        $input = array(
            'link' => "/api/export/app/99/2021",
            // 'link' => "/api/export/APPOffice/",
            'apiKey' =>  $key,
            'flash' => $flash,
            'flash2' => $SD,
            'user' => $SD,
        );

        // dd($this->getdebug($input));
        $received = $this->getcurl($input);
        // dd($received['app']);

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load("xls/APP-PER-CATEGORY.xlsx");

        // START INSERT VALUES
        // START INSERT VALUES
        // START INSERT VALUES

        $data = [
            'header' => ['entity_name' => 'DOH-CHD-CAR','section' => $mySD[0],'date' => Carbon::today()->toFormattedDateString()
            ,'prepared_by_name' => strtok( $mySD[4], '||' ),'prepared_by_designation' => ltrim(strstr($mySD[4], '||'), '||')
            ,'recommended_by_name' => strtok( $mySD[2], '||' ),'recommended_by_designation' => ltrim(strstr($mySD[2], '||'), '||')
            ],
            'data' => []
        ];
        foreach($received['app'] as $key => $v){
            $data['mops'][$key] = $v['parent_type_abbr'].'-'.$v['source_abbr'].'-'.$v['type_abbr'];
        }
        foreach ($received['app'] as $key => $k) {
            $data['data'][$key] = ['general_description' => $k['item_name'],'qty' => $k['qtyTotal'],'estimated_budget' => $k['qtyxabc']
            ,'mode' => $k['mode'], 'abc' => $k['abc'], 'itemUnit' => $k['itemUnit']
            , 'FundSource' => $k['Fund Source Type']
            , 'FundType' => $k['type_abbr']
            , 'SourceType' => $k['source_abbr']
            , 'FINALPPMPID' => $k['FINALPPMPID']
            ,'quar1' => ($k['milestones1']+$k['milestones2']+$k['milestones3'])
            ,'quar2' => ($k['milestones4']+$k['milestones5']+$k['milestones6'])
            ,'quar3' => ($k['milestones7']+$k['milestones8']+$k['milestones10'])
            ,'quar4' => ($k['milestones9']+$k['milestones2']+$k['milestones3'])
            ,'divi' => $k['division_abbr'], 'sect' => $k['section_abbr']
            ,'quarAmount1' => $k['ppmpq1'],'quarAmount2' => $k['ppmpq2'],'quarAmount3' => $k['ppmpq3'],'quarAmount4' => $k['ppmpq4']
            ,'m1' => $k['milestones1'],'m2' => $k['milestones2'],'m3' => $k['milestones3'],'m4' => $k['milestones4'],'m5' => $k['milestones5'],'m6' => $k['milestones6']
            ,'m7' => $k['milestones7'],'m8' => $k['milestones8'],'m9' => $k['milestones9'],'m10' => $k['milestones10'],'m11' => $k['milestones11'],'m12' => $k['milestones12']];
        }

        // dd($data['data']);


        // Table Header
        // Table Header
        // $spreadsheet->getActiveSheet()
        //     ->setCellValue('A5', 'END-USER/UNIT: '.$data['header']['section']);
        // $spreadsheet->getActiveSheet()
        //     ->setCellValue('A6', 'Charged To:'. $data['mops'][0]);

        // Table Footer
        // Table Footer
        $spreadsheet->getActiveSheet()
            ->setCellValue('A14',$data['header']['prepared_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('A15',$data['header']['prepared_by_designation']);

        $spreadsheet->getActiveSheet()
            ->setCellValue('D14',$data['header']['recommended_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D15',$data['header']['recommended_by_designation']);

        $spreadsheet->getActiveSheet()
            ->setCellValue('S14',$data['header']['recommended_by_name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('S15',$data['header']['recommended_by_designation']);

        // Table insert
        $start = 3;

        //quick fix//
        //quick fix//
        foreach ($data['data'] as $key => $j) {
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('A'.$start, $j['function']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('A'.$start, $j['general_description']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('B'.$start, $j['itemUnit']);


            $spreadsheet->getActiveSheet()
                ->setCellValue('c'.$start, $j['m1']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('D'.$start, $j['m2']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('E'.$start, $j['m3']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('F'.$start, $j['quar1']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('G'.$start, $j['quarAmount1']);


            $spreadsheet->getActiveSheet()
                ->setCellValue('H'.$start, $j['m4']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('I'.$start, $j['m5']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('J'.$start, $j['m6']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('K'.$start, $j['quar2']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('L'.$start, $j['quarAmount2']);



            $spreadsheet->getActiveSheet()
                ->setCellValue('M'.$start, $j['m7']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('N'.$start, $j['m8']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('O'.$start, $j['m9']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('P'.$start, $j['quar3']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('Q'.$start, $j['quarAmount3']);


            $spreadsheet->getActiveSheet()
                ->setCellValue('R'.$start, $j['m10']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('S'.$start, $j['m11']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('T'.$start, $j['m12']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('U'.$start, $j['quar4']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('V'.$start, $j['quarAmount4']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('W'.$start, $j['qty']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('X'.$start, $j['abc']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('Y'.$start, $j['estimated_budget']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('Z'.$start, $j['FundSource']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('AA'.$start, $j['FundType']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('AB'.$start, $j['SourceType']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('AC'.$start, $j['mode']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('AD'.$start, $j['FINALPPMPID']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('AE'.$start, $j['divi']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('AF'.$start, $j['section_abbr']);





            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('C'.$start, $j['qty']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('D'.$start, $j['estimated_budget']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('E'.$start, $j['mode']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('F'.$start, $j['m1']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('G'.$start, $j['m2']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('H'.$start, $j['m3']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('I'.$start, $j['m4']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('J'.$start, $j['m5']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('K'.$start, $j['m6']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('L'.$start, $j['m7']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('M'.$start, $j['m8']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('N'.$start, $j['m9']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('O'.$start, $j['m10']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('P'.$start, $j['m11']);
            // $spreadsheet->getActiveSheet()
            //     ->setCellValue('Q'.$start, $j['m12']);
            $spreadsheet->getActiveSheet()
                ->insertNewRowBefore($start+1, 1);
            $start++;
        }
        //quick fix//
        //quick fix//

        // END INSERT VALUES
        // END INSERT VALUES
        // END INSERT VALUES
        // return $StratSRow.'|'.$CoreSRow.'|'.$SupportSRow;
        $response = response()->streamDownload(function() use ($spreadsheet) {
            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
        });
        $response->setStatusCode(200);
        $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $response->headers->set('Content-Disposition', 'attachment; filename="OFFICE APP CONSOLIDATED '.'.xlsx"');
        $response->send();
     }

     public function wfpConsolidated(Request $request, $year = null, $key, $SD, $flash)
     {
        $mySD = explode("xx",$SD);

        if($key !== 'export_key_K2hKr7D0H9u')
            return 'NO_EXPORT_KEY';
        if ($year == null)
            $year = date('Y');
        if ($flash == 'all') {
            $input = array(
                'link' => "/api/export/generateMasterConsolidatedWFPReport/",
                'apiKey' =>  $key,
                'flash' => $flash,
                'flash2' => $SD,
                'user' => $SD,
            );
        }else{
            $input = array(
                'link' => "/api/export/generateConsolidatedWFPReport/",
                'apiKey' =>  $key,
                'flash' => $flash,
                'flash2' => $SD,
                'user' => $SD,
            );
        }


        $received = $this->getcurl($input);
        // dd($received);
        // if ($received['']) {
        //     # code...
        // }
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load("xls/ConsolidatedWFP-blank.xlsx");
        // START INSERT VALUES
        // START INSERT VALUES
        // START INSERT VALUES

        $data = [
            'header' => ['entity_name' => 'DOH-CHD-CAR','section' => $mySD[0],'date' => Carbon::today()->toFormattedDateString()
            ,'prepared_by_name' => strtok( $mySD[4], '||' ),'prepared_by_designation' => ltrim(strstr($mySD[4], '||'), '||')
            ,'recommended_by_name' => strtok( $mySD[2], '||' ),'recommended_by_designation' => ltrim(strstr($mySD[2], '||'), '||')
            ],
            'data' => []
        ];


        foreach ($received['allwfp'] as $key => $k) {
            if ($k['cost'] == 0) {
                $k['fund_source'] = 'SRDS';
            }
            $data['data'][$key] = ['activities' => $k['activities'],'timeframe' => $k['timeframe'],'function' => $k['function']
            ,'q1' => $k['q1']
            ,'q2' => $k['q2']
            ,'q3' => $k['q3']
            ,'q4' => $k['q4']
            ,'item' => $k['item']
            ,'cost' => $k['cost']
            ,'function_type' => $k['function_type']
            ,'name_family' => $k['name_family']
            ,'name' => $k['name']
            ,'program_abbr' => $k['program_abbr']
            ,'fund_source' => $k['parent_type']. '|' . $k['source_name']. '|' . $k['type_abbr']];

            $data['sources'][$key] = [
                'fund_source' => $k['parent_type']. '|' . $k['source_name']. '|' . $k['type_abbr'],
                'source_name' => $k['source_name'],
                // 'parent_type' => $k['parent_type'],
            ];

            $data['programs'][$key] = [
                'program_abbr' => $k['program_abbr'],
            ];
        }
        $start = 1;
        $data['sources'] = array_values(array_map("unserialize", array_unique(array_map("serialize", $data['sources']))));
        $data['programs'] = array_map("unserialize", array_unique(array_map("serialize", $data['programs'])));
        // dd(array_map("unserialize", array_unique(array_map("serialize", $data['sources']))));

                                                            // create new sheet
                                                            // create new sheet
                                                            // create new sheet
                                                                        // foreach ($data['sources'] as $key => $value22) {
                                                                        //     $spreadsheet->setActiveSheetIndex($key);
                                                                        //     $spreadsheet->getActiveSheet()->setTitle($value22['source_name']);

                                                                        //     // Table Header
                                                                        //     // Table Header
                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->setCellValue('A4', 'Name of DOH Unit:'.$data['header']['entity_name']);
                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->setCellValue('A5', 'Budget Line Item:'.$value22['fund_source']);

                                                                        //     // Table Footer
                                                                        //     // Table Footer
                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->setCellValue('A19',$data['header']['prepared_by_name']);
                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->setCellValue('A20',$data['header']['prepared_by_designation']);

                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->setCellValue('D19',$data['header']['recommended_by_name']);
                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->setCellValue('D20',$data['header']['recommended_by_designation']);

                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->setCellValue('C19',$data['header']['recommended_by_name']);
                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->setCellValue('C20',$data['header']['recommended_by_designation']);

                                                                        //     // Table insert
                                                                        //     $StratSRow = 10;
                                                                        //     $CoreSRow = 12;
                                                                        //     $SupportSRow = 14;

                                                                        //     $a = 0; $b = 0; $c = 0;
                                                                        //     foreach ($received['allwfp'] as $key => $value) {
                                                                        //         if ($value['function_type'] == 'A. Strategic Functions') {
                                                                        //             $a = $a + 1;
                                                                        //         }
                                                                        //         if ($value['function_type'] == 'B. Core Functions') {
                                                                        //             $b = $b + 1;
                                                                        //         }
                                                                        //         if ($value['function_type'] == 'C. Support Functions') {
                                                                        //             $c = $c + 1;
                                                                        //         }
                                                                        //     }

                                                                        //     if ($a == 0 ) $a = 1;
                                                                        //     if ($b == 0 ) $b = 1;
                                                                        //     if ($c == 0 ) $c = 1;

                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->insertNewRowBefore($StratSRow+1,$a);

                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->insertNewRowBefore($CoreSRow+1+$a, $b);

                                                                        //     $spreadsheet->getActiveSheet()
                                                                        //         ->insertNewRowBefore($SupportSRow+1+$a+$b, $c);

                                                                        //     foreach ($data['data'] as $keyj => $j) {
                                                                        //         if ($j['fund_source'] == $value22['fund_source']) {
                                                                        //             if ($j['function_type'] === 'A. Strategic Functions') {
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('A'.$StratSRow, $j['function']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('B'.$StratSRow, $j['activities']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('C'.$StratSRow, $j['timeframe']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('D'.$StratSRow, $j['q1']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('E'.$StratSRow, $j['q2']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('F'.$StratSRow, $j['q3']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('G'.$StratSRow, $j['q4']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('H'.$StratSRow, $j['cost']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('I'.$StratSRow, $j['fund_source']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('J'.$StratSRow, $j['name_family'].' ' .$j['name']);
                                                                        //                 $StratSRow++;
                                                                        //             }
                                                                        //             if($j['function_type'] === 'B. Core Functions') {
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('A'.($CoreSRow+$a), $j['function']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('B'.($CoreSRow+$a), $j['activities']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('C'.($CoreSRow+$a), $j['timeframe']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('D'.($CoreSRow+$a), $j['q1']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('E'.($CoreSRow+$a), $j['q2']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('F'.($CoreSRow+$a), $j['q3']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('G'.($CoreSRow+$a), $j['q4']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('H'.($CoreSRow+$a), $j['cost']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('I'.($CoreSRow+$a), $j['fund_source']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('J'.($CoreSRow+$a), $j['name_family'].' ' .$j['name']);
                                                                        //                 $CoreSRow++;
                                                                        //             }
                                                                        //             if($j['function_type'] === 'C. Support Functions') {
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('A'.($SupportSRow + $a + $b), $j['function']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('B'.($SupportSRow + $a + $b), $j['activities']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('C'.($SupportSRow + $a + $b), $j['timeframe']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('D'.($SupportSRow + $a + $b), $j['q1']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('E'.($SupportSRow + $a + $b), $j['q2']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('F'.($SupportSRow + $a + $b), $j['q3']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('G'.($SupportSRow + $a + $b), $j['q4']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('H'.($SupportSRow + $a + $b), $j['cost']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('I'.($SupportSRow + $a + $b), $j['fund_source']);
                                                                        //                 $spreadsheet->getActiveSheet()
                                                                        //                     ->setCellValue('J'.($SupportSRow + $a + $b), $j['name_family'].' ' .$j['name']);
                                                                        //                 $SupportSRow++;
                                                                        //             }
                                                                        //         }
                                                                        //     }
                                                                        // }
                                                            // create new sheet
                                                            // create new sheet
                                                            // create new sheet

        // dd($data);

        ///
        ///start
        ///
        ///

        // styles
        $programHead = [
            'font' => [
                'size' => 20,
                'bold'  => true,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startcolor' => [
                    'rgb' => 'd6d6d6',
                ]
            ]
        ];

        $functionHead = array(
            'font'  => array(
                'bold'  => true,
                'color' => array('rgb' => '4287f5'),
                'size'  => 15,
                'name'  => 'Verdana')
        );

        $normal = array(
            'font'  => array(
                'bold'  => false,
                'color' => array('rgb' => '000000'),
                'size'  => 11,
                'name'  => 'Calibri',
            // 'fill' => [
            //     'fillType' => Fill::FILL_SOLID,
            //     'startcolor' => [
            //         'rgb' => 'ffffff',
            //     ]
            // ]
        ));
        $spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(-1);

        foreach ($data['sources'] as $key => $value22) {
            $spreadsheet->setActiveSheetIndex($key);
            $spreadsheet->getActiveSheet()->setTitle(substr($value22['source_name'], 0, 31));

            // Table Header
            // Table Header
            $spreadsheet->getActiveSheet()
                ->setCellValue('A4', 'Name of DOH Unit:'.$data['header']['entity_name']);
            $spreadsheet->getActiveSheet()
                ->setCellValue('A5', 'Budget Line Item:'.$value22['fund_source']);

            // Table Footer
            // Table Footer
                    // $spreadsheet->getActiveSheet()
                    //     ->setCellValue('A19',$data['header']['prepared_by_name']);
                    // $spreadsheet->getActiveSheet()
                    //     ->setCellValue('A20',$data['header']['prepared_by_designation']);

                    // $spreadsheet->getActiveSheet()
                    //     ->setCellValue('D19',$data['header']['recommended_by_name']);
                    // $spreadsheet->getActiveSheet()
                    //     ->setCellValue('D20',$data['header']['recommended_by_designation']);

                    // $spreadsheet->getActiveSheet()
                    //     ->setCellValue('C19',$data['header']['recommended_by_name']);
                    // $spreadsheet->getActiveSheet()
                    //     ->setCellValue('C20',$data['header']['recommended_by_designation']);


            // Table insert
            $masterStart = 9;
            $srt = true;
            $cor = true;
            $sup = true;
            foreach ($data['programs'] as $kpro => $vpro) {
                $spreadsheet->getActiveSheet()
                    ->insertNewRowBefore($masterStart,1);
                // $spreadsheet->getActiveSheet()
                //     ->mergeCells('A'.$masterStart.':K'.$masterStart);
                $spreadsheet->getActiveSheet()
                    ->getStyle('A'.$masterStart.':K'.$masterStart)
                    ->applyFromArray($programHead);
                $spreadsheet->getActiveSheet()
                    ->setCellValue('A'.$masterStart,$vpro['program_abbr'])
                    ->getStyle('A'.($masterStart))
                    // ->getFill()->getStartColor()->setRGB('FF0000');
                    // ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                    ->applyFromArray($programHead);
                foreach ($data['data'] as $keydata => $vdata) {
                    if ($vpro['program_abbr'] == $vdata['program_abbr']) {
                        if ($vdata['fund_source'] == $value22['fund_source']) {
                            if ($vdata['function_type'] === 'A. Strategic Functions') {
                                if ($srt == true) {
                                    $spreadsheet->getActiveSheet()
                                        ->insertNewRowBefore($masterStart+1,2);
                                    // $spreadsheet->getActiveSheet()
                                    //     ->mergeCells('A'.($masterStart+1).':K'.($masterStart+1));
                                    $spreadsheet->getActiveSheet()
                                        ->setCellValue('A'.($masterStart+1), 'STRATEGIC FUNCTIONS')
                                        ->getStyle('A'.($masterStart+1))
                                        ->applyFromArray($functionHead);
                                    $spreadsheet->getActiveSheet()
                                        ->getRowDimension('A'.($masterStart+1))
                                        ->setRowHeight(50);
                                    $srt = false;
                                    $masterStart++;
                                }
                                $spreadsheet->getActiveSheet()
                                    ->insertNewRowBefore($masterStart+1,2);
                                $spreadsheet->getActiveSheet()
                                    ->getRowDimension($masterStart+1)->setRowHeight(-1);
                                $spreadsheet->getActiveSheet()
                                    // ->getStyle('A'.$masterStart.':K'.$masterStart)
                                    ->getStyle($masterStart)
                                    ->applyFromArray($normal);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('A'.($masterStart+1), $vdata['function'])
                                    ->getStyle('A'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('B'.($masterStart+1), $vdata['activities'])
                                    ->getStyle('B'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('C'.($masterStart+1), $vdata['timeframe'])
                                    ->getStyle('C'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('D'.($masterStart+1), $vdata['q1'])
                                    ->getStyle('D'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('E'.($masterStart+1), $vdata['q2'])
                                    ->getStyle('E'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('F'.($masterStart+1), $vdata['q3'])
                                    ->getStyle('F'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('G'.($masterStart+1), $vdata['q4'])
                                    ->getStyle('G'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('H'.($masterStart+1), $vdata['cost'])
                                    ->getStyle('H'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->getStyle('H'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getNumberFormat()
                                    ->setFormatCode(NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('I'.($masterStart+1), $vdata['fund_source'])
                                    ->getStyle('I'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('J'.($masterStart+1), $vdata['item'])
                                    ->getStyle('J'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('K'.($masterStart+1), $vdata['name_family'].' ' .$vdata['name'])
                                    ->getStyle('K'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $masterStart++;
                            }
                            if($vdata['function_type'] === 'B. Core Functions') {
                                if ($cor == true) {
                                    $spreadsheet->getActiveSheet()
                                        ->insertNewRowBefore($masterStart+1,2);
                                    // $spreadsheet->getActiveSheet()
                                    //     ->mergeCells('A'.($masterStart+1).':K'.($masterStart+1));
                                    $spreadsheet->getActiveSheet()
                                        ->setCellValue('A'.($masterStart+1), 'CORE FUNCTIONS')
                                        ->getStyle('A'.($masterStart+1))
                                        ->applyFromArray($functionHead);
                                    $spreadsheet->getActiveSheet()
                                        ->getRowDimension('A'.($masterStart+1))
                                        ->setRowHeight(50);
                                    $cor = false;
                                    $masterStart++;
                                }
                                $spreadsheet->getActiveSheet()
                                    ->insertNewRowBefore(($masterStart+1)+1,2);
                                $spreadsheet->getActiveSheet()
                                    ->getRowDimension($masterStart+1)->setRowHeight(-1);
                                $spreadsheet->getActiveSheet()
                                    ->getStyle($masterStart+1)
                                    ->applyFromArray($normal);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('A'.($masterStart+1), $vdata['function'])
                                    ->getStyle('A'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('B'.($masterStart+1), $vdata['activities'])
                                    ->getStyle('B'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('C'.($masterStart+1), $vdata['timeframe'])
                                    ->getStyle('C'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('D'.($masterStart+1), $vdata['q1'])
                                    ->getStyle('D'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('E'.($masterStart+1), $vdata['q2'])
                                    ->getStyle('E'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('F'.($masterStart+1), $vdata['q3'])
                                    ->getStyle('F'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('G'.($masterStart+1), $vdata['q4'])
                                    ->getStyle('G'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('H'.($masterStart+1), $vdata['cost'])
                                    ->getStyle('H'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->getStyle('H'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getNumberFormat()
                                    ->setFormatCode(NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('I'.($masterStart+1), $vdata['fund_source'])
                                    ->getStyle('I'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                    $spreadsheet->getActiveSheet()
                                    ->setCellValue('J'.($masterStart+1), $vdata['item'])
                                    ->getStyle('J'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('K'.($masterStart+1), $vdata['name_family'].' ' .$vdata['name'])
                                    ->getStyle('K'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $masterStart++;
                            }
                            if($vdata['function_type'] === 'C. Support Functions') {
                                if ($sup == true) {
                                    $spreadsheet->getActiveSheet()
                                        ->insertNewRowBefore($masterStart+1,2);
                                    // $spreadsheet->getActiveSheet()
                                    //     ->mergeCells('A'.($masterStart+1).':K'.($masterStart+1));
                                    $spreadsheet->getActiveSheet()
                                        ->setCellValue('A'.($masterStart+1), 'SUPPORT FUNCTIONS')
                                        ->getStyle('A'.($masterStart+1))
                                        ->applyFromArray($functionHead);
                                    $spreadsheet->getActiveSheet()
                                        ->getRowDimension('A'.($masterStart+1))
                                        ->setRowHeight(50);

                                    $sup = false;
                                    $masterStart++;
                                }
                                $spreadsheet->getActiveSheet()
                                    ->insertNewRowBefore($masterStart+1,2);

                                $spreadsheet->getActiveSheet()
                                    ->getRowDimension($masterStart+1)->setRowHeight(-1);

                                $spreadsheet->getActiveSheet()
                                    ->getStyle($masterStart+1)
                                    ->applyFromArray($normal);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('A'.($masterStart+1), $vdata['function'])
                                    ->getStyle('A'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('B'.($masterStart+1), $vdata['activities'])
                                    ->getStyle('B'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('C'.($masterStart+1), $vdata['timeframe'])
                                    ->getStyle('C'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('D'.($masterStart+1), $vdata['q1'])
                                    ->getStyle('D'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('E'.($masterStart+1), $vdata['q2'])
                                    ->getStyle('E'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('F'.($masterStart+1), $vdata['q3'])
                                    ->getStyle('F'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('G'.($masterStart+1), $vdata['q4'])
                                    ->getStyle('G'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('H'.($masterStart+1), $vdata['cost'])
                                    ->getStyle('H'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->getStyle('H'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getNumberFormat()
                                    ->setFormatCode(NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('I'.($masterStart+1), $vdata['fund_source'])
                                    ->getStyle('I'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                    $spreadsheet->getActiveSheet()
                                    ->setCellValue('J'.($masterStart+1), $vdata['item'])
                                    ->getStyle('J'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('K'.($masterStart+1), $vdata['name_family'].' ' .$vdata['name'])
                                    ->getStyle('K'.($masterStart+1))
                                    ->applyFromArray($normal)
                                    ->getAlignment()->setWrapText(true);
                                $masterStart++;
                            }
                        }
                    }
                }
                $srt = true;
                $cor = true;
                $sup = true;
                $masterStart++;
            }
        }
        ///
        ///
        ///end
        ///




        // END INSERT VALUES
        // END INSERT VALUES
        // END INSERT VALUES
        $response = response()->streamDownload(function() use ($spreadsheet) {
            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
        });
        $filename = date('Y-m-d H:i:s');

        $response->setStatusCode(200);
        $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $response->headers->set('Content-Disposition', 'attachment; filename="WFP '.$filename.'.xlsx"');
        $response->send();
     }



     public function timeCheck($type, $time, $yyear, $mmonth){
        $data = null;
        $sundays = $this->getSundays($yyear,$mmonth);
        $saturdays = $this->getSaturdays($yyear,$mmonth);
       // in: 7:00 - 9:00
       // out: 12:00 12:30
       // in: 12:30 - 1:00
       // out: 16:00 18:00
        switch ($type) {
            case 'morningIn':
                foreach ($time as $key => $v) {
                    if ($v >= '05:00' && $v <= '09:00') {
                        $data = $v;
                        break;
                    }
                }
                break;

            case 'morningOut':
                foreach ($time as $key => $v) {
                    if ($v >= '12:00' && $v <= '12:30') {
                        $data = $v;
                        break;
                    }
                }
                break;
            case 'afternoonIn':
                foreach ($time as $key => $v) {
                    if ($v >= '12:30' && $v <= '13:00') {
                        $data = $v;
                        break;
                    }
                }
                break;

            case 'afternoonOut':
                foreach ($time as $key => $v) {
                    if ($v >= '16:00' && $v <= '22:00') {
                        $data = $v;
                        break;
                    }
                }
                break;

            default:
                # code...
                break;
        }

       return $data;
    }

     public function DTRR(Request $request, $year = null, $key, $flash) {
        if($key !== 'export_key_K2hKr7D0H9u')
            return 'NO_EXPORT_KEY';
        if ($year == null)
            $year = date('Y');
        $flash = str_replace("!!!!", '/', $flash);
        list($firstWord) = explode('/', $flash);

        $monthyear = substr($flash, strpos($flash, "/") + 1);
        $monthyear = substr($monthyear, strpos($monthyear, "/") + 1);
        $yyear = substr($monthyear, strpos($monthyear, "/") + 1);
        $mmonth = strtok($monthyear, '/');

        $input = array(
            // 'link' => "/api/export/pr/".$year,
            'link' => '/'.$flash,
            'apiKey' =>  $key,
            'flash' => $flash,
            'flash2' => '$SD',
            'user' => 'user_d',
        );
        $received = $this->HRGet($input);
        $arr = explode("/", $flash, 2);

        $monthNumber = explode("/", $monthyear, 2);
        $monthName = date('F', mktime(0, 0, 0, $monthNumber[0], 10));
        // dd($monthNumber[0]);

        $HRDetails = array(
            'link' => '/api/v1/employee/'.$arr[0],
            'apiKey' => $this->getAppkey($this->ControllerKey),
            'user' => auth()->user(),
        );
        $personDetails = $this->HRGet($HRDetails);
// dd($this->HRGet($HRDetails));

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load("xls/DTR_NEW.xlsx");

        // START INSERT VALUES
        // START INSERT VALUES
        // START INSERT VALUES

        // $data = [
        //     'header' => ['entity_name' => 'DOH-CHD-CAR','fund_cluster' => $received['allppmp'][0]['source_abbr'].' - '.$received['allppmp'][0]['parent_type_abbr'].' - '.$received['allppmp'][0]['type_abbr'],'section' => $received['user'][6],'PR_number' => $received['allppmp'][0]['prNumber'],'RCC' => 'RCC','date' => Carbon::today()->toFormattedDateString()
        //     ,'requested_by_name' => strtok( $received['user'][4], '||' ),'requested_by_designation' => ltrim(strstr($received['user'][4], '||'), '||')
        //     ,'approved_by_name' => strtok( $received['user'][2], '||' ),'approved_by_designation' => ltrim(strstr($received['user'][2], '||'), '||')
        //     ],
        //     'data' => []

        // ];

        // foreach ($received['allppmp'] as $key => $k) {
        //     $data['data'][$key] = ['unit' => $k['unit'],'item_description' => $k['item_name'].'.                                                           Charge to: '.$k['source_abbr'].' - '.$k['parent_type_abbr'].' - '.$k['type_abbr'],'quantity' => $k['qty'],'unit_cost' => $k['abc'],'total_cost' => $k['estimated_budget']];
        // }
        $dates = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"];
        $groupedDates = [];
        // for ($i=1; $i <= 31; $i++) {
        //     array_push($dates, $i);
        // }
        // dd($dates);
        foreach ($received as $key => $k) {
            foreach ($dates as $key2 => $k2) {
                $timeeee = substr($k['checktime'], strpos($k['checktime'], " ") + 1);
                list($completeDate) = explode(' ', $k['checktime']);

                // list($date) = explode('-', $completeDate);
                // $date = substr($completeDate, 0, strrpos($completeDate, "-"));
                $date = substr($completeDate, strpos($completeDate, "-") + 1);
                $date = substr($date, strpos($date, "-") + 1);
                // dd($date);
                // $date = str_replace("0","", $date);
                // dd($completeDate);
                if ($date == $k2) {
                    // dd($completeDate);
                    $groupedDates[$key2]['completeDate'] = $completeDate;
                    $groupedDates[$key2]['date'] = $date;
                    $groupedDates[$key2]['times'][$key] = $timeeee;
                    // unset($dates[$key2]);
                }else{
                    // $emptyDates[$key2][$k2];
                }
                // dd($k2);
            }
            // $data['data'][$key] = ['unit' => $k['unit'],'item_description' => $k['item_name'].'.                                                           Charge to: '.$k['source_abbr'].' - '.$k['parent_type_abbr'].' - '.$k['type_abbr'],'quantity' => $k['qty'],'unit_cost' => $k['abc'],'total_cost' => $k['estimated_budget']];
        }
        $nullvalue = '00:00:00.000';
        foreach ($groupedDates as $key => $value) {
            $groupedDates[$key]['Final'] = ['min'=> $nullvalue, 'mout'=> $nullvalue, 'ain'=> $nullvalue,'aout'=> $nullvalue];
        }
        // $groupedDates['Final'] = ['min'=> '', 'mout'=> '', 'ain'=> '','aout'=> ''];
        // dd($groupedDates);
        // dd($finalDays);
        foreach ($groupedDates as $key => $value) {
                // dd($value['times']);
            // if (count($value['times']) <= 4) {
                // $finalDays['min' => $this->timeCheck("morningIn",$value['times'])];
                $groupedDates[$key]['Final']['min'] = $this->timeCheck("morningIn",$value['times'], $yyear, $mmonth);
                $groupedDates[$key]['Final']['mout'] = $this->timeCheck("morningOut",$value['times'], $yyear, $mmonth);
                $groupedDates[$key]['Final']['ain'] = $this->timeCheck("afternoonIn",$value['times'], $yyear, $mmonth);
                $groupedDates[$key]['Final']['aout'] = $this->timeCheck("afternoonOut",$value['times'], $yyear, $mmonth);
                // dd($groupedDates);

                // dd(count($value['times']));
            // }
        }
//         dd($this->getSundays($yyear,$mmonth));
$sundays = $this->getSundays($yyear,$mmonth);
$saturdays = $this->getSaturdays($yyear,$mmonth);

        $spreadsheet->getActiveSheet()
            ->setCellValue('e3', $personDetails[0]['name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('m3', $personDetails[0]['name']);
        $spreadsheet->getActiveSheet()
            ->setCellValue('e4', $monthName.', ' . $monthNumber[1]);
        $spreadsheet->getActiveSheet()
            ->setCellValue('m4', $monthName.', ' . $monthNumber[1]);
                    foreach ($groupedDates as $key => $k) {
                        $min = substr($k['Final']['min'], 0, strrpos( $k['Final']['min'], ':') );
                        $ain = substr($k['Final']['ain'], 0, strrpos( $k['Final']['ain'], ':') );
                        $aout = substr($k['Final']['aout'], 0, strrpos( $k['Final']['aout'], ':') );
                        $mout = substr($k['Final']['mout'], 0, strrpos( $k['Final']['mout'], ':') );
                        // if ($min == null && $mout == null && $ain == null && $aout == null) {
                        //     if (in_array($k['date'], $sundays)) {
                        //         $min =  "SUN.";
                        //     }elseif(in_array($k['date'], $saturdays)){
                        //         $min  = "SAT.";
                        //     }else {
                        //         $min = '-';
                        //     }
                        // }

                        // if ($min == null) $min = '-';
                        // if ($mout == null) $mout = '-';
                        // if ($ain == null) $ain = '-';
                        // if ($aout == null) $aout = '-';
                        switch ($k['date']) {
                            case '01':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c9', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d9', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e9', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f9', $aout);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k9', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l9', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m9', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n9', $aout);
                                break;
                            case '02':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c10', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d10', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e10', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f10', $aout);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k10', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l10', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m10', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n10', $aout);
                                break;
                            case '03':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c11', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d11', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e11', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f11', $aout);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k11', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l11', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m11', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n11', $aout);
                                break;
                            case '04':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c12', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d12', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e12', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f12', $aout);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k12', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l12', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m12', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n12', $aout);
                                break;
                            case '05':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c13', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d13', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e13', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f13', $aout);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k13', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l13', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m13', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n13', $aout);
                                break;
                            case '06':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c14', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d14', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e14', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f14', $aout);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k14', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l14', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m14', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n14', $aout);
                                break;
                            case '07':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c15', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d15', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e15', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f15', $aout);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k15', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l15', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m15', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n15', $aout);
                                break;
                            case '08':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c16', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d16', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e16', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f16', $aout);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k16', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l16', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m16', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n16', $aout);
                                break;
                            case '09':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c17', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d17', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e17', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f17', $aout);

                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k17', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l17', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m17', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n17', $aout);
                                break;
                            case '10':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c18', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d18', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e18', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f18', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k18', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l18', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m18', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n18', $aout);
                                break;
                            case '11':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c19', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d19', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e19', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f19', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k19', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l19', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m19', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n19', $aout);
                                break;
                            case '12':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c20', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d20', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e20', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f20', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k20', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l20', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m20', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n20', $aout);
                                break;
                            case '13':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c21', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d21', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e21', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f21', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k21', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l21', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m21', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n21', $aout);
                                break;
                            case '14':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c22', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d22', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e22', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f22', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k22', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l22', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m22', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n22', $aout);
                                break;
                            case '15':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c23', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d23', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e23', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f23', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k23', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l23', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m23', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n23', $aout);
                                break;
                            case '16':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c24', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d24', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e24', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f24', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k24', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l24', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m24', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n24', $aout);
                                break;
                            case '17':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c25', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d25', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e25', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f25', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k25', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l25', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m25', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n25', $aout);
                                break;
                            case '18':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c26', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d26', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e26', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f26', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k26', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l26', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m26', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n26', $aout);
                                break;
                            case '19':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c27', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d27', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e27', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f27', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k27', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l27', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m27', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n27', $aout);
                                break;
                            case '20':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c28', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d28', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e28', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f28', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k28', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l28', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m28', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n28', $aout);
                                break;
                            case '21':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c29', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d29', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e29', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f29', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k29', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l29', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m29', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n29', $aout);
                                break;
                            case '22':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c30', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d30', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e30', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f30', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k30', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l30', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m30', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n30', $aout);
                                break;
                            case '23':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c31', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d31', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e31', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f31', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k31', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l31', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m31', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n31', $aout);
                                break;
                            case '24':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c32', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d32', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e32', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f32', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k32', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l32', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m32', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n32', $aout);
                                break;
                            case '25':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c33', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d33', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e33', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f33', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k33', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l33', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m33', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n33', $aout);
                                break;
                            case '26':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c34', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d34', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e34', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f34', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k34', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l34', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m34', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n34', $aout);
                                break;
                            case '27':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c35', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d35', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e35', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f35', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k35', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l35', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m35', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n35', $aout);
                                break;
                            case '28':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c36', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d36', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e36', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f36', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k36', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l36', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m36', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n36', $aout);
                                break;
                            case '29':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c37', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d37', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e37', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f37', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k37', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l37', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m37', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n37', $aout);
                                break;
                            case '30':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c38', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d38', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e38', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f38', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k38', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l38', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m38', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n38', $aout);
                                break;
                            case '31':
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('c39', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('d39', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('e39', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('f39', $aout);


                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('k39', $min);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('l39', $mout);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('m39', $ain);
                                $spreadsheet->getActiveSheet()
                                    ->setCellValue('n39', $aout);
                                break;

                            default:
                                # code...
                                break;
                        }
                    }

                                                                            // foreach($spreadsheet->getActiveSheet()->getRowDimensions() as $rd) {
                                                                            //     $rd->setRowHeight(-1);
                                                                            // }
                                                                            // // Table Header
                                                                            // // Table Header
                                                                            // $spreadsheet->getActiveSheet()
                                                                            //     ->setCellValue('A6', 'Entity Name: '. $data['header']['entity_name']);
                                                                            // $spreadsheet->getActiveSheet()
                                                                            //     ->setCellValue('D6', 'Fund Cluster: '. $data['header']['fund_cluster']);
                                                                            // $spreadsheet->getActiveSheet()
                                                                            //     ->setCellValue('A7', 'Office/Section: '. $data['header']['section']);
                                                                            // $spreadsheet->getActiveSheet()
                                                                            //     ->setCellValue('C7', 'PR Number: '. $data['header']['PR_number']);
                                                                            // $spreadsheet->getActiveSheet()
                                                                            //     ->setCellValue('C8', 'Responsibility Center Code: '. $data['header']['RCC']);
                                                                            // $spreadsheet->getActiveSheet()
                                                                            //     ->setCellValue('E7', 'Date: '. $data['header']['date']);

                                                                            // // Table Footer
                                                                            // // Table Footer
                                                                            // // $spreadsheet->getActiveSheet()
                                                                            // //     ->setCellValue('B18',$data['header']['requested_by_name']);
                                                                            // // $spreadsheet->getActiveSheet()
                                                                            // //     ->setCellValue('B19',$data['header']['requested_by_designation']);
                                                                            // // $spreadsheet->getActiveSheet()
                                                                            // //     ->setCellValue('D18',$data['header']['approved_by_name']);
                                                                            // // $spreadsheet->getActiveSheet()
                                                                            // //     ->setCellValue('D19',$data['header']['approved_by_designation']);

                                                                            // // Table insert
                                                                            // $startRow = 11;
                                                                            // foreach ($data['data'] as $keyj => $j) {

                                                                            //     $spreadsheet->getActiveSheet()->insertNewRowBefore($startRow, 1);

                                                                            //     // Unit
                                                                            //     $spreadsheet->getActiveSheet()
                                                                            //         ->setCellValue('B'.$startRow, $j['unit']);//A11
                                                                            //     // Item Description
                                                                            //     $spreadsheet->getActiveSheet()
                                                                            //         ->setCellValue('C'.$startRow, $j['item_description']);
                                                                            //     // Quantity
                                                                            //     $spreadsheet->getActiveSheet()
                                                                            //         ->setCellValue('D'.$startRow, $j['quantity']);
                                                                            //     // Unit Cost
                                                                            //     $spreadsheet->getActiveSheet()
                                                                            //         ->setCellValue('E'.$startRow, $j['unit_cost']);
                                                                            //     // Total Cost
                                                                            //     $spreadsheet->getActiveSheet()
                                                                            //         ->setCellValue('F'.$startRow, $j['total_cost']);
                                                                            //     $spreadsheet->getActiveSheet()->getRowDimension(1)->setRowHeight(-1);
                                                                            //     $startRow++;
                                                                            // }
        // }
        // END INSERT VALUES
        // END INSERT VALUES
        // END INSERT VALUES



        // PRINT DTR IN EXCEL
        // PRINT DTR IN EXCEL
        // PRINT DTR IN EXCEL
        // $response = response()->streamDownload(function() use ($spreadsheet) {
        //     $writer = new Xlsx($spreadsheet);
        //     $writer->save('php://output');
        // });
        // $response->setStatusCode(200);
        // $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        // $response->headers->set('Content-Disposition', 'attachment; filename="DTR.xlsx"');
        // $response->send();


        // PRINT DTR IN PDF
        // PRINT DTR IN PDF
        // PRINT DTR IN PDF
        // $response1 = response()->streamDownload(function() use ($spreadsheet) {
        //     $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Mpdf');
        //     $writer->save('php://output');
        // });
        // $response1->setStatusCode(200);
        // $response1->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        // $response1->headers->set('Content-Disposition', 'attachment; filename="DTR.pdf"');
        // // $response1->headers->set('Content-Disposition', 'attachment; filename="DTR.xlsx"');
        // $response1->send();


        // test
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="myfile.pdf"');
        header('Cache-Control: max-age=0');
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Mpdf');
        $writer->save('php://output');
        dd('sdfsdf');
                        // header('Content-Type: application/vnd.ms-excel');
                        // header('Content-Disposition: attachment;filename="myfile.xls"');
                        // header('Cache-Control: max-age=0');
                        // $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Mpdf');
                        // $writer->save('php://output');
        // $pdf = PDF::loadView ($response, $flash);
        // return $pdf->download('invoice.pdf');
     }

     function getSundays($y,$m){
        $date = "$y-$m-01";
        $first_day = date('N',strtotime($date));
        $first_day = 7 - $first_day + 1;
        $last_day =  date('t',strtotime($date));
        $days = array();
        for($i=$first_day; $i<=$last_day; $i=$i+7 ){
            $days[] = $i;
        }
        return  $days;
    }
     function getSaturdays($y,$m){
        $date = "$y-$m-01";
        $first_day = date('N',strtotime($date));
        $first_day = 7 - $first_day;
        $last_day =  date('t',strtotime($date));
        $days = array();
        for($i=$first_day; $i<=$last_day; $i=$i+7 ){
            $days[] = $i;
        }
        return  $days;
    }

}
