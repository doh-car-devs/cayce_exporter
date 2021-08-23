<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Traits\APIAccess;
use App\Http\Controllers\Traits\CurlNow;
use App\Http\Controllers\Controller;
use Illuminate\Http\Request;

class WordController extends Controller
{
    use CurlNow, APIAccess;

    function postQualEvalReport($year = null, $key)
    {
        if($key !== 'export_key_K2hKr7D0H9u')
            return 'NO_EXPORT_KEY';

        if ($year == null)
            $year = date('Y');

        $input = array(
            'link' => "/api/twg/export/postqualevalreport/".$year,
            'apiKey' =>  $key,
            'user' => 'user_d',
        );

        $data = $this->getcurl($input);
        $itemKeys = array_column($data['item'], 'qtyxabc');
        array_multisort($itemKeys, SORT_ASC, $data['item']);
        // dd($data);
        $wordTest = new \PhpOffice\PhpWord\PhpWord();

        // STYLES
            $wordTest->addFontStyle(
                'fontStyleP',array('name' => 'Times New Roman', 'size' => 12, 'color' => '000000', 'bold' => false)
            );
            $wordTest->addParagraphStyle(
                'paragraphStyleP',array('align' => 'lowKashida')
            );

            $wordTest->addFontStyle(
                'fontStyleH',array('name' => 'Times New Roman', 'size' => 12, 'color' => '000000', 'bold' => true)
            );

            $TfontStyle = array('bold'=>true, 'italic'=> false, 'size'=>11, 'name' => 'Times New Roman', 'afterSpacing' => 0, 'Spacing'=> 0, 'cellMargin'=>0, 'spaceAfter' => 0);
            $fontStyle = array('italic'=> true, 'size'=>11, 'name'=>'Times New Roman','afterSpacing' => 0, 'Spacing'=> 0, 'cellMargin'=>0 , 'spaceAfter' => 0);
            $cfontStyle = array('allCaps'=>false,'italic'=> false, 'size'=>11, 'name' => 'Times New Roman','afterSpacing' => 0, 'Spacing'=> 0, 'cellMargin'=>0, 'spaceAfter' => 0);
            $styleCell = array('borderTopSize'=>1 ,'borderTopColor' =>'black','borderLeftSize'=>1,'borderLeftColor' =>'black','borderRightSize'=>1,'borderRightColor'=>'black','borderBottomSize' =>1,'borderBottomColor'=>'black');
            $styleTransCell = array('borderTopSize'=>1 ,'borderTopColor' =>'white','borderLeftSize'=>1,'borderLeftColor' =>'white','borderRightSize'=>1,'borderRightColor'=>'white','borderBottomSize' =>1,'borderBottomColor'=>'white');
        // STYLES
        $bidsSection = $wordTest->addSection();

        $bidsstyle = [
            'borderSize' => 1,
        ];
        $wordTest->addTableStyle('bidstable', $bidsstyle);
        $bidstable = $bidsSection->addTable('bidstable');

        // for ($i=0; $i < 10; $i++) {
        foreach ($data['app']['items'] as $i => $value) {
            $bidrow = $bidstable->addRow();
                $bidrow->addCell(null, ['gridSpan' => 8,'borderTopSize'=>1 ,'borderTopColor' =>'white','borderLeftSize'=>1,'borderLeftColor' =>'white','borderRightSize'=>1,'borderRightColor'=>'white','borderBottomSize' =>1,'borderBottomColor'=>'white'])
                    ->addText(null);
                $bidrow->addCell(null, ['gridSpan' => 8,'borderTopSize'=>1 ,'borderTopColor' =>'white','borderLeftSize'=>1,'borderLeftColor' =>'white','borderRightSize'=>1,'borderRightColor'=>'white','borderBottomSize' =>1,'borderBottomColor'=>'white'])
                    ->addText(null);
            $bidrow = $bidstable->addRow();
                $bidrow->addCell(null, $styleTransCell)
                    ->addText('Item No.    '.($i+1), null, ['align' => 'left']);
                $bidrow->addCell(null,['gridSpan' => 7,'borderTopSize'=>1 ,'borderTopColor' =>'white','borderLeftSize'=>1,'borderLeftColor' =>'white','borderRightSize'=>1,'borderRightColor'=>'white','borderBottomSize' =>1,'borderBottomColor'=>'white'])
                    ->addText($value['item'], null, ['align' => 'left']);

            $bidrow = $bidstable->addRow();
                $bidrow->addCell(null, $styleTransCell)
                    ->addText('ABC:', null, ['align' => 'left']);
                $bidrow->addCell(null, ['gridSpan' => 7,'borderTopSize'=>1 ,'borderTopColor' =>'white','borderLeftSize'=>1,'borderLeftColor' =>'white','borderRightSize'=>1,'borderRightColor'=>'white','borderBottomSize' =>1,'borderBottomColor'=>'white'])
                    ->addText("{item_price_total}"."@"."{item_price}"."per"."{item_unit}", null, ['align' => 'left']);

            $bidrow = $bidstable->addRow();
                $bidrow->addCell(1500)
                    ->addText('Bidder No.',['align' => 'left', 'vMerge' => 'restart']);
                $bidrow->addCell(null)
                    ->addText('Name of Bidder Identification', null, ['align' => 'center']);
                $bidrow->addCell(null, ['gridSpan' => 5])
                    ->addText('Name of Bidder Identification', null, ['align' => 'center']);
                $bidrow->addCell(null)
                    ->addText('Rank', null, ['align' => 'center']);

            // for ($j=0; $j < 5; $j++) {
            $k=1;
            foreach ($data['item'] as $j => $valuej) {
                if ($valuej['item_id'] == $value['id']) {
                    $k = $k + 1;
                    $bidrow = $bidstable->addRow();
                        $bidrow->addCell(1500,$styleCell)
                            ->addText($k,$TfontStyle);
                        $bidrow->addCell(4000,$styleCell)
                            ->addText($valuej['bidder_name'],$cfontStyle);
                        $bidrow->addCell(1500,$styleCell)
                            ->addText('₱'.number_format($valuej['qtyxabc'], 2, '.', ','),$cfontStyle, ['align' => 'center']);
                        $bidrow->addCell(500,$styleCell)
                            ->addText('@',$cfontStyle, ['align' => 'center']);
                        $bidrow->addCell(1000,$styleCell)
                            ->addText('₱'.number_format($valuej['bid_amount'], 2, '.', ','),$cfontStyle, ['align' => 'center']);
                        $bidrow->addCell(500,$styleCell)
                            ->addText('per',$cfontStyle, ['align' => 'center']);
                        $bidrow->addCell(500,$styleCell)
                            ->addText($valuej['unit'],$cfontStyle, ['align' => 'center']);
                        $bidrow->addCell(1500,$styleCell)
                            ->addText($j+1,$TfontStyle, ['align' => 'center']);
                }
            }
        }

        $objectWriter = \PhpOffice\PhpWord\IOFactory::createWriter($wordTest, 'Word2007');

        try {
            $objectWriter->save(storage_path('Post Qualification Evaluation Summary Report.docx'));
        } catch (Exception $e) {
        }

        return response()->download(storage_path('Post Qualification Evaluation Summary Report.docx'));
    }

    function index()
    {
        $wordTest = new \PhpOffice\PhpWord\PhpWord();

        // STYLES
            $wordTest->addFontStyle(
                'fontStyleP',array('name' => 'Times New Roman', 'size' => 12, 'color' => '000000', 'bold' => false)
            );
            $wordTest->addParagraphStyle(
                'paragraphStyleP',array('align' => 'lowKashida')
            );

            $wordTest->addFontStyle(
                'fontStyleH',array('name' => 'Times New Roman', 'size' => 12, 'color' => '000000', 'bold' => true)
            );


            $TfontStyle = array('bold'=>true, 'italic'=> false, 'size'=>11, 'name' => 'Times New Roman', 'afterSpacing' => 0, 'Spacing'=> 0, 'cellMargin'=>0, 'spaceAfter' => 0);
            $fontStyle = array('italic'=> true, 'size'=>11, 'name'=>'Times New Roman','afterSpacing' => 0, 'Spacing'=> 0, 'cellMargin'=>0 , 'spaceAfter' => 0);
            $cfontStyle = array('allCaps'=>false,'italic'=> false, 'size'=>11, 'name' => 'Times New Roman','afterSpacing' => 0, 'Spacing'=> 0, 'cellMargin'=>0, 'spaceAfter' => 0);
            $styleCell = array('borderTopSize'=>1 ,'borderTopColor' =>'black','borderLeftSize'=>1,'borderLeftColor' =>'black','borderRightSize'=>1,'borderRightColor'=>'black','borderBottomSize' =>1,'borderBottomColor'=>'black');
            $styleTransCell = array('borderTopSize'=>1 ,'borderTopColor' =>'white','borderLeftSize'=>1,'borderLeftColor' =>'white','borderRightSize'=>1,'borderRightColor'=>'white','borderBottomSize' =>1,'borderBottomColor'=>'white');

            // rowhelper
            $cellRowSpan = array('vMerge' => 'restart');
            $cellRowContinue = array('vMerge' => 'continue');
            $cellColSpan = array('gridSpan' => 2);
        // STYLES

        $bidsSection = $wordTest->addSection();

        $bidsstyle = [
            'borderSize' => 1,
        ];
        $wordTest->addTableStyle('bidstable', $bidsstyle);
        $bidstable = $bidsSection->addTable('bidstable');


        // $bidrow = $bidstable->addRow();
        //     $row->addCell(null, ['gridSpan' => 4])
        //         ->addText('1', null, ['align' => 'center']);
        for ($i=0; $i < 10; $i++) {
            $bidrow = $bidstable->addRow();
                $bidrow->addCell(null, ['gridSpan' => 8,'borderTopSize'=>1 ,'borderTopColor' =>'white','borderLeftSize'=>1,'borderLeftColor' =>'white','borderRightSize'=>1,'borderRightColor'=>'white','borderBottomSize' =>1,'borderBottomColor'=>'white'])
                    ->addText(null);
                $bidrow->addCell(null, ['gridSpan' => 8,'borderTopSize'=>1 ,'borderTopColor' =>'white','borderLeftSize'=>1,'borderLeftColor' =>'white','borderRightSize'=>1,'borderRightColor'=>'white','borderBottomSize' =>1,'borderBottomColor'=>'white'])
                    ->addText(null);
            $bidrow = $bidstable->addRow();
                $bidrow->addCell(null, $styleTransCell)
                    ->addText('Item No.', null, ['align' => 'left']);
                $bidrow->addCell(null,['gridSpan' => 7,'borderTopSize'=>1 ,'borderTopColor' =>'white','borderLeftSize'=>1,'borderLeftColor' =>'white','borderRightSize'=>1,'borderRightColor'=>'white','borderBottomSize' =>1,'borderBottomColor'=>'white'])
                    ->addText("{item_name}", null, ['align' => 'left']);

            $bidrow = $bidstable->addRow();
                $bidrow->addCell(null, $styleTransCell)
                    ->addText('ABC:', null, ['align' => 'left']);
                $bidrow->addCell(null, ['gridSpan' => 7,'borderTopSize'=>1 ,'borderTopColor' =>'white','borderLeftSize'=>1,'borderLeftColor' =>'white','borderRightSize'=>1,'borderRightColor'=>'white','borderBottomSize' =>1,'borderBottomColor'=>'white'])
                    ->addText("{item_price_total}"."@"."{item_price}"."per"."{item_unit}", null, ['align' => 'left']);

            $bidrow = $bidstable->addRow();
                $bidrow->addCell(1500)
                    ->addText('Bidder No.',['align' => 'left', 'vMerge' => 'restart']);
                $bidrow->addCell(null)
                    ->addText('Name of Bidder Identification', null, ['align' => 'center']);
                $bidrow->addCell(null, ['gridSpan' => 5])
                    ->addText('Name of Bidder Identification', null, ['align' => 'center']);
                $bidrow->addCell(null)
                    ->addText('Rank', null, ['align' => 'center']);

            for ($j=0; $j < 5; $j++) {
                $bidrow = $bidstable->addRow();
                    $bidrow->addCell(1500,$styleCell)
                        ->addText($j+1,$TfontStyle);
                    $bidrow->addCell(4000,$styleCell)
                        ->addText('Name of bidder'.($j+1),$cfontStyle);
                    $bidrow->addCell(1500,$styleCell, ['align' => 'center'])
                        ->addText('{bid_price_total}',$cfontStyle);
                    $bidrow->addCell(500,$styleCell, ['align' => 'center'])
                        ->addText('@',$cfontStyle);
                    $bidrow->addCell(1000,$styleCell, ['align' => 'center'])
                        ->addText('{bid_price}',$cfontStyle);
                    $bidrow->addCell(500,$styleCell, ['align' => 'center'])
                        ->addText('per',$cfontStyle);
                    $bidrow->addCell(500,$styleCell, ['align' => 'center'])
                        ->addText('{unit}',$cfontStyle);
                    $bidrow->addCell(1500,$styleCell, ['align' => 'center'])
                        ->addText($j+1,$TfontStyle);
            }
        }
        ////////////////
        ////////////////
        ////////////////
        ////////////////

        $section = $wordTest->addSection(
            [
                'orientation' => 'portrait',
                'marginTop' => 200,
                'marginLeft' => 200,
                'marginRight' => 200,
                'marginBottom' => 200,
            ]
        );
        $tableStyle = [
            'width' => 100,
            'borderSize' => 1,
            'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER,
        ];
        $wordTest->addTableStyle('myTable', $tableStyle);
        $table = $section->addTable('myTable');

        $row = $table->addRow();

            $row->addCell(4200, ['valign' => 'center', 'vMerge' => 'restart'])
                ->addText('A', null, ['align' => 'center']);

            $row->addCell(null, ['gridSpan' => 4, 'vMerge' => 'restart'])
                ->addText('B', null, ['align' => 'center']);

            $row->addCell(null, ['gridSpan' => 4])
                ->addText('1', null, ['align' => 'center']);

        $row = $table->addRow();
            $row->addCell(null, ['vMerge' => 'continue']);
            $row->addCell(null, ['vMerge' => 'continue','gridSpan' => 3]);
            $row->addCell(null, ['vMerge' => 'continue']);
            $row->addCell(null, ['gridSpan' => 2])->addText('2');
            $row->addCell(null, ['gridSpan' => 2]);
        $row = $table->addRow();
            $row->addCell(null, ['vMerge' => 'continue']);
            $row->addCell()->addText('C');
            $row->addCell()->addText('D');
            $row->addCell()->addText('3');
            // $row->addCell();
            // $row->addCell();
            // $row->addCell();
        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////

        $mytable = $wordTest->addSection();
        // $space = $wordTest->addSection();
        // $wordTest->addTableStyle('myTable', $tableStyle);
        $table = $mytable->addTable('myOwnTableStyle',array('borderSize' => 1, 'borderColor' => '999999', 'afterSpacing' => 0, 'Spacing'=> 0, 'cellMargin'=>0  ,'spaceAfter' => 0));
        for ($i=0; $i < 5; $i++) {
            $table->addRow();
            $table->addCell(null,$styleTransCell);
            $table->addCell(null,$styleTransCell);

            $table->addRow();
            $table->addCell(1500,$styleTransCell)->addText('Item No.',$TfontStyle);
            $table->addCell(500,$styleTransCell)->addText($i+1,$TfontStyle);
            $table->addCell(4000,$styleTransCell)->addText("{item_name}",$cfontStyle);
            $table->addCell(5500,$styleTransCell, $cellRowContinue);
            $table->addCell(null,$styleTransCell);


            $table->addRow();
            $table->addCell(1500,$styleTransCell)->addText('ABC',$TfontStyle);
            $table->addCell(500,$cellRowSpan);
            $table->addCell(4000,$styleTransCell)->addText("{item_price_total}"."@"."{item_price}"."per"."{item_unit}",$cfontStyle);
            $table->addCell(4000,$styleTransCell);
            $table->addCell(1500,$styleTransCell);


            $table->addRow();
            $table->addCell(1500,$styleCell)->addText('Bidder No.',$TfontStyle);
            $table->addCell(null,$styleCell);
            $table->addCell(4000,$styleCell)->addText('Name of Bidder Identification',$TfontStyle);
            $table->addCell(4000,$styleCell, ['gridSpan' => 4])->addText('Total Bid as Read Amount (php)',$TfontStyle);
            $table->addCell(1500,$styleCell)->addText('Rank',$TfontStyle);
            $j =1;
            // for ($j=0; $j < 5; $j++) {
                $table->addRow();
                $table->addCell(1500,$styleCell)->addText($j+1,$TfontStyle);
                $table->addCell(null,$styleCell);
                $table->addCell(4000,$styleCell)->addText('Name of bidder'.($j+1),$TfontStyle);
                // $table->addCell(4000,$styleCell)->addText('Total Bid as Read Amount (php)',$TfontStyle);
                $table->addCell(1500,$styleCell)->addText('{bid_price_total}',$TfontStyle);
                $table->addCell(500,$styleCell)->addText('@',$TfontStyle);
                $table->addCell(1000,$styleCell)->addText('{bid_price}',$TfontStyle);
                $table->addCell(500,$styleCell)->addText('per',$TfontStyle);
                $table->addCell(500,$styleCell)->addText('{unit}',$TfontStyle);
                $table->addCell(1500,$styleCell)->addText($j+1,$TfontStyle);
            // }
        }
        // $table->addCell(6000,$styleCell)->addText($news['CompanyDetails']['CompanyName'],$cfontStyle);
        // $table->addCell(2000,$styleCell)->addText($news['CompanyDetails']['Tel'],$fontStyle);

        // table2
        // $table2 = $mytable->addTable('mytable2',array('borderSize' => 1, 'borderColor' => '999999', 'afterSpacing' => 0, 'Spacing'=> 0, 'cellMargin'=>0  ));
        // $table2->addRow();
        // $table2->addCell(2500,$styleCell)->addText('Company Name',$TfontStyle);
        // $table2->addCell(6000,$styleCell)->addText('test',$cfontStyle);
        // $table2->addCell(1500,$styleCell)->addText('Tel',$TfontStyle);
        // $table2->addCell(2000,$styleCell)->addText('test',$fontStyle);

        // $headerSection = $wordTest->createSection();
        // $header = $headerSection->createHeader();
        // $headerText = "Department of Health Cordillera Administrative Regional Office";
        // $header->addText(htmlspecialchars($headerText));

        $aTitl = '1.0	PROJECT IDENTIFICATION';
        $aDesc = "The Department of Health-Cordillera Administrative Regional Office, through the General Appropriations Act intends to apply the sum of Six Million Sixty Two Thousand Three Hundred Four (Php 6,062,304.00) being the Approved Budget for the Contract (ABC) for the Procurement of Information Technology Supplies & equipment and Paper Materials & Products for 2017 Requirements under Invitation to Bid No. 2017-005 (Regular) published on conducted through open competitive bidding procedures using non-discretionary pass/fail criteria as specified in the Revised Implementing Rules and Regulations Act (R.A. 9184).";


        $main = $wordTest->addSection();

        $main->addText('POST QUALIFICATION EVALUATION SUMMARY REPORT', array('bold' =>true, 'size' => 12, 'align' => 'center'));
        $main->addText($aTitl, array('bold' =>true, 'size' => 12));
        $main->addText(htmlspecialchars($aDesc), 'fontStyleP','paragraphStyleP');


        $objectWriter = \PhpOffice\PhpWord\IOFactory::createWriter($wordTest, 'Word2007');

        try {
            $objectWriter->save(storage_path('Exported.docx'));
        } catch (Exception $e) {
        }

        return response()->download(storage_path('Exported.docx'));

    }

    function sample()
    {
        $wordTest = new \PhpOffice\PhpWord\PhpWord();

        $docx = $wordTest->addSection();

        $docx->addText('Quisque ullamcorper, dolor eget eleifend consequat,
        justo nunc ultricies quam, sed ullamcorper lectus urna ac justo.
        Phasellus sed libero ut dui hendrerit tempus. Mauris tincidunt
        laoreet sapien, feugiat viverra justo dictum eu. Cras at eros ac
        urna accumsan varius. Vestibulum cursus gravida sollicitudin.
        Donec vestibulum lectus id sem malesuada volutpat. Praesent et ipsum orci.
        Sed rutrum eros id erat fermentum in auctor velit auctor.
        Nam bibendum rutrum augue non pellentesque. Donec in mauris dui,
        non sagittis dui. Phasellus quam leo, ultricies sit amet cursus nec,
        elementum at est. Proin blandit volutpat odio ac dignissim.
        In at lacus dui, sed scelerisque ante. Aliquam tempor,
        metus sed malesuada vehicula, neque massa malesuada dolor,
        vel semper massa ante eu nibh.');

        // create a simple Word fragment to insert into the table
        $textFragment = $wordTest->WordFragment($docx);
        // $textFragment = new WordFragment($docx);
        $text = array();
        $text[] = array('text' => 'Fit text and ');
        $text[] = array('text' => 'Word fragment', 'bold' => true);
        $textFragment->addText($text);

        // establish some row properties for the first row
        $trProperties = array();
        $trProperties[0] = array(
            'minHeight' => 1000,
            'tableHeader' => true,
        );

        $col_1_1 = array(
            'rowspan' => 4,
            'value' => '1_1',
            'backgroundColor' => 'cccccc',
            'borderColor' => 'b70000',
            'border' => 'single',
            'borderTopColor' => '0000FF',
            'borderWidth' => 16,
            'cellMargin' => 200,
        );
        $col_2_2 = array(
            'rowspan' => 2,
            'colspan' => 2,
            'width' => 200,
            'value' => $textFragment,
            'backgroundColor' => 'ffff66',
            'borderColor' => 'b70000',
            'border' => 'single',
            'cellMargin' => 200,
            'fitText' => 'on',
            'vAlign' => 'bottom',
        );
        $col_2_4 = array(
            'rowspan' => 3,
            'value' => 'Some rotated text',
            'backgroundColor' => 'eeeeee',
            'borderColor' => 'b70000',
            'border' => 'single',
            'borderWidth' => 16,
            'textDirection' => 'tbRl',
        );

        //set the global table properties
        $options = array(
            'columnWidths' => array(400,1400,400,400,400),
            'border' => 'single',
            'borderWidth' => 4,
            'borderColor' => 'cccccc',
            'borderSettings' => 'inside',
            'float' => array(
                'align' => 'right',
                'textMargin_top' => 300,
                'textMargin_right' => 400,
                'textMargin_bottom' => 300,
                'textMargin_left' => 400,
            ),
            'tableWidth' => array('type' => 'pct', 'value' => 70),
        );

        $values = array(
            array($col_1_1, '1_2', '1_3', '1_4', '1_5'),
            array($col_2_2, $col_2_4, '2_5'),
            array('3_5'),
            array('4_2', '4_3', '4_5'),
        );

        $docx->addTable($values, $options, $trProperties);

        $docx->addText('In pretium neque vitae sem posuere volutpat.
        Class aptent taciti sociosqu ad litora torquent per conubia nostra,
        per inceptos himenaeos. Quisque eget ultricies ipsum. Cras vitae suscipit
        erat. Nullam fermentum risus sed urna fermentum placerat laoreet arcu lobortis.
        Integer nisl erat, vehicula eget posuere id, mollis fermentum mi.
        Phasellus quis nulla orci. Suspendisse malesuada lectus et turpis facilisis
        id imperdiet tellus luctus. In hac habitasse platea dictumst. Proin a mattis turpis.
        Aliquam sit amet velit a lacus hendrerit bibendum. Mauris euismod dictum augue eget condimentum.');

        $docx->createDocx('output');
    }
}
