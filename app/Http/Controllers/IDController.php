<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpWord\TemplateProcessor;

class IDController extends Controller
{
    public function generate()
    {
        // $wordTest = new \PhpOffice\PhpWord\PhpWord();
 
        // // $newSection = $wordTest->addSection();
     
        // // $desc1 = "The Portfolio details is a very useful feature of the web page. You can establish your archived details and the works to the entire web community. It was outlined to bring in extra clients, get you selected based on this details.";
     
        // // $newSection->addText($desc1, array('name' => 'Tahoma', 'size' => 15, 'color' => 'red'));
     
        // // $objectWriter = \PhpOffice\PhpWord\IOFactory::createWriter($wordTest, 'Word2007');
        // // try {
        // //     $objectWriter->save(storage_path('TestWordFile.docx'));
        // // } catch (Exception $e) {
        // // }
     
        // // return response()->download(storage_path('TestWordFile.docx'));
        // return asset('storage/template.docx');
        // $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor(asset('storage/template.docx'));
        // $templateProcessor->setValue('firstname', 'John');
        // $templateProcessor->setValue('lastname', 'Doe');
        // // $templateProcessor->setImageValue('foto', array('path' => 'dummy_foto.jpg', 'width' => 100, 'height' => 100, 'ratio' => true));
        // $templateProcessor->saveAs('result.docx');

		// return( asset('storage/images/1.png'));
		$pic1 = asset('storage/images/1.png');
		$values = [
		    ['user_id' => 1, 'name' => 'Rolf Cayce P. Dy', 'position' => 'Coputer Programmer I', 'nname' => 'Cayce'                      , 'division' => 'RLED'],
		    ['user_id' => 2, 'name' => 'Keanu F. Acierto', 'position' => 'Disease Surveillance Officer ', 'nname' => 'Keanu'             , 'division' => 'MSD'],
		    ['user_id' => 3, 'name' => 'Kelvin K. Valdez', 'position' => 'Development Management Officer IV', 'nname' => 'Kelvs'         , 'division' => 'LHSD'],
		    ['user_id' => 4, 'name' => 'Hobel C. Pacatiw', 'position' => 'Dengvaxia and Vector Surveillance Officer', 'nname' => 'Khobel', 'division' => 'RD'],
		];

        // $templateProcessor = new TemplateProcessor('template.docx');
        $templateProcessor = new TemplateProcessor('word/user.docx');
        // $templateProcessor->setValue('id', '1');
        // $templateProcessor->setValue('name', 'cayce');
        // $templateProcessor->setValue('email', 'caycedy@gmail.com');
        // $templateProcessor->setValue('address', 'bakakeng');

        $templateProcessor->setValue('id', 		$values[0]['user_id']);
        $templateProcessor->setValue('name', 	$values[0]['name']);
        $templateProcessor->setValue('nname', 	$values[0]['nname']);
        $templateProcessor->setValue('position', $values[0]['position']);

        $templateProcessor->setValue('id2', 		$values[1]['user_id']);
        $templateProcessor->setValue('name2', 	$values[1]['name']);
        $templateProcessor->setValue('nname2', 	$values[1]['nname']);
        $templateProcessor->setValue('position2', $values[1]['position']);

        $templateProcessor->setValue('id3', 		$values[2]['user_id']);
        $templateProcessor->setValue('name3', 	$values[2]['name']);
        $templateProcessor->setValue('nname3', 	$values[2]['nname']);
        $templateProcessor->setValue('position3', $values[2]['position']);

        $templateProcessor->setValue('id4', 		$values[3]['user_id']);
        $templateProcessor->setValue('name4', 	$values[3]['name']);
        $templateProcessor->setValue('nname4', 	$values[3]['nname']);
        $templateProcessor->setValue('position4', $values[3]['position']);
        // $templateProcessor->setValue('email2', 	$values[1]['email']);
        // $templateProcessor->setImageValue('logo', $pic1);
        // $templateProcessor->setImageValue('pic1', 'https://images-na.ssl-images-amazon.com/images/I/41dd%2Bg%2BIxvL.jpg');
        $templateProcessor->setImageValue('pic1', array('path' => 'https://images-na.ssl-images-amazon.com/images/I/41dd%2Bg%2BIxvL.jpg', 
        	'width' => '100%', 'height' => '100%', 'ratio' => true
        ));
        $templateProcessor->setImageValue('division', array('path' => 'http://localhost/images/user.png', 
        	'width' => '595px', 'height' => '842px', 'ratio' => true
        ));
        $fileName = 'user';
        $templateProcessor->saveAs('user' . '.docx');
        return response()->download('user' . '.docx')->deleteFileAfterSend(true);
        
    }

    public function setImageValue($search, $replace)
    {
        // Sanity check
        if (!file_exists($replace))
        {
            return;
        }

        // Delete current image
        $this->zipClass->deleteName('word/media/' . $search);

        // Add a new one
        $this->zipClass->addFile($replace, 'word/media/' . $search);
    }
}
