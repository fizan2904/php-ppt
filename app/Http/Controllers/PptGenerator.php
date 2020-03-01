<?php

namespace App\Http\Controllers;

require_once(__DIR__ . '/../../../vendor/autoload.php');

use Illuminate\Http\Request;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;

class PptGenerator extends Controller
{
    public function index(Request $request) {
        $objPHPPowerPoint = new PhpPresentation();
        
        $currentSlide = $objPHPPowerPoint->getActiveSlide();
        $shape = $currentSlide->createDrawingShape();
        
        $shape->setName('PHPPresentation logo')
            ->setDescription('PHPPresentation logo')
            ->setPath('/Users/fizan/Downloads/phppowerpoint_logo.gif')
            ->setHeight(36)
            ->setOffsetX(10)
            ->setOffsetY(10);
        $shape->getShadow()->setVisible(true)
                        ->setDirection(45)
                        ->setDistance(10);
        $shape = $currentSlide->createRichTextShape()
            ->setHeight(300)
            ->setWidth(600)
            ->setOffsetX(170)
            ->setOffsetY(180);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );

        $name = $request->input('title');

        $textRun = $shape->createTextRun($name);
        $textRun->getFont()->setBold(true)->setSize(60)->setColor( new Color( 'FFE06B20' ) );

        $currentSlide = $objPHPPowerPoint->createSlide();
        
        // $currentSlide->setSlideLayout();

        $portraitPlaceholder = $currentSlide->createRichTextShape();
        $portraitPlaceholder->setOffsetX(676)->setOffsetY(215)->setHeight(100)->setWidth(250);
        $portraitPlaceholder->createTextRun('Insert portrait here')->getFont()->setName('Arial')->setSize(18)->setColor(new Color(Color::COLOR_WHITE));
        
        $determineCircle = $currentSlide->createDrawingShape();
        $determineCircle->setPath('/Users/fizan/Downloads/color-1.jpg')->setOffsetX(19)->setoffsetY(0)->setHeight(402)->setWidth(520);

        $currentSlide = $objPHPPowerPoint->createSlide();
        $outlineShape = $currentSlide->createRichTextShape()
            ->setHeight(300)
            ->setWidth(600)
            ->setOffsetX(170)
            ->setOffsetY(180);

        $description = $request->input('description');
        
        $outlineShape->createTextRun($description)->getFont()->setBold(true)->setSize(60)->setColor(new Color('FFE06B20'));
        $outlineShape->createBreak();

        
        $oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
        $pathToFile = __DIR__ . "/sample.pptx";
        $oWriterPPTX->save($pathToFile);
        return response()->download($pathToFile)->deleteFileAfterSend();
    }
}
