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
        
        $documentLayout = $objPHPPowerPoint->getLayout();
        $documentLayout->setDocumentLayout($documentLayout::LAYOUT_SCREEN_16X9);

        $currentSlide = $objPHPPowerPoint->getActiveSlide();
        $currentSlide->setSlideLayout(Layout::BLANK);
        $this->setSlideBackgroundColor($currentSlide);
        $imagePlaceholder = $currentSlide->createRichTextShape();
        $imagePlaceholder->setOffsetX(676)->setOffsetY(215)->setHeight(100)->setWidth(250);
        $imagePlaceholder->createTextRun('Insert product or brand image here')->getFont()->setName('Arial')->setSize(18)->setColor(new Color(Color::COLOR_WHITE));
        $initiativeLogo = $currentSlide->createDrawingShape();
        $initiativeLogo->setPath($this->get('kernel')->getRootDir() . self::INITIATIVE_LOGO_BLUE)->setOffsetX(812)->setoffsetY(487)->setHeight(21)->setWidth(103);
        $separateLine = $currentSlide->createLineShape(786, 476, 786, 524);
        $separateLine->getBorder()->setColor(new Color('43BBEF'));
        $campaignInfoShape = $currentSlide->createRichTextShape();
        $campaignInfoShape->setHeight(133)->setWidth(500)->setOffsetX(0)->setOffsetY(282);
        $fill = new Fill();
        $fill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('43BBEF'));
        $campaignInfoShape->setFill($fill);
        $firstP = $campaignInfoShape->getActiveParagraph();
        $firstP->createTextRun($campaign->getName())->getFont()->setName('KG Second Chances Sketch')->setSize(32)->setColor(new Color(Color::COLOR_WHITE));
        $secondP = $campaignInfoShape->createParagraph();
        $secondP->createTextRun('Planning Cycle - ' . $campaign->getBrand()->getName())->getFont()->setName('KG Second Chances Sketch')->setSize(18)->setColor(new Color(Color::COLOR_WHITE));
        $thirdP = $campaignInfoShape->createParagraph();
        $thirdP->createTextRun($campaign->getCountry()->getName())->getFont()->setName('KG Second Chances Sketch')->setSize(29)->setColor(new Color(Color::COLOR_WHITE));
        
        $oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
        $pathToFile = __DIR__ . "/sample.pptx";
        $oWriterPPTX->save($pathToFile);
        return response()->download($pathToFile)->deleteFileAfterSend();
    }
}
