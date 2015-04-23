<?php    
 /* CAT:Bar Chart */ 

 /* pChart library inclusions */ 
 include("pChart2.1.3/class/pData.class.php"); 
 include("pChart2.1.3/class/pDraw.class.php"); 
 include("pChart2.1.3/class/pImage.class.php"); 

 /* Create and populate the pData object */ 
 $MyData = new pData();   
 $MyData->loadPalette("pChart2.1.3/palettes/light.color",TRUE); 
 $MyData->addPoints(array(150,220,300,-250,-420,-200,300,200,100),"Server A"); 
 $MyData->addPoints(array(140,0,340,-300,-320,-300,200,100,50),"Server B"); 
 $MyData->setAxisName(0,"Hits"); 
 $MyData->addPoints(array("������","�������","����","������","���","����","����","������","��������"),"������"); 
 $MyData->setSerieDescription("Months","Month"); 
 $MyData->setAbscissa("Months"); 

 /* Create the pChart object */ 
 $myPicture = new pImage(700,230,$MyData); 
 $myPicture->drawGradientArea(0,0,700,230,DIRECTION_VERTICAL,array("StartR"=>240,"StartG"=>240,"StartB"=>240,"EndR"=>180,"EndG"=>180,"EndB"=>180,"Alpha"=>100)); 
 $myPicture->drawGradientArea(0,0,700,230,DIRECTION_HORIZONTAL,array("StartR"=>240,"StartG"=>240,"StartB"=>240,"EndR"=>180,"EndG"=>180,"EndB"=>180,"Alpha"=>20)); 
 $myPicture->setFontProperties(array("FontName"=>"pChart2.1.3/fonts/verdana.ttf","FontSize"=>8)); 

 /* Draw the scale  */ 
 $myPicture->setGraphArea(50,30,680,200); 
 $myPicture->drawScale(array("CycleBackground"=>TRUE,"DrawSubTicks"=>TRUE,"GridR"=>0,"GridG"=>0,"GridB"=>0,"GridAlpha"=>10)); 

 /* Turn on shadow computing */  
 $myPicture->setShadow(TRUE,array("X"=>1,"Y"=>1,"R"=>0,"G"=>0,"B"=>0,"Alpha"=>10)); 

 /* Draw the chart */ 
 $settings = array("Gradient"=>TRUE,"GradientMode"=>GRADIENT_EFFECT_CAN,"DisplayPos"=>LABEL_POS_INSIDE,"DisplayValues"=>TRUE,"DisplayR"=>255,"DisplayG"=>255,"DisplayB"=>255,"DisplayShadow"=>TRUE,"Surrounding"=>10); 
 $myPicture->drawBarChart($settings); 

 /* Write the chart legend */ 
 $myPicture->drawLegend(580,12,array("Style"=>LEGEND_NOBORDER,"Mode"=>LEGEND_HORIZONTAL)); 

 /* Render the picture (choose the best way) */ 
 $myPicture->autoOutput("pictures/example.drawBarChart.can.png"); 