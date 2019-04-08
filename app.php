<?php
  //starting excel
  $excel = new COM("excel.application") or die("Unable to instanciate excel");
  print "Loaded excel, version {$excel->Version}\n";

  //bring it to front
  #$excel->Visible = 1;//NOT

  //dont want alerts ... run silent
  $excel->DisplayAlerts = 0;

  //create a new workbook
  $wkb = $excel->Workbooks->Add();

  //select the default sheet
  $sheet=$wkb->Worksheets(1);

  $xlMod = $wkb->VBProject->VBComponents->Add(1);
  $xlMod->Name = "Module1";
  $macroCode = "Sub main()\r\n" .
         "   MsgBox \"Hello world\"\r\n" .
         "end Sub";
 
  $xlMod->CodeModule->AddFromString($macroCode);
    

  //make it the active sheet
  $sheet->activate;

  //fill it with some bogus data
  for($row=1;$row<=7;$row++){
      for ($col=1;$col<=5;$col++){
         $sheet->activate;
         $cell=$sheet->Cells($row,$col);
         $cell->Activate;
        //  $cell->value = 'pool4tool 4eva ' . $row . ' ' . $col . ' ak';
         $cell->value = '=now()';
      }//end of colcount for loop
  }

  ///////////
  // Select Rows 2 to 5
  $r = $sheet->Range("2:5")->Rows;

  // group them baby, yeah
  $r->Cells->Group;

  // save the new file
  $strPath = 'D:\Workspace\php-example\out\example.xls';
  if (file_exists($strPath)) {unlink($strPath);}
  $wkb->SaveAs($strPath, 1);

  //close the book
  $wkb->Close(false);
  $excel->Workbooks->Close();

  //free up the RAM
  unset($sheet);

  //closing excel
  $excel->Quit();

  //free the object
  $xlMod = null;
  $excel = null; 
?>
