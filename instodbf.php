<?

function empty_data()
{
global $balkazna_file;
$balkazna_file_data = dbase_open($balkazna_file, 2);

$balkazna_file_data_record_num = dbase_numrecords($balkazna_file_data);


  for ($y = 1; $y <= $balkazna_file_data_record_num; $y++)
  {
  	$row = dbase_get_record_with_names($balkazna_file_data, $y);
    $row['ZALK'] = '0';
	unset($row['deleted']);
 	$row_ins = array_values($row);
	dbase_replace_record($balkazna_file_data, $row_ins, $y);
  }
echo "количество записей: ".$y;
// закрыть dbf файла баланса

dbase_close($balkazna_file_data);
}

function insert_data ()
{

global $balbars_file, $balkazna_file, $workdir;

$checked_balacc = array ("1241", "3125", "3261", "3262", "3282", "4456", "4465");

// print_r($checked_balacc);

// открыть dbf файла баланса для записи

$balkazna_file_data = dbase_open($balkazna_file, 2);

$balkazna_file_data_record_num = dbase_numrecords($balkazna_file_data);

// открытие XLS файла

// $filename = "D:\Visual Studio 2008\Projects\zvitexp\Debug\bars\Баланс 10 02.xls";

$filename = str_replace("\\", "/" , $balbars_file);

$sheet1 = "Лист1";

$excel_app = new COM("Excel.application") or Die ("Did not connect");

$excel_app->Visible = 1;

$Workbook = $excel_app->Workbooks->Open("$filename") or Die("Did not open $filename $Workbook");


$i=5;
$excel_result_balacc = '0000';
$pattern = "/[1-4][0-9]{3}/";

while ($excel_result_balacc !='')
{
        // $coord_razom = "C" . $i;

        $coord_balacc = "B" . $i;
        $coord_bars_deb = "E" . $i;
        $coord_bars_kred = "F" . $i;

        $Worksheet = $Workbook->Worksheets($sheet1);
        $Worksheet->activate;

        /*
        $excel_cell_razom = $Worksheet->Range($coord_razom);
        $excel_cell_razom->activate;
        $excel_result_razom = $excel_cell_razom->value;
        */

        $excel_cell_balacc = $Worksheet->Range($coord_balacc);
        $excel_cell_balacc->activate;
        $excel_result_balacc = $excel_cell_balacc->value;

        $excel_cell_bars_deb = $Worksheet->Range($coord_bars_deb);
        $excel_cell_bars_deb->activate;
        $excel_result_bars_deb = $excel_cell_bars_deb->value;

// $bars_deb_p = explode(".", $excel_result_bars_deb);
// $bars_deb = implode(",", $bars_deb_p);


        $excel_cell_bars_kred = $Worksheet->Range($coord_bars_kred);
        $excel_cell_bars_kred->activate;
        $excel_result_bars_kred = $excel_cell_bars_kred->value;

// $bars_kred_p = explode(".", $excel_result_bars_kred);
// $bars_kred = implode(",", $bars_kred_p);

// print $excel_cell_balacc. "\n";

// echo "\n";


  for ($y = 1; $y <= $balkazna_file_data_record_num; $y++)
  {

  $row = dbase_get_record_with_names($balkazna_file_data, $y);

  if ($row['KOBL'] == '4')
  	{

  	 $row['ZALK'] = '0';

     if ($row['NRAX'] == trim($excel_result_balacc))
       	{
         if(preg_match($pattern, $row['NRAX'])==1)
         	{


        		if (trim($row['PR']) == 'А')
          		{          			if(in_array($row['NRAX'], $checked_balacc))
          			{          			echo "балансовий рахунок ".$row['NRAX']." має залишок ".$excel_result_bars_deb."\n";
          			$row['ZALK'] = 0;
          			}
          			else
          			{          			echo "балансовий рахунок ".$row['NRAX']."\n";          			$row['ZALK'] = $excel_result_bars_deb/1000;
          			}

          		}

          		if (trim($row['PR']) == 'П')
          		{
          		    if(in_array($row['NRAX'], $checked_balacc))
          			{
          			echo "балансовий рахунок ".$row['NRAX']." має залишок ".$excel_result_bars_kred."\n";
          			$row['ZALK'] = 0;
          			}
          			else
          			{          			echo "балансовий рахунок ".$row['NRAX']."\n";
          			$row['ZALK'] = $excel_result_bars_kred/1000;
          			}

          		}

            	unset($row['deleted']);

          		$row_ins = array_values($row);

          		// print_r($row_ins);

          		dbase_replace_record($balkazna_file_data, $row_ins, $y);
         	}
         }
  	}
  }

        $i = $i + 1;
}

// закрыть dbf файла баланса

dbase_close($balkazna_file_data);

// закрыть excel

$excel_app->Quit();

// освободить объект

//$excel_app->Release();

$excel_app = null;

}

$workdir = str_replace("\\", "/" , $work_dir);

// $balkazna_blank_file = $workdir."/form/F11M.DBF";
// $balkazna_file = $workdir."/F11M.DBF";

$balkazna_file = "D:\zvitp\F11M.DBF";

// копирование бланка dbf файла баланса для программы zvitp

if (!copy($balkazna_file, $balkazna_file.".bak")) {
    echo "failed to copy $file...\n";
}

if ($empty == 1)
{empty_data();
}


// чтение XLS файла и вставка данных в

insert_data ();

?>
