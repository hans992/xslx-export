<?php
    require 'vendor/autoload.php';
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

    $db_conx = mysqli_connect('127.0.0.1', 'root', 'koliko031', 'zadatak');
        

    if(isset($_POST['exportButton'])){
        $conditions = "";

        if($_POST['start'] != "") {
            $start = $_POST['start'];
            $conditions .= "date >= '$start'";
        }else{
            $start = "0000-00-00";
            $conditions .= "date >= '$start'";
        }

        if( $_POST['end'] != "") {
            $conditions .= " AND ";
            $end = $_POST['end'];
            $conditions .= "date <= '$end'";
        }else{
            $conditions .= " AND ";
            $end = date("Y-m-d");
            $conditions .= "date <= '$end'";
        }


        $message = "Successfully exported data from ".$start." to ".$end;


        /**************************************************************************************
         ********* SHEET 1
         **************************************************************************************/

            $sql = mysqli_query($db_conx, "SELECT * FROM delivery_stats WHERE $conditions");

            $tel = 0;
            $mail = 0;
            $lieferung = 0;
            $average = 0;
            $iterator_partial = 0;
            $iterator_total = 0;



            while ($row=mysqli_fetch_array($sql)){

                if($row['method'] == 'Telefon') $tel += $row['count'];
                elseif($row['method'] == 'E-Mail') $mail += $row['count'];
                elseif($row['method'] == 'Lieferung') $lieferung += $row['count'];

                if($row['cartValue'] != '0') {
                    $average += $row['cartValue'];
                    $iterator_partial++;

                }

                $iterator_total++;
            }



            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $sheet->setCellValue('A1', "$start - $end");

            $sheet->setCellValue('B1', "via Telefon");
            $sheet->setCellValue('C1', "via E-Mail");
            $sheet->setCellValue('D1', "via Lieferung");
            $sheet->setCellValue('E1', "Total orders");
            $sheet->setCellValue('F1', "Average order");

            $sheet->setCellValue('B2', $tel);
            $sheet->setCellValue('C2', $mail);
            $sheet->setCellValue('D2', $lieferung);
            $sheet->setCellValue('E2', $iterator_total);
            $sheet->setCellValue('F2', $average/$iterator_partial);





        /**************************************************************************************
         ********* SHEET 2
         **************************************************************************************/

            $sql = mysqli_query($db_conx, "SELECT date, method, count(count) as orders 
                                                  FROM delivery_stats 
                                                  WHERE $conditions 
                                                  GROUP BY DATE_FORMAT(date, '%Y-%m'), method");

            $spreadsheet->createSheet();
            $sheet = $spreadsheet->getSheetByName('Worksheet 1');

            $sheet->setCellValue('A1', "Mjesec");
            $sheet->setCellValue('B1', "Metoda");
            $sheet->setCellValue('C1', "Count");

            $i = 2;

            while ($row=mysqli_fetch_array($sql)){
                $method = $row['method'];
                $count = $row['orders'];
                $format_date = explode('-', $row['date']);

                $sheet->setCellValue("A$i", "$format_date[1]-$format_date[0]");
                $sheet->setCellValue("B$i", "$method");
                $sheet->setCellValue("C$i", "$count");

                $i++;
            }


            $writer = new Xlsx($spreadsheet);
            $writer->save('zadatak.xlsx');
    }
?>

<html>
    <body>
        <form method="post" action="">
            <input id="start" type="date" name="start"><br>
            <input id="end" type="date" name="end"><br>
            <input type="submit" name="exportButton" value="Export to xmls"><br>
            <?php if(isset($_POST['exportButton'])){ echo $message; } ?>
        </form>
    </body>
</html>