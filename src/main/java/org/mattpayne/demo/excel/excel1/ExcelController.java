package org.mattpayne.demo.excel.excel1;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.FileOutputStream;
import java.util.Random;

@Controller
public class ExcelController {

    @GetMapping("/download-excel")
    public ResponseEntity<byte[]> downloadExcel() {
        try {
            // Create a new workbook
            Workbook workbook = new XSSFWorkbook();

            // Create "summary" sheet
            Sheet summarySheet = workbook.createSheet("summary");
            // Create "data" sheet
            Sheet dataSheet = workbook.createSheet("data");

            // Generate random data
            Random random = new Random();
            CellStyle decimalStyle = workbook.createCellStyle();
            DataFormat df = workbook.createDataFormat();
            decimalStyle.setDataFormat(df.getFormat("0.00"));

            // Create headers for the "summary" sheet
            Row summaryHeader = summarySheet.createRow(0);
            summaryHeader.createCell(0).setCellValue("Row");
            summaryHeader.createCell(1).setCellValue("Mean");
            summaryHeader.createCell(2).setCellValue("Median");
            summaryHeader.createCell(3).setCellValue("Max");
            summaryHeader.createCell(4).setCellValue("Min");
            summaryHeader.createCell(5).setCellValue("Sum");

            for (int rowIndex = 0; rowIndex < 10; rowIndex++) {
                Row dataRow = dataSheet.createRow(rowIndex);
                for (int colIndex = 0; colIndex < 13; colIndex++) {
                    Cell dataCell = dataRow.createCell(colIndex);
                    if (colIndex == 0) {
                        dataCell.setCellValue("A" + (rowIndex + 1));
                    } else {
                        double randomValue = -5 + (100 + 5) * random.nextDouble();
                        dataCell.setCellValue(randomValue);
                        dataCell.setCellStyle(decimalStyle);
                    }
                }
            }

            // Write Excel formulas to calculate statistics in the "summary" sheet
            for (int rowIndex = 0; rowIndex < 10; rowIndex++) {
                Row summaryRow = summarySheet.createRow(rowIndex + 1); // Add 1 for the header row
                summaryRow.createCell(0).setCellValue("Row " + (rowIndex + 1));
                summaryRow.createCell(1).setCellFormula("AVERAGE(data!B" + (rowIndex + 1) + ":N" + (rowIndex + 1) + ")");
                summaryRow.createCell(2).setCellFormula("MEDIAN(data!B" + (rowIndex + 1) + ":N" + (rowIndex + 1) + ")");
                summaryRow.createCell(3).setCellFormula("MAX(data!B" + (rowIndex + 1) + ":N" + (rowIndex + 1) + ")");
                summaryRow.createCell(4).setCellFormula("MIN(data!B" + (rowIndex + 1) + ":N" + (rowIndex + 1) + ")");
                summaryRow.createCell(5).setCellFormula("SUM(data!B" + (rowIndex + 1) + ":N" + (rowIndex + 1) + ")");
            }

            // Create a temporary file to save the workbook
            java.io.File tempFile = java.io.File.createTempFile("sample", ".xlsx");
            try (FileOutputStream outputStream = new FileOutputStream(tempFile)) {
                workbook.write(outputStream);
            }

            // Prepare the response for download
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.setContentDispositionFormData("attachment", "sample.xlsx");

            byte[] fileContent = org.apache.commons.io.FileUtils.readFileToByteArray(tempFile);

            return new ResponseEntity<>(fileContent, headers, org.springframework.http.HttpStatus.OK);
        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<>(org.springframework.http.HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}
