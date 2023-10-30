package org.mattpayne.demo.excel.excel1;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.FileOutputStream;

@Controller
public class ExcelController {

    @GetMapping("/download-excel")
    public ResponseEntity<byte[]> downloadExcel() {
        try {
            // Create a new workbook
            Workbook workbook = new XSSFWorkbook();

            // Create "summary" sheet
            Sheet summarySheet = workbook.createSheet("summary");
            Row summaryRow = summarySheet.createRow(0);
            Cell summaryCell = summaryRow.createCell(0);
            summaryCell.setCellValue("This is the summary sheet.");

            // Create "data" sheet
            Sheet dataSheet = workbook.createSheet("data");
            Row dataRow = dataSheet.createRow(0);
            Cell dataCell = dataRow.createCell(0);
            dataCell.setCellValue("This is the data sheet.");

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
