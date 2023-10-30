package org.mattpayne.demo.excel.excel1;


import org.apache.commons.math3.stat.descriptive.rank.Median;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
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

            // Generate random data and calculate statistics
            Random random = new Random();
            CellStyle decimalStyle = workbook.createCellStyle();
            DataFormat df = workbook.createDataFormat();
            decimalStyle.setDataFormat(df.getFormat("0.00"));
            List<List<Double>> data = new ArrayList<>();

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
                List<Double> rowData = new ArrayList<>();
                for (int colIndex = 0; colIndex < 13; colIndex++) {
                    Cell dataCell = dataRow.createCell(colIndex);
                    if (colIndex == 0) {
                        dataCell.setCellValue("A" + (rowIndex + 1));
                    } else {
                        double randomValue = -5 + (100 + 5) * random.nextDouble();
                        dataCell.setCellValue(randomValue);
                        dataCell.setCellStyle(decimalStyle);
                        rowData.add(randomValue);
                    }
                }
                data.add(rowData);
            }

            // Write summary statistics to the "summary" sheet
            for (int rowIndex = 0; rowIndex < 10; rowIndex++) {
                Row summaryRow = summarySheet.createRow(rowIndex);
                Cell meanCell = summaryRow.createCell(0);
                Cell medianCell = summaryRow.createCell(1);
                Cell maxCell = summaryRow.createCell(2);
                Cell minCell = summaryRow.createCell(3);
                Cell sumCell = summaryRow.createCell(4);

                List<Double> rowData = data.get(rowIndex);
                double mean = calculateMean(rowData);
                double median = calculateMedian(rowData);
                double max = calculateMax(rowData);
                double min = calculateMin(rowData);
                double sum = calculateSum(rowData);

                meanCell.setCellValue(mean);
                medianCell.setCellValue(median);
                maxCell.setCellValue(max);
                minCell.setCellValue(min);
                sumCell.setCellValue(sum);

                meanCell.setCellStyle(decimalStyle);
                medianCell.setCellStyle(decimalStyle);
                maxCell.setCellStyle(decimalStyle);
                minCell.setCellStyle(decimalStyle);
                sumCell.setCellStyle(decimalStyle);
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

    private double calculateMean(List<Double> data) {
        double sum = 0;
        for (Double value : data) {
            sum += value;
        }
        return sum / data.size();
    }

    private double calculateMedian(List<Double> data) {
        double[] values = new double[data.size()];
        for (int i = 0; i < data.size(); i++) {
            values[i] = data.get(i);
        }
        Median medianFunction = new Median();
        return medianFunction.evaluate(values);
    }

    private double calculateMax(List<Double> data) {
        return data.stream().max(Double::compare).orElse(0.0);
    }

    private double calculateMin(List<Double> data) {
        return data.stream().min(Double::compare).orElse(0.0);
    }

    private double calculateSum(List<Double> data) {
        double sum = 0;
        for (Double value : data) {
            sum += value;
        }
        return sum;
    }
}

