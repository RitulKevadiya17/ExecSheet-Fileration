package com.example.execsheetfileration.Controller;

import com.example.execsheetfileration.ResponseDTO.DataResponse;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.BufferedReader;
import java.util.ArrayList;
import java.util.List;

@Controller
class ExcelController {

    @GetMapping("uploadPage")
    public String exam(){
        return "File-Upload-Page";
    }

    @PostMapping("/downloadExel")
    public ResponseEntity<Resource> uploadFile(@RequestParam("file") MultipartFile file) throws IOException {
        String fileName = file.getOriginalFilename();

//        if (fileName != null && (fileName.endsWith(".xlsx") || fileName.endsWith(".xls"))) {
//            // Handle Excel file
//            return readExcelFile(file.getInputStream());
//        } else
            if (fileName != null && fileName.endsWith(".csv")) {
            // Handle CSV file
            return readCsvFile(file.getInputStream());
        } else {
            // Unsupported file format
            throw new IllegalArgumentException("Unsupported file format: " + fileName);
        }
    }

    private void readExcelFile(java.io.InputStream inputStream) throws IOException {
        Workbook workbook = new XSSFWorkbook(inputStream);
//        return readSheet(workbook.getSheetAt(0));
    }

    private ResponseEntity<Resource> readCsvFile(java.io.InputStream inputStream) throws IOException {
        List<List<String>> data = new ArrayList<>();
        DataResponse dataResponse = new DataResponse();

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream))) {
            String line;
            int countEMERALD = 0;
            int countSPEAR = 0;
            int countCUSHION = 0;
            int countOVAL = 0;
            int countPRINCESS = 0;
            int countSMARQUISE = 0;
            while ((line = reader.readLine()) != null) {

                String[] values = line.split(",");
//              List<String> rowData = List.of(values);
                List<String> diamondType = new ArrayList<>();
                diamondType.add(List.of(values).get(6));
                for(String list : diamondType){
                    if(list.equals("EMERALD 4STEP")){
                        countEMERALD++;
                    }
                    if(list.equals("S.PEAR")){
                        countSPEAR++;
                    }
                    if(list.equals("CUSHION")){
                        countCUSHION++;
                    }
                    if(list.equals("OVAL")){
                        countOVAL++;
                    }
                    if(list.equals("PRINCESS")){
                        countPRINCESS++;
                    }
                    if(list.equals("S.MARQUISE")){
                        countSMARQUISE++;
                    }
                }
//                data.add(rowData);
                data.add(diamondType);
            }
            dataResponse.setCountEMERALD(countEMERALD);
            dataResponse.setCountSPEAR(countSPEAR);
            dataResponse.setCountCUSHION(countCUSHION);
            dataResponse.setCountOVAL(countOVAL);
            dataResponse.setCountPRINCESS(countPRINCESS);
            dataResponse.setCountSMARQUISE(countSMARQUISE);
            dataResponse.setTotalType(countEMERALD+countSPEAR+countCUSHION
                    +countOVAL+countPRINCESS+countSMARQUISE);
        }


        return generateExcelFile(dataResponse);
    }

//    private int readSheet(Sheet sheet) {
//        int rows = sheet.getPhysicalNumberOfRows();
//        int cols = sheet.getRow(0).getPhysicalNumberOfCells();
//
//        List<List<String>> data = new ArrayList<>();
//
//        for (int i = 1; i < rows; i++) {
//            List<String> rowData = new ArrayList<>();
//            for (int j = 0; j < cols; j++) {
//                Cell cell = sheet.getRow(i).getCell(j);
//                String value = cellToString(cell);
//                rowData.add(value);
//            }
//            data.add(rowData);
//        }
//        return data;
//    }

    private String cellToString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private ResponseEntity<Resource> generateExcelFile(DataResponse dataResponse) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Diamond Data");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Diamond Type");
        headerRow.createCell(1).setCellValue("Count");

        // Add data rows
        addDataRow(sheet, 1, "EMERALD", dataResponse.getCountEMERALD());
        addDataRow(sheet, 2, "S.PEAR", dataResponse.getCountSPEAR());
        addDataRow(sheet, 3, "CUSHION", dataResponse.getCountCUSHION());
        addDataRow(sheet, 4, "OVAL", dataResponse.getCountOVAL());
        addDataRow(sheet, 5, "PRINCESS", dataResponse.getCountPRINCESS());
        addDataRow(sheet, 6, "S.MARQUISE", dataResponse.getCountSMARQUISE());

        // Create total row
        Row totalRow = sheet.createRow(7);
        totalRow.createCell(0).setCellValue("Total");
        totalRow.createCell(1).setCellValue(dataResponse.getTotalType());

        // Convert workbook to byte array
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();

        // Create ByteArrayResource from byte array
        ByteArrayResource resource = new ByteArrayResource(outputStream.toByteArray());

        // Set appropriate headers for Excel file download
        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=diamond_data.xlsx");

        return ResponseEntity.ok()
                .headers(headers)
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(resource);
    }

    private void addDataRow(Sheet sheet, int rowNum, String diamondType, int count) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(diamondType);
        row.createCell(1).setCellValue(count);
    }
}
