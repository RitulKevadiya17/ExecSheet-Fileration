package com.example.execsheetfileration.Controller;

import com.example.execsheetfileration.ResponseDTO.DataResponse;
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
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.BufferedReader;
import java.text.DecimalFormat;
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

        if (fileName != null && (fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || fileName.endsWith(".csv"))) {
            // Handle Excel file
            return readCsvFile(file.getInputStream());
        } else {
            // Unsupported file format
            throw new IllegalArgumentException("Unsupported file format: " + fileName);
        }
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
            List<Double> kachuListOfEmerald = new ArrayList<>();
            List<Double> kachuListOfPear = new ArrayList<>();
            List<Double> kachuListOfCushion = new ArrayList<>();
            List<Double> kachuListOfOval = new ArrayList<>();
            List<Double> kachuListOfPrincess = new ArrayList<>();
            List<Double> kachuListOfMarquise = new ArrayList<>();

            List<Double> polishedListOfEmerald = new ArrayList<>();
            List<Double> polishedListOfPear = new ArrayList<>();
            List<Double> polishedListOfCushion = new ArrayList<>();
            List<Double> polishedListOfOval = new ArrayList<>();
            List<Double> polishedListOfPrincess = new ArrayList<>();
            List<Double> polishedListOfMarquise = new ArrayList<>();

            while ((line = reader.readLine()) != null) {

                String[] values = line.split(",");
                List<String> diamondType = new ArrayList<>();
                diamondType.add(List.of(values).get(6));
                for(String list : diamondType){
                    if(list.equals("EMERALD 4STEP")){
                        kachuListOfEmerald.add(Double.valueOf(List.of(values).get(8)));
                        polishedListOfEmerald.add(Double.valueOf(List.of(values).get(9)));
                        countEMERALD++;
                    }
                    if(list.equals("S.PEAR")){
                        kachuListOfPear.add(Double.valueOf(List.of(values).get(8)));
                        polishedListOfPear.add(Double.valueOf(List.of(values).get(9)));
                        countSPEAR++;
                    }
                    if(list.equals("CUSHION")){
                        kachuListOfCushion.add(Double.valueOf(List.of(values).get(8)));
                        polishedListOfCushion.add(Double.valueOf(List.of(values).get(9)));
                        countCUSHION++;
                    }
                    if(list.equals("OVAL")){
                        kachuListOfOval.add(Double.valueOf(List.of(values).get(8)));
                        polishedListOfOval.add(Double.valueOf(List.of(values).get(9)));
                        countOVAL++;
                    }
                    if(list.equals("PRINCESS")){
                        kachuListOfPrincess.add(Double.valueOf(List.of(values).get(8)));
                        polishedListOfPrincess.add(Double.valueOf(List.of(values).get(9)));
                        countPRINCESS++;
                    }
                    if(list.equals("S.MARQUISE")){
                        kachuListOfMarquise.add(Double.valueOf(List.of(values).get(8)));
                        polishedListOfMarquise.add(Double.valueOf(List.of(values).get(9)));
                        countSMARQUISE++;
                    }
                }
                data.add(diamondType);
            }
            double totalWeightOfEmerald = kachuListOfEmerald.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();
            double totalWeightOfPear = kachuListOfPear.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();
            double totalWeightOfCushion = kachuListOfCushion.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();
            double totalWeightOfOval = kachuListOfOval.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();
            double totalWeightOfPrincess = kachuListOfPrincess.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();
            double totalWeightOfMarquise = kachuListOfMarquise.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();

            dataResponse.setTotalWeightOfEmerald(totalWeightOfEmerald);
            dataResponse.setTotalWeightOfPear(totalWeightOfPear);
            dataResponse.setTotalWeightOfCushion(totalWeightOfCushion);
            dataResponse.setTotalWeightOfOval(totalWeightOfOval);
            dataResponse.setTotalWeightOfPrincess(totalWeightOfPrincess);
            dataResponse.setTotalWeightOfMarquise(totalWeightOfMarquise);
            dataResponse.setTotalkachuWeight(totalWeightOfEmerald+totalWeightOfPear+totalWeightOfCushion+
                    totalWeightOfOval+totalWeightOfPrincess+totalWeightOfMarquise);

            double totalPolishedWeightOfEmerald = polishedListOfEmerald.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();
            double totalPolishedWeightOfPear = polishedListOfPear.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();
            double totalPolishedWeightOfCushion = polishedListOfCushion.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();
            double totalPolishedWeightOfOval = polishedListOfOval.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();
            double totalPolishedWeightOfPrincess = polishedListOfPrincess.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();
            double totalPolishedWeightOfMarquise = polishedListOfMarquise.stream()
                    .mapToDouble(Double::doubleValue)
                    .sum();

            dataResponse.setTotalPolishedWeightOfEmerald(totalPolishedWeightOfEmerald);
            dataResponse.setTotalPolishedWeightOfPear(totalPolishedWeightOfPear);
            dataResponse.setTotalPolishedWeightOfCushion(totalPolishedWeightOfCushion);
            dataResponse.setTotalPolishedWeightOfOval(totalPolishedWeightOfOval);
            dataResponse.setTotalPolishedWeightOfPrincess(totalPolishedWeightOfPrincess);
            dataResponse.setTotalPolishedWeightOfMarquise(totalPolishedWeightOfMarquise);
            dataResponse.setTotalPolishedWeight(totalPolishedWeightOfEmerald+totalPolishedWeightOfPear+totalPolishedWeightOfCushion+
                    totalPolishedWeightOfOval+totalPolishedWeightOfPrincess+totalPolishedWeightOfMarquise);

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

    private ResponseEntity<Resource> generateExcelFile(DataResponse dataResponse) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Diamond Data");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(4).setCellValue("DIAMOND-TYPE");
        headerRow.createCell(5).setCellValue("COUNT");
        headerRow.createCell(6).setCellValue("KACHU-WEIGHT");
        headerRow.createCell(7).setCellValue("POLISHED-WEIGHT");
        headerRow.createCell(8).setCellValue("PRCENTAGE");
        headerRow.createCell(9).setCellValue("PRCENTAGE1");

        DecimalFormat decimalFormat = new DecimalFormat("#.00");
        String a =decimalFormat.format((dataResponse.getTotalPolishedWeightOfEmerald()/dataResponse.getTotalWeightOfEmerald())*100)+"%";
        String b =decimalFormat.format((dataResponse.getTotalPolishedWeightOfPear()/dataResponse.getTotalWeightOfPear())*100)+"%";
        String c =decimalFormat.format((dataResponse.getTotalPolishedWeightOfCushion()/dataResponse.getTotalWeightOfCushion())*100)+"%";
        String d =decimalFormat.format((dataResponse.getTotalPolishedWeightOfOval()/dataResponse.getTotalWeightOfOval())*100)+"%";
        String e =decimalFormat.format((dataResponse.getTotalPolishedWeightOfPrincess()/dataResponse.getTotalWeightOfPrincess())*100)+"%";
        String f =decimalFormat.format((dataResponse.getTotalPolishedWeightOfMarquise()/dataResponse.getTotalWeightOfMarquise())*100)+"%";

        String a1 = decimalFormat.format(dataResponse.getTotalPolishedWeightOfEmerald()/dataResponse.getTotalPolishedWeight()*100)+"%";
        String b1 = decimalFormat.format(dataResponse.getTotalPolishedWeightOfPear()/dataResponse.getTotalPolishedWeight()*100)+"%";
        String c1 = decimalFormat.format(dataResponse.getTotalPolishedWeightOfCushion()/dataResponse.getTotalPolishedWeight()*100)+"%";
        String d1 = decimalFormat.format(dataResponse.getTotalPolishedWeightOfOval()/dataResponse.getTotalPolishedWeight()*100)+"%";
        String e1 = decimalFormat.format(dataResponse.getTotalPolishedWeightOfPrincess()/dataResponse.getTotalPolishedWeight()*100)+"%";
        String f1 = decimalFormat.format(dataResponse.getTotalPolishedWeightOfMarquise()/dataResponse.getTotalPolishedWeight()*100)+"%";

        // Add data rows
        addDataRow(sheet, 3, "EMERALD", dataResponse.getCountEMERALD(),dataResponse.getTotalWeightOfEmerald(),dataResponse.getTotalPolishedWeightOfEmerald(),
                a,a1);
        addDataRow(sheet, 4, "S.PEAR", dataResponse.getCountSPEAR(), dataResponse.getTotalWeightOfPear(), dataResponse.getTotalPolishedWeightOfPear(),
                b,b1);
        addDataRow(sheet, 5, "CUSHION", dataResponse.getCountCUSHION(), dataResponse.getTotalWeightOfCushion(), dataResponse.getTotalPolishedWeightOfCushion(),
                c,c1);
        addDataRow(sheet, 6, "OVAL", dataResponse.getCountOVAL(), dataResponse.getTotalWeightOfOval(), dataResponse.getTotalPolishedWeightOfOval(),
                d,d1);
        addDataRow(sheet, 7, "PRINCESS", dataResponse.getCountPRINCESS(), dataResponse.getTotalWeightOfPrincess(), dataResponse.getTotalPolishedWeightOfPrincess(),
                e,e1);
        addDataRow(sheet, 8, "S.MARQUISE", dataResponse.getCountSMARQUISE(), dataResponse.getTotalWeightOfMarquise(), dataResponse.getTotalPolishedWeightOfMarquise(),
                f,f1);
        // Create total row
        Row totalRow = sheet.createRow(10);
        totalRow.createCell(4).setCellValue("TOTAL");
        totalRow.createCell(5).setCellValue(dataResponse.getTotalType());
        totalRow.createCell(6).setCellValue(dataResponse.getTotalkachuWeight());
        totalRow.createCell(7).setCellValue(dataResponse.getTotalPolishedWeight());

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

    private void addDataRow(Sheet sheet, int rowNum, String diamondType, int count, double totalKachuWeight ,
                            double totalPolishedWeight , String percentage, String percentage1) {
        Row row = sheet.createRow(rowNum);
        row.createCell(4).setCellValue(diamondType);
        row.createCell(5).setCellValue(count);
        row.createCell(6).setCellValue(totalKachuWeight);
        row.createCell(7).setCellValue(totalPolishedWeight);
        row.createCell(8).setCellValue(percentage);
        row.createCell(9).setCellValue(percentage1);
    }
}
