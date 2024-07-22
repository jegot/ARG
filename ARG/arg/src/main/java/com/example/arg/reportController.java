package com.example.arg;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Controller
public class reportController {

    @Autowired
    private ExcelService excelService;

    @GetMapping("/")
    public String index() {
        return "index";
    }

    public String reportPath = "C:/Users/jgolling/Desktop/programming/ARG/thisReport.pdf";

    @GetMapping("/generate-report")
    public ResponseEntity<InputStreamResource> generateReport(@RequestParam("date") String date) {
        try {
            reportPath = "C:/Users/jgolling/Desktop/programming/ARG/Report_.pdf";
    
            List<String> machineNames = new ArrayList<>();
            machineNames.add("INKJET1");
            machineNames.add("INKJET2");
            machineNames.add("INKJET3");
            machineNames.add("INKJET4");
            machineNames.add("INKJET5");
            machineNames.add("INKJET6");
            machineNames.add("INSERT1");
            machineNames.add("INSERTER 9");
    
            List<List<String[]>> allData = new ArrayList<>();
    
            for (String machine : machineNames) {
                // Read data from the Excel file for the specified date
                List<String[]> data = excelService.readExcel(date, machine);
    
                if (data != null && !data.isEmpty()) {
                    allData.add(data);
                }
            }
    
            if (!allData.isEmpty()) {
                // Generate the PDF report with all data
                excelService.generatePdfReport(allData, reportPath, date);
            }
    
            // Create the PDF file response
            File file = new File(reportPath);
            System.out.println("Creating file: " + reportPath);
    
            InputStreamResource resource = new InputStreamResource(new FileInputStream(file));
    
            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=report.pdf");
            return ResponseEntity.ok()
                    .headers(headers)
                    .contentLength(file.length())
                    .contentType(MediaType.APPLICATION_PDF)
                    .body(resource);
        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.status(500).build();
        }

}



    
}
