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

    @GetMapping("/generate-report")
    public ResponseEntity<InputStreamResource> generateReport(@RequestParam("date") String date) {
        try {
            // Replace the slashes in the date string with dashes
            String pathDate = date.replace('/', '-');
    
            //String userHome = System.getProperty("user.home");
            String reportFilename = "Report_" + pathDate + ".pdf";

            String reportPath = "Y:/Machine Reports/reports/" + reportFilename;

            

    
            // Check if the report already exists
            File file = new File(reportPath);
            
            List<String> machineNames = new ArrayList<>();
            machineNames.add("INKJET1");
            machineNames.add("INKJET2");
            machineNames.add("INKJET3");
            machineNames.add("INKJET4");
            machineNames.add("INKJET5");
            machineNames.add("INKJET6");
            machineNames.add("INSERT1");
            machineNames.add("INSERT2");
            machineNames.add("INSERT3");
            machineNames.add("INSERT4");
            machineNames.add("INSERT5");
            machineNames.add("INSERT6");
            machineNames.add("INSERT7");
            machineNames.add("INSERT8");
            machineNames.add("INSERT9");
            machineNames.add("INSERT10");
            machineNames.add("INSERT11");
    
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

            
            if (!file.exists()) {
                    throw new IOException("Failed to create the report file.");
                }
         

            // Create the PDF file response
            InputStreamResource resource = new InputStreamResource(new FileInputStream(file));

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + reportFilename);
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
