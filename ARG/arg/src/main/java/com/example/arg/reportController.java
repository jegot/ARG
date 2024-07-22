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
import java.util.List;

@Controller
public class reportController {

    @Autowired
    private ExcelService excelService;

    @GetMapping("/")
    public String index() {
        return "index";
    }

    private String reportPath = "C:\\Users\\jgoll\\Desktop\\mmtest\\thisReport.pdf";

    @GetMapping("/generate-report")
    public ResponseEntity<InputStreamResource> generateReport(@RequestParam("date") String date) {
        try {
            // Read data from the Excel file for the specified date
            List<String[]> data = excelService.readExcel(date);

            // Generate the PDF report
            excelService.generatePdfReport(data, reportPath);

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
