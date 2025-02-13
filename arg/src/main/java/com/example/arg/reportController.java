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
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;



@Controller
public class reportController {

    @GetMapping("/")
    public String index() {
        return "index";
    }

    @GetMapping("/generate-report")
    public ResponseEntity<InputStreamResource> generateReport(@RequestParam("date") String date) {

        reportGenerator generator = new reportGenerator();

        LocalDate localDate;

        try {
            localDate = LocalDate.parse(date); // Assumes date is in ISO YYYY-MM-DD format
        } catch (Exception e) {
            throw new IllegalArgumentException("Invalid date format. Expected format: YYYY-MM-DD");
        }

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
        String formattedDate = localDate.format(formatter);

        String reportFilename = "Report_" + date + ".pdf";
        String reportPath = "" + reportFilename;

        try {
            generator.generatePdf(reportPath, formattedDate);
        } catch (Exception e) {
            System.out.println("An error occurred while generating the report PDF");
            e.printStackTrace();
            return ResponseEntity.internalServerError().build();
        }


        // Attempt to create and return the response with the new PDF file as an attachment
        try {
            File file = new File(reportPath);
            InputStreamResource resource = new InputStreamResource(new FileInputStream(file));
            
            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + reportFilename);

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentLength(file.length())
                    .contentType(MediaType.APPLICATION_PDF)
                    .body(resource);
        } catch (FileNotFoundException e) {
            System.out.println("Generated file not found when trying to create response");
            e.printStackTrace();
            return ResponseEntity.notFound().build();
        } catch (Exception e) {
            System.out.println("An unexpected error occurred");
            e.printStackTrace();
            return ResponseEntity.internalServerError().build();
        }
    
    }



    @GetMapping("/generate-job-specific-report")
    public ResponseEntity<InputStreamResource> generateJobSpecificReport(@RequestParam("jobNumber") String jobNumber) {

        System.out.println("attempting to create job report");

        reportGenerator generator = new reportGenerator();

        String reportFilename = "Report_" + jobNumber + ".pdf";
        String reportPath = "Y:/Machine Reports/NewReports/" + reportFilename;

        try {
            generator.generateJobReport(reportPath, jobNumber);
        } catch (Exception e) {
            System.out.println("An error occurred while generating the report PDF");
            e.printStackTrace();
            return ResponseEntity.internalServerError().build();
        }


        // Attempt to create and return the response with the new PDF file as an attachment
        try {
            File file = new File(reportPath);
            InputStreamResource resource = new InputStreamResource(new FileInputStream(file));
            
            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + reportFilename);

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentLength(file.length())
                    .contentType(MediaType.APPLICATION_PDF)
                    .body(resource);
        } catch (FileNotFoundException e) {
            System.out.println("Generated file not found when trying to create response");
            e.printStackTrace();
            return ResponseEntity.notFound().build();
        } catch (Exception e) {
            System.out.println("An unexpected error occurred");
            e.printStackTrace();
            return ResponseEntity.internalServerError().build();
        }
    
    }



}
