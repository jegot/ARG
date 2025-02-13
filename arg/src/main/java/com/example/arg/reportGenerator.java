package com.example.arg;

import java.util.List;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Paragraph;

import java.sql.*;
import java.util.ArrayList;

public class reportGenerator {
    public static String[] HEADERS_TO_INCLUDE = {
        "timestamp", "devicename", "job #", "operator 1", 
        "operator 2", "job count", "%/HR", "Job Efficiency %", 
        "count/HR", "target rate/HR", "Run Time Min", "job avg rate"
    };

    private DatabaseConnector dbConnector;

    public reportGenerator() {
        dbConnector = new DatabaseConnector();
    }

    public static List<String> fetchUniqueMachines(String dateFilter, Connection conn) throws SQLException {
        List<String> machineNames = new ArrayList<>();
        String query = "SELECT DISTINCT devicename FROM tbl_datapool WHERE DATE(timestamp) = STR_TO_DATE(?, '%m/%d/%Y')";
        
        try (PreparedStatement pstmt = conn.prepareStatement(query)) {
            pstmt.setString(1, dateFilter);
            ResultSet rs = pstmt.executeQuery();
            while (rs.next()) {
                machineNames.add(rs.getString("devicename"));
            }
        } catch (SQLException e) {
            System.err.println("SQL Exception: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    
        return machineNames;
    }





    public static List<String> fetchUniqueOperators(String dateFilter, Connection conn, String machineName) throws SQLException {
        List<String> operators = new ArrayList<>();
        String queryOperator1 = "SELECT DISTINCT `operator 1` FROM tbl_datapool WHERE DATE(timestamp) = STR_TO_DATE(?, '%m/%d/%Y') AND devicename = ?";
        String queryOperator2 = "SELECT DISTINCT `operator 2` FROM tbl_datapool WHERE DATE(timestamp) = STR_TO_DATE(?, '%m/%d/%Y') AND devicename = ?";
    
        try (PreparedStatement pstmt1 = conn.prepareStatement(queryOperator1);
             PreparedStatement pstmt2 = conn.prepareStatement(queryOperator2)) {
    
            
            pstmt1.setString(1, dateFilter);
            pstmt1.setString(2, machineName);
            
            // Execute the first query
            try (ResultSet rs1 = pstmt1.executeQuery()) {
                while (rs1.next()) {
                    String operator1 = rs1.getString("operator 1");
                    if (operator1 != null) { 
                        operators.add(operator1);
                    }
                }
            }
    
            
            pstmt2.setString(1, dateFilter);
            pstmt2.setString(2, machineName);
            
            // Execute the second query
            try (ResultSet rs2 = pstmt2.executeQuery()) {
                while (rs2.next()) {
                    String operator2 = rs2.getString("operator 2");
                    if (operator2 != null && !operators.contains(operator2)) { // Avoid duplicates
                        operators.add(operator2);
                    }
                }
            }
    
        } catch (SQLException e) {
            System.err.println("SQL Exception: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    
        return operators;
    }
    



    
    public static List<String[]> fetchData(String dateFilter, String machineName, Connection conn) throws SQLException {
        List<String[]> data = new ArrayList<>();
    
        String query = "SELECT * FROM tbl_datapool WHERE DATE(timestamp) = STR_TO_DATE(?, '%m/%d/%Y') AND devicename = ?";
        try (PreparedStatement pstmt = conn.prepareStatement(query)) {
            pstmt.setString(1, dateFilter);
            pstmt.setString(2, machineName);
    
            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    String[] row = new String[HEADERS_TO_INCLUDE.length];
                    for (int i = 0; i < HEADERS_TO_INCLUDE.length; i++) {
                        row[i] = rs.getString(HEADERS_TO_INCLUDE[i]);
                    }
                    data.add(row);
                }
            }
        } catch (SQLException e) {
            System.err.println("SQL Exception for machine: " + machineName + " - " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    
        return data;
    }




    public void generatePdf(String dest, String dateFilter) throws Exception {
        dbConnector.connect(); 
        Connection conn = null;
        
        try {
            // Open the connection here and pass it down
            conn = DatabaseConnector.getConnection();
    
            PdfWriter writer = new PdfWriter(dest);
            PdfDocument pdf = new PdfDocument(writer);

            PageSize ltrLandscape = PageSize.LETTER.rotate();
            Document document = new Document(pdf, ltrLandscape);
            
            
            List<String> machineNames = fetchUniqueMachines(dateFilter, conn);
    
            
            for (String machineName : machineNames) {
                
                List<String[]> dataForMachine = fetchData(dateFilter, machineName, conn);
    
                // Create a new table for the current machine
                Table table = new Table(HEADERS_TO_INCLUDE.length);
                
                // Add headers to the table
                for (String header : HEADERS_TO_INCLUDE) {
                    table.addCell(header);
                }
    
                // Populate table rows with data
                for (String[] rowData : dataForMachine) {
                    for (String cellData : rowData) {
                        table.addCell(cellData);
                    }
                }
    
                document.add(new Paragraph("Report for " + machineName + " on " + dateFilter));
                document.add(table.setFontSize(10));
                

                //***** GENERATE MACHINE SUMMARY HERE AFTER TABLE */
                List<String>allOperators = fetchUniqueOperators(dateFilter, conn, machineName);

                for (String operator : allOperators) {
                    if(operator != null && operator != "") {
                        document.add(new Paragraph(generateOperatorStatLine(dateFilter, conn, machineName, operator)).setFontSize(11));
                    }
                }

                
                // New page begins here
                document.add(new AreaBreak());
            }

            //delete last page
            if (pdf.getNumberOfPages() > 0) {
                pdf.removePage(pdf.getNumberOfPages());
            }

            // Close the document
            document.close();
        } finally {
            if (conn != null && !conn.isClosed()) {
                conn.close();  
            }
        }
    }



    public static String generateOperatorStatLine(String dateFilter, Connection conn, String machineName, String operator) throws SQLException {
        StringBuilder operatorLine = new StringBuilder();
    
        String query = "SELECT `job #`, " +
                       "MIN(`job count`) AS begin_count, " + // First count at the earliest timestamp
                       "MAX(`job count`) AS end_count, " +   // Last count at the latest timestamp
                       "MAX(`job count`) - MIN(`job count`) AS difference, " + // Difference between last and first counts
                       "MIN(`target rate/HR`) AS min_target_rate_per_hour, " + 
                       "MAX(`Run Time Min`) AS max_run_time_min, " + // Max run time min
                       "MIN(`Run Time Min`) AS min_run_time_min " + // Min run time min
                       "FROM tbl_datapool " +
                       "WHERE DATE(timestamp) = STR_TO_DATE(?, '%m/%d/%Y') AND devicename = ? AND (`operator 1` = ? OR `operator 2` = ?) " +
                       "GROUP BY `job #`"; 
                       
    
        try (PreparedStatement pstmt = conn.prepareStatement(query)) {
            pstmt.setString(1, dateFilter);
            pstmt.setString(2, machineName);
            pstmt.setString(3, operator);
            pstmt.setString(4, operator);
            
            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    String jobNumber = rs.getString("job #");
                    int beginCount = rs.getInt("begin_count");
                    int endCount = rs.getInt("end_count");
                    int difference = rs.getInt("difference");
                    //double avgCountPerHour = rs.getDouble("avg_count_per_hour");
                    int minTargetPerHour = rs.getInt("min_target_rate_per_hour");
                    int maxRunTime = rs.getInt("max_run_time_min");
                    int minRunTime = rs.getInt("min_run_time_min");
                


                    int totalJobTimeInMins = maxRunTime - minRunTime;
                    double totalJobTimeInHrs = totalJobTimeInMins/60.0;

                    //default value to zero as a fallback
                    double avgCountPerHour = 0.0;

                    //"avg Percent per hour" calculated by avg count/hr divided by target rate/hr * 100
                    double avgPercentPerHour = 0.0;

                    //insure no division by 0 occurs, will throw error
                    if(totalJobTimeInHrs > 0) {
                        avgCountPerHour = difference/totalJobTimeInHrs;
                    }

                    if( minTargetPerHour > 0) {
                        avgPercentPerHour = avgCountPerHour/minTargetPerHour;
                        avgPercentPerHour = avgPercentPerHour * 100.00;
                    }





                    if(jobNumber != null && operator != null && !operator.isEmpty() && difference > 0) {
                        operatorLine.append(String.format("Operator: %s, Job #: %s\n" +
                                                       "Avg Count/HR: %.0f, Difference: %d (Begin: %d to End: %d)    Avg Percent/HR: %.2f%%, Operator Time: %.2fhrs ( %d mins)\n",
                            operator, jobNumber, avgCountPerHour, difference, beginCount, endCount, avgPercentPerHour, totalJobTimeInHrs, totalJobTimeInMins));
                    }
                    
                }
            }
        } catch (SQLException e) {
            System.err.println("SQL Exception: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    
        return operatorLine.toString();
    }
    



    public void generateMonthlyPdf(String dest, String month, String year) throws Exception {
        dbConnector.connect(); 
        Connection conn = null;
    
        try {
          
            String numericMonth = convertMonthToNumber(month);
    
            // Open the connection here and pass it down
            conn = DatabaseConnector.getConnection();
    
            PdfWriter writer = new PdfWriter(dest);
            PdfDocument pdf = new PdfDocument(writer);
    
            PageSize ltrLandscape = PageSize.LETTER.rotate();
            Document document = new Document(pdf, ltrLandscape);
    
            document.add(new Paragraph("Report for " + month + " " + year));
    
            // 1. Query for operators and their average count/HR
            String queryOperators = "SELECT `operator 1` AS operator, AVG(`count/HR`) AS avg_count_hr " +
                                    "FROM tbl_datapool " +
                                    "WHERE YEAR(timestamp) = ? AND MONTH(timestamp) = ? " +
                                    "GROUP BY `operator 1` HAVING operator IS NOT NULL";
    
            PreparedStatement pstmtOperators = conn.prepareStatement(queryOperators);
            pstmtOperators.setString(1, year);
            pstmtOperators.setString(2, numericMonth);
            ResultSet rsOperators = pstmtOperators.executeQuery();
    

            document.add(new Paragraph("Operators and their Average Count/HR:"));
            while (rsOperators.next()) {
                String operator = rsOperators.getString("operator");
                double avgCountHR = rsOperators.getDouble("avg_count_hr");
                document.add(new Paragraph("Operator: " + operator + ", Avg Count/HR: " + String.format("%.2f", avgCountHR)));
            }
    
            
            
            document.add(new AreaBreak());


            // 2. Query for machines and their average job efficiency
            String queryMachines = "SELECT `devicename`, AVG(`Job Efficiency %`) AS avg_job_efficiency " +
                                   "FROM tbl_datapool " +
                                   "WHERE YEAR(timestamp) = ? AND MONTH(timestamp) = ? " +
                                   "GROUP BY `devicename` HAVING devicename IS NOT NULL";
    
            PreparedStatement pstmtMachines = conn.prepareStatement(queryMachines);
            pstmtMachines.setString(1, year);
            pstmtMachines.setString(2, numericMonth);
            ResultSet rsMachines = pstmtMachines.executeQuery();
    
            document.add(new Paragraph("Machines and their Average Job Efficiency %:"));
            while (rsMachines.next()) {
                String deviceName = rsMachines.getString("devicename");
                double avgJobEfficiency = rsMachines.getDouble("avg_job_efficiency");
                document.add(new Paragraph("Device: " + deviceName + ", Avg Job Efficiency %: " + String.format("%.2f", avgJobEfficiency)));
            }

            document.add(new AreaBreak());

    
            // 3. Query for jobs and their max Run Time Min
            String queryJobs = "SELECT `job #`, MAX(`Run Time Min`) AS max_run_time " +
                               "FROM tbl_datapool " +
                               "WHERE YEAR(timestamp) = ? AND MONTH(timestamp) = ? " +
                               "GROUP BY `job #` HAVING `job #` IS NOT NULL";
    
            PreparedStatement pstmtJobs = conn.prepareStatement(queryJobs);
            pstmtJobs.setString(1, year);
            pstmtJobs.setString(2, numericMonth);
            ResultSet rsJobs = pstmtJobs.executeQuery();
    

            document.add(new AreaBreak());

            document.add(new Paragraph("Jobs and their Maximum Run Time:"));
            while (rsJobs.next()) {
                String jobNumber = rsJobs.getString("job #");
                int maxRunTime = rsJobs.getInt("max_run_time");
                if(maxRunTime < 0) {
                    document.add(new Paragraph("Job #: " + jobNumber + ", Max Run Time Min: " + maxRunTime));
                }
            }
    

            // Close the document
            document.close();
        } finally {
            if (conn != null && !conn.isClosed()) {
                conn.close();  
            }
        }
    }
    


    public static String convertMonthToNumber(String monthName) {
        switch (monthName.toLowerCase()) {
            case "january": return "01";
            case "february": return "02";
            case "march": return "03";
            case "april": return "04";
            case "may": return "05";
            case "june": return "06";
            case "july": return "07";
            case "august": return "08";
            case "september": return "09";
            case "october": return "10";
            case "november": return "11";
            case "december": return "12";
            default: throw new IllegalArgumentException("Invalid month: " + monthName);
        }
    }
    
    
    public void generateJobReport(String dest, String jobNumber) throws Exception {
        dbConnector.connect();
        Connection conn = null;

        int totalNumOfOperators = 1;
    
        try {
            conn = DatabaseConnector.getConnection(); // Connect to DB
           
            PdfWriter writer = new PdfWriter(dest);
            PdfDocument pdf = new PdfDocument(writer);

            PageSize ltrLandscape = PageSize.LETTER.rotate();
            Document document = new Document(pdf, ltrLandscape);

    
            // Header for the job-specific report
            document.add(new Paragraph("Job Report for: " + jobNumber).setBold().setFontSize(14));
    
            // Query to fetch data for the given job number (case-insensitive match)
            String query = "SELECT timestamp, devicename, `operator 1`, `operator 2`, `job count`, `Run Time Min`, `count/HR`, `rate/HR` " +
                           "FROM tbl_datapool WHERE `job #` LIKE ? COLLATE utf8mb4_general_ci";
    
            try (PreparedStatement pstmt = conn.prepareStatement(query)) {
                pstmt.setString(1, "%" + jobNumber + "%"); // Use LIKE for partial matches
                ResultSet rs = pstmt.executeQuery();
    
                // Create a table with columns for the data
                Table table = new Table(9); // 
                table.addCell("Timestamp");
                table.addCell("Device");
                table.addCell("Operator 1");
                table.addCell("Operator 2");
                table.addCell("Job Count");
                table.addCell("Run Time (Min)");
                table.addCell("Pieces per Hour");
                table.addCell("Rate per Hour");
                table.addCell("Number of operators");
    
                // Populate the table with the data for each device run
                boolean dataFound = false;
                
                while (rs.next()) {
                    dataFound = true;
                    table.addCell(rs.getString("timestamp"));
                    table.addCell(rs.getString("devicename"));
                    table.addCell(rs.getString("operator 1"));
                    table.addCell(rs.getString("operator 2"));
                    table.addCell(String.valueOf(rs.getInt("job count")));
                    table.addCell(String.valueOf(rs.getInt("Run Time Min")));
                    table.addCell(String.format("%.0f", rs.getDouble("count/HR")));
                    table.addCell(String.format("%.0f", rs.getDouble("rate/HR")));


                    if (rs.getString("operator 2") != null && !rs.getString("operator 2").isEmpty()) {
                        totalNumOfOperators = 2;
                        table.addCell("2");
                    } else {
                        totalNumOfOperators = 1;
                        table.addCell("1");
                    }
                    
                }
    
                if (!dataFound) {
                    document.add(new Paragraph("No data found for Job Number: " + jobNumber).setFontSize(12));
                } else {
                    // Add the table to the PDF document
                    document.add(table.setFontSize(9));
                }
            }
    
            // Query for total statistics across all devices for this job, including the latest job avg rate
            String summaryQuery = "SELECT MAX(`job count`) AS total_pieces, " +
                                  "MAX(`Run Time Min`) AS total_run_time, " +
                                  "(SELECT `job avg rate` FROM tbl_datapool " +
                                  " WHERE `job #` LIKE ? COLLATE utf8mb4_general_ci " +
                                  " ORDER BY timestamp DESC LIMIT 1) AS latest_avg_pieces_per_hour " +
                                  "FROM tbl_datapool WHERE `job #` LIKE ? COLLATE utf8mb4_general_ci";
    
            try (PreparedStatement summaryStmt = conn.prepareStatement(summaryQuery)) {
                summaryStmt.setString(1, "%" + jobNumber + "%");
                summaryStmt.setString(2, "%" + jobNumber + "%");
                ResultSet summaryRs = summaryStmt.executeQuery();
    
                if (summaryRs.next()) {
                    int totalPieces = summaryRs.getInt("total_pieces");
                    int totalRunTime = summaryRs.getInt("total_run_time");
                    double latestAvgPiecesPerHour = summaryRs.getDouble("latest_avg_pieces_per_hour");

                    int totalOperatorHours = totalRunTime * totalNumOfOperators;
    
                    // Add the summary to the document
                    document.add(new Paragraph("\nSummary").setBold().setFontSize(12));
                    document.add(new Paragraph("Total Pieces: " + totalPieces));
                    document.add(new Paragraph("Total Run Time (Min): " + totalRunTime + "     //    Estimated Operator Minutes: " + totalOperatorHours));
                    document.add(new Paragraph("Average Pieces/Hour: " + String.format("%.2f", latestAvgPiecesPerHour)));


                }
            }
    
            // Close the document
            document.close();
            System.out.println("Job Report generated successfully: " + dest);
    
        } catch (SQLException e) {
            System.err.println("SQL Exception: " + e.getMessage());
            e.printStackTrace();
            throw e;
        } finally {
            if (conn != null && !conn.isClosed()) {
                conn.close(); // Close DB connection
            }
        }
    }
    
    
    
}
