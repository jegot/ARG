package com.example.arg;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelService {

    private static final String FILE_PATH = "C:/Users/jgolling/Desktop/programming/ARG/COUNTER DATASTORE MICROSOFT QUERY.xlsx";
    private static final String[] HEADERS_TO_INCLUDE = {"timestamp", "devicename", "job #", "operator 1", "operator 2", "Job Efficiency %", "count/HR", "target rate/HR", "Run Time Min"};

    public List<String[]> readExcel(String dateFilter, String machineName) throws IOException {
        List<String[]> data = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("M/d/yyyy");

        try (FileInputStream fis = new FileInputStream(FILE_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) { throw new IOException("Header row is missing in the Excel sheet"); }

            int timeColumnIndex = -1;
            int machineColumnIndex = -1;
            int[] columnIndexes = new int[HEADERS_TO_INCLUDE.length];

            for (int i = 0; i < HEADERS_TO_INCLUDE.length; i++) {
                columnIndexes[i] = -1;
            }

            for (Cell cell : headerRow) {
                String header = cell.getStringCellValue();
                for (int i = 0; i < HEADERS_TO_INCLUDE.length; i++) {
                    if (HEADERS_TO_INCLUDE[i].equalsIgnoreCase(header)) {
                        columnIndexes[i] = cell.getColumnIndex();
                        if ("timestamp".equalsIgnoreCase(header)) {
                            timeColumnIndex = cell.getColumnIndex();
                        } else if ("devicename".equalsIgnoreCase(header)) {
                            machineColumnIndex = cell.getColumnIndex();
                        }
                        break;
                    }
                }
            }

            if (timeColumnIndex == -1) { throw new IOException("Timestamp column is missing in the header row"); }
            if (machineColumnIndex == -1) { throw new IOException("Machine column is missing in the header row"); }

            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) { continue; }

                Cell thisCell = row.getCell(timeColumnIndex);
                Cell otherCell = row.getCell(machineColumnIndex);
                if ((thisCell != null && sdf.format(thisCell.getDateCellValue()).equals(dateFilter)) && (otherCell != null && otherCell.getStringCellValue().equals(machineName))) {
                    String[] rowData = new String[HEADERS_TO_INCLUDE.length];

                    for (int i = 0; i < columnIndexes.length; i++) {
                        int cellIndex = columnIndexes[i];
                        if (cellIndex != -1) {
                            Cell cell = row.getCell(cellIndex);
                            rowData[i] = cell == null ? "" : getCellValue(cell);
                        } else {
                            rowData[i] = "";
                        }
                    }
                    data.add(rowData);
                }
            }
        }

        return data;
    }

    private String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new SimpleDateFormat("M/d/yyyy H:mm").format(cell.getDateCellValue());
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "Unknown Cell Type";
        }
    }

    public void generatePdfReport(List<List<String[]>> allData, String outputPath, String date) throws IOException {
        if (allData == null || allData.isEmpty()) {
            throw new IllegalArgumentException("Input data (allData) cannot be null or empty");
        }

        PdfWriter writer = new PdfWriter(outputPath);
        com.itextpdf.kernel.pdf.PdfDocument pdf = new com.itextpdf.kernel.pdf.PdfDocument(writer);
        Document document = new Document(pdf, PageSize.A4);

        try {
            for (List<String[]> data : allData) {
                document.add(new Paragraph("Report for " + date).setFontSize(15).setBold());
                Table table = new Table(HEADERS_TO_INCLUDE.length);

                for (String header : HEADERS_TO_INCLUDE) {
                    table.addCell(header);
                }

                Map<String, List<Double>> operatorEfficiencyMap = new HashMap<>();

                for (String[] rowData : data) {
                    for (String cellData : rowData) {
                        table.addCell(cellData);
                    }

                    String operator1 = rowData[3]; // Assuming "operator 1" is at index 3
                    String operator2 = rowData[4]; // Assuming "operator 2" is at index 4
                    String jobEfficiencyStr = rowData[5]; // Assuming "Job Efficiency %" is at index 5

                    double jobEfficiency = jobEfficiencyStr.isEmpty() ? 0.0 : Double.parseDouble(jobEfficiencyStr);

                    if (!operator1.isEmpty()) {
                        operatorEfficiencyMap.putIfAbsent(operator1, new ArrayList<>());
                        operatorEfficiencyMap.get(operator1).add(jobEfficiency);
                    }

                    if (!operator2.isEmpty()) {
                        operatorEfficiencyMap.putIfAbsent(operator2, new ArrayList<>());
                        operatorEfficiencyMap.get(operator2).add(jobEfficiency);
                    }
                }

                document.add(table);

                // Calculate and add statistics
                for (Map.Entry<String, List<Double>> entry : operatorEfficiencyMap.entrySet()) {
                    String operator = entry.getKey();
                    List<Double> efficiencies = entry.getValue();
                    double averageEfficiency = efficiencies.stream().mapToDouble(Double::doubleValue).average().orElse(0.0);

                    document.add(new Paragraph(String.format("Average Job Efficiency for %s: %.2f%%", operator, averageEfficiency)));
                }

                document.add(new AreaBreak());
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            document.close();
        }
    }
}
