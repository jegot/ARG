package com.example.arg;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelService {

    private static final String FILE_PATH = "D:/Andrews/ARG/COUNTER DATASTORE MICROSOFT QUERY.xlsx";

    public List<String[]> readExcel(String dateFilter) throws IOException {
        List<String[]> data = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("M/d/yyyy");

        try (FileInputStream fis = new FileInputStream(FILE_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                throw new IOException("Header row is missing in the Excel sheet");
            }

            int columnIndex = -1;

            for (Cell cell : headerRow) {
                if ("timestamp".equalsIgnoreCase(cell.getStringCellValue())) {
                    columnIndex = cell.getColumnIndex();
                    break;
                }
            }

            if (columnIndex == -1) {
                throw new IOException("Timestamp column is missing in the header row");
            }

            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }

                Cell thisCell = row.getCell(columnIndex);
                if (thisCell != null && sdf.format(thisCell.getDateCellValue()).equals(dateFilter)) {
                    String[] rowData = new String[row.getLastCellNum()];

                    for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                        Cell cell = row.getCell(cellIndex);
                        rowData[cellIndex] = cell == null ? "" : getCellValue(cell);
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

    public void generatePdfReport(List<String[]> data, String outputPath) throws IOException {
        // Initialize PDF writer
        PdfWriter writer = new PdfWriter(outputPath);

        // Initialize PDF document
        com.itextpdf.kernel.pdf.PdfDocument pdf = new com.itextpdf.kernel.pdf.PdfDocument(writer);

        // Initialize document
        Document document = new Document(pdf);

        // Add title
        document.add(new Paragraph("Filtered Report").setFontSize(15).setBold());

        // Create a table with the same number of columns as the data
        if (data.size() > 0) {
            int numColumns = data.get(0).length;
            Table table = new Table(numColumns);

            // Add data to the table
            for (String[] rowData : data) {
                for (String cellData : rowData) {
                    table.addCell(cellData);
                }
            }

            // Add table to document
            document.add(table);
        }

        // Close document
        document.close();
    }
}
