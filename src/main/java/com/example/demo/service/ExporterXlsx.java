package com.example.demo.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

public class ExporterXlsx<T> {

    private List<String> headers = new ArrayList<>();

    private List<Function<T, String>> functions = new ArrayList<>();

    public void addColumnString(String headerName, Function<T, String> function) {
        headers.add(headerName);
        functions.add(function);
    }

    public void addColumnLong(String headerName, Function<T, Long> function) {
        headers.add(headerName);
        functions.add(function.andThen(integerValue -> integerValue == null ? "" : integerValue.toString()));
    }

    public void addColumnInteger(String headerName, Function<T, Integer> function) {
        headers.add(headerName);
        functions.add(function.andThen(integerValue -> integerValue == null ? "" : integerValue.toString()));
    }

    public void createXlsx(OutputStream outputStream, List<T> objects) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet");

        Row rowHeader = sheet.createRow(0);
        int idxHeaderColumn = 0;
        for (String header : headers) {
            Cell cell = rowHeader.createCell(idxHeaderColumn);
            cell.setCellValue(header);
            idxHeaderColumn++;
        }

        int idxRow = 1;
        for (T object : objects) {
            Row row = sheet.createRow(idxRow);
            int idxDataColumn = 0;
            for (Function<T, String> function : functions) {
                String value = function.apply(object);
                Cell cell = rowHeader.createCell(idxDataColumn);
                cell.setCellValue(value);
                idxDataColumn++;
            }
            idxRow++;
        }
    }

}