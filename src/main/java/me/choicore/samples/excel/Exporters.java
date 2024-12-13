package me.choicore.samples.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;

public class Exporters {
    public static <T> Workbook excel(List<T> data, Class<T> clazz) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        List<Field> fields = getAnnotatedFields(clazz);

        Row headerRow = sheet.createRow(0);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        for (int i = 0; i < fields.size(); i++) {
            Cell cell = headerRow.createCell(i);
            Column column = fields.get(i).getDeclaredAnnotation(Column.class);
            cell.setCellValue(column.header());
            cell.setCellStyle(headerStyle);

            if (column.width() > 0) {
                sheet.setColumnWidth(i, column.width() * 256);
            }
        }

        if (data.isEmpty()) {
            noContent(workbook, sheet, fields);
            return workbook;
        }

        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            T item = data.get(i);

            for (int j = 0; j < fields.size(); j++) {
                Cell cell = row.createCell(j);
                Field field = fields.get(j);
                Object value = field.get(item);
                Column column = field.getDeclaredAnnotation(Column.class);
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setAlignment(column.align());

                switch (column.style()) {
                    case DATE -> {
                        if (value instanceof LocalDate date) {
                            cell.setCellValue(date);
                            cellStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd"));


                        } else if (value instanceof LocalDateTime dateTime) {
                            cell.setCellValue(dateTime);
                            cellStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));

                        }
                    }
                    case NUMBER -> {
                        if (value instanceof Number number) {
                            cell.setCellValue(number.doubleValue());
                        }
                    }
                    default -> cell.setCellValue(value != null ? value.toString() : "");
                }
                cell.setCellStyle(cellStyle);
            }
        }

        for (int i = 0; i < fields.size(); i++) {
            Column column = fields.get(i).getDeclaredAnnotation(Column.class);
            if (column.width() < 0) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, (int) (sheet.getColumnWidth(i) * 1.2));
            }
        }

        return workbook;
    }

    private static <T> List<Field> getAnnotatedFields(Class<T> clazz) {
        List<Field> fields = Arrays.stream(clazz.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(Column.class))
                .peek(field -> field.setAccessible(true))
                .toList();
        return fields;
    }

    private static void noContent(Workbook workbook, Sheet sheet, List<Field> fields) {
        CellStyle noDataStyle = workbook.createCellStyle();
        noDataStyle.setAlignment(HorizontalAlignment.CENTER);
        noDataStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Row noDataRow = sheet.createRow(1);
        Cell noDataCell = noDataRow.createCell(0);
        noDataCell.setCellValue("데이터가 없습니다.");
        noDataCell.setCellStyle(noDataStyle);

        sheet.addMergedRegion(new CellRangeAddress(
                1,
                1,
                0,
                fields.size() - 1
        ));
    }

}