package me.choicore.samples;

import lombok.Data;
import me.choicore.samples.excel.Column;
import me.choicore.samples.excel.Exporters;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

public class Application {

    public static void main(String[] args) {
        Record record = new Record();
        record.setStringColumn("문자");
        record.setDateColumn(LocalDate.now());
        record.setDateTimeColumn(LocalDateTime.now());
        record.setIntColumn(1);
        record.setDoubleColumn(1.1);
        record.setLongColumn(1L);
        record.setBigDecimalColumn(new java.math.BigDecimal("1.1"));
        record.setBooleanColumn(true);
        try (Workbook workbook = Exporters.excel(List.of(record), Record.class)) {
            workbook.write(new FileOutputStream("./record_" + System.currentTimeMillis() + ".xlsx"));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Data
    public static class Record {
        @Column(header = "문자", width = 10, align = HorizontalAlignment.CENTER)
        private String stringColumn;
        @Column(header = "날짜", style = Column.Style.DATE, align = HorizontalAlignment.CENTER)
        private LocalDate dateColumn;
        @Column(header = "날짜 + 시간", style = Column.Style.DATE)
        private LocalDateTime dateTimeColumn;
        @Column(header = "정수", style = Column.Style.NUMBER)
        private int intColumn;
        @Column(header = "실수", style = Column.Style.NUMBER)
        private double doubleColumn;
        @Column(header = "긴 정수", style = Column.Style.NUMBER)
        private long longColumn;
        @Column(header = "실수", style = Column.Style.NUMBER)
        private BigDecimal bigDecimalColumn;
        @Column(header = "부울", style = Column.Style.NORMAL)
        private boolean booleanColumn;
    }

}
