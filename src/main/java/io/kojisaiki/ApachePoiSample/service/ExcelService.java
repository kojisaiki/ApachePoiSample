package io.kojisaiki.ApachePoiSample.service;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

@Service
public class ExcelService {

    public String generateSimpleExcel() throws Exception {
        String filepath = "./out.xlsx";
        FileOutputStream fout = null;

        Workbook book = null;

        book = new SXSSFWorkbook();

        Font font = book.createFont();
        font.setFontName("Meiryo UI");
        font.setFontHeightInPoints((short) 9);

        DataFormat format = book.createDataFormat();

        //ヘッダ文字列用のスタイル
        CellStyle style_header = book.createCellStyle();
        style_header.setBorderBottom(BorderStyle.THIN);
        ExcelService.setBorder(style_header, BorderStyle.THIN);
        style_header.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getIndex());
        style_header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style_header.setVerticalAlignment(VerticalAlignment.TOP);
        style_header.setFont(font);

        //文字列用のスタイル
        CellStyle style_string = book.createCellStyle();
        ExcelService.setBorder(style_string, BorderStyle.THIN);
        style_string.setVerticalAlignment(VerticalAlignment.TOP);
        style_string.setFont(font);

        //改行が入った文字列用のスタイル
        CellStyle style_string_wrap = book.createCellStyle();
        ExcelService.setBorder(style_string_wrap, BorderStyle.THIN);
        style_string_wrap.setVerticalAlignment(VerticalAlignment.TOP);
        style_string_wrap.setWrapText(true);
        style_string_wrap.setFont(font);

        //整数用のスタイル
        CellStyle style_int = book.createCellStyle();
        ExcelService.setBorder(style_int, BorderStyle.THIN);
        style_int.setDataFormat(format.getFormat("#,##0;-#,##0"));
        style_int.setVerticalAlignment(VerticalAlignment.TOP);
        style_int.setFont(font);

        //小数用のスタイル
        CellStyle style_double = book.createCellStyle();
        ExcelService.setBorder(style_double, BorderStyle.THIN);
        style_double.setDataFormat(format.getFormat("#,##0.0;-#,##0.0"));
        style_double.setVerticalAlignment(VerticalAlignment.TOP);
        style_double.setFont(font);

        //円表示用のスタイル
        CellStyle style_yen = book.createCellStyle();
        ExcelService.setBorder(style_yen, BorderStyle.THIN);
        style_yen.setDataFormat(format.getFormat("\"\\\"#,##0;\"\\\"-#,##0"));
        style_yen.setVerticalAlignment(VerticalAlignment.TOP);
        style_yen.setFont(font);

        //パーセント表示用のスタイル
        CellStyle style_percent = book.createCellStyle();
        ExcelService.setBorder(style_percent, BorderStyle.THIN);
        style_percent.setDataFormat(format.getFormat("0.0%"));
        style_percent.setVerticalAlignment(VerticalAlignment.TOP);
        style_percent.setFont(font);

        //日時表示用のスタイル
        CellStyle style_datetime = book.createCellStyle();
        ExcelService.setBorder(style_datetime, BorderStyle.THIN);
        style_datetime.setDataFormat(format.getFormat("yyyy/mm/dd hh:mm:ss"));
        style_datetime.setVerticalAlignment(VerticalAlignment.TOP);
        style_datetime.setFont(font);

        Row row;
        int rowNumber;
        Cell cell;
        int colNumber;

        //シートの作成(3シート作ってみる)
        Sheet sheet;

        for (int i = 0; i < 3; i++) {
            sheet = book.createSheet();
            if (sheet instanceof SXSSFSheet) {
                ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
            }

            //シート名称の設定
            book.setSheetName(i, "シート" + (i + 1));

            //ヘッダ行の作成
            rowNumber = 0;
            colNumber = 0;
            row = sheet.createRow(rowNumber);
            cell = row.createCell(colNumber++);
            cell.setCellStyle(style_header);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("No.");

            cell = row.createCell(colNumber++);
            cell.setCellStyle(style_header);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("文字列");

            cell = row.createCell(colNumber++);
            cell.setCellStyle(style_header);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("改行の入った文字列");

            cell = row.createCell(colNumber++);
            cell.setCellStyle(style_header);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("整数");

            cell = row.createCell(colNumber++);
            cell.setCellStyle(style_header);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("小数");

            cell = row.createCell(colNumber++);
            cell.setCellStyle(style_header);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("円");

            cell = row.createCell(colNumber++);
            cell.setCellStyle(style_header);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("パーセント");

            cell = row.createCell(colNumber++);
            cell.setCellStyle(style_header);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("日時");

            cell = row.createCell(colNumber);
            cell.setCellStyle(style_header);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("円(8%の税込)");

            //ウィンドウ枠の固定
            sheet.createFreezePane(1, 1);

            //ヘッダ行にオートフィルタの設定
            sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, colNumber));

            //列幅の自動調整
            for (int j = 0; j <= colNumber; j++) {
                sheet.autoSizeColumn(j, true);
            }

            //データ行の生成(10行作ってみる)
            for (int j = 0; j < 10; j++) {
                rowNumber++;
                colNumber = 0;
                row = sheet.createRow(rowNumber);
                cell = row.createCell(colNumber++);
                cell.setCellStyle(style_int);
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(j + 1);

                cell = row.createCell(colNumber++);
                cell.setCellStyle(style_string);
                cell.setCellType(CellType.STRING);
                cell.setCellValue("これは" + (j + 1) + "行目のデータです。");

                cell = row.createCell(colNumber++);
                cell.setCellStyle(style_string_wrap);
                cell.setCellType(CellType.STRING);
                cell.setCellValue("これは\n" + (j + 1) + "行目の\nデータです。");

                cell = row.createCell(colNumber++);
                cell.setCellStyle(style_int);
                cell.setCellType(CellType.STRING);
                cell.setCellValue((j + 1) * 1000);

                cell = row.createCell(colNumber++);
                cell.setCellStyle(style_double);
                cell.setCellType(CellType.STRING);
                cell.setCellValue((double) (j + 1) * 1000);

                cell = row.createCell(colNumber++);
                cell.setCellStyle(style_yen);
                cell.setCellType(CellType.STRING);
                cell.setCellValue((j + 1) * 1000);

                cell = row.createCell(colNumber++);
                cell.setCellStyle(style_percent);
                cell.setCellType(CellType.STRING);
                cell.setCellValue((double) (j + 1));

                cell = row.createCell(colNumber++);
                cell.setCellStyle(style_datetime);
                cell.setCellType(CellType.STRING);
                cell.setCellValue(new Date());

                cell = row.createCell(colNumber);
                cell.setCellStyle(style_yen);
                cell.setCellType(CellType.FORMULA);
                cell.setCellFormula("ROUND(" + ExcelService.getExcelColumnString(colNumber - 3) + (rowNumber + 1) + "*1.08, 0)");

                //列幅の自動調整
                for (int k = 0; k <= colNumber; k++) {
                    sheet.autoSizeColumn(k, true);
                }
            }
        }

        //シート3を消してみる
        book.removeSheetAt(2);

        fout = new FileOutputStream("./out.xlsx");
        book.write(fout);

        return filepath;
    }

    private static void setBorder(CellStyle style, BorderStyle border) {
        style.setBorderBottom(border);
        style.setBorderTop(border);
        style.setBorderLeft(border);
        style.setBorderRight(border);
    }

    private final static String[] LIST_ALPHA = {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
    };

    private static String getExcelColumnString(int column) {
        String result = "";

        if (column >= 0) {
            if (column / ExcelService.LIST_ALPHA.length > 0) {
                result += getExcelColumnString(column / ExcelService.LIST_ALPHA.length - 1);
            }
            result += ExcelService.LIST_ALPHA[column % ExcelService.LIST_ALPHA.length];
        }

        return result;
    }
}
