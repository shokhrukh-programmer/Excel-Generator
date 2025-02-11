package org.example.excelgenerator.fonts;

import org.apache.poi.ss.usermodel.*;

public class Styles {
    public static CellStyle getItalicStyle(Workbook workbook) {
        // 🔴 Create italic bold font (For specific rows & columns)
        Font italicBoldFont = workbook.createFont();
        italicBoldFont.setFontName("Times New Roman");
        italicBoldFont.setBold(true);
        italicBoldFont.setItalic(true);
        italicBoldFont.setFontHeightInPoints((short) 18);

        // 🔴 Create italic cell style
        CellStyle italicStyle = workbook.createCellStyle();
        italicStyle.setFont(italicBoldFont);
        italicStyle.setAlignment(HorizontalAlignment.CENTER);
        italicStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        italicStyle.setBorderBottom(BorderStyle.THIN);
        italicStyle.setBorderTop(BorderStyle.THIN);
        italicStyle.setBorderLeft(BorderStyle.THIN);
        italicStyle.setBorderRight(BorderStyle.THIN);

        return italicStyle;
    }

    public static CellStyle getCellStyle(Workbook workbook) {
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Times New Roman");
        boldFont.setBold(true);
        boldFont.setFontHeightInPoints((short) 22);

        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFont(boldFont);
        boldStyle.setAlignment(HorizontalAlignment.CENTER);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.THIN);
        boldStyle.setBorderLeft(BorderStyle.THIN);
        boldStyle.setBorderRight(BorderStyle.THIN);
        return boldStyle;
    }

    public static CellStyle getBottomBorder(Workbook workbook) {
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Times New Roman");
        boldFont.setBold(true);
        boldFont.setFontHeightInPoints((short) 22);

        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFont(boldFont);
        boldStyle.setAlignment(HorizontalAlignment.CENTER);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.NONE);
        boldStyle.setBorderLeft(BorderStyle.NONE);
        boldStyle.setBorderRight(BorderStyle.NONE);
        return boldStyle;
    }
}
