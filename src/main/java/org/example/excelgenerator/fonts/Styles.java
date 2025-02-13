package org.example.excelgenerator.fonts;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
        italicStyle.setWrapText(true);
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
        boldStyle.setWrapText(true);
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
        boldStyle.setWrapText(true);
        boldStyle.setFont(boldFont);
        boldStyle.setAlignment(HorizontalAlignment.CENTER);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.NONE);
        boldStyle.setBorderLeft(BorderStyle.NONE);
        boldStyle.setBorderRight(BorderStyle.NONE);
        return boldStyle;
    }

    public static CellStyle getCellBasicStyle(Workbook workbook) {
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Times New Roman");
        boldFont.setBold(false);
        boldFont.setFontHeightInPoints((short) 22);

        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setWrapText(true);
        boldStyle.setFont(boldFont);
        boldStyle.setAlignment(HorizontalAlignment.CENTER);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.THIN);
        boldStyle.setBorderLeft(BorderStyle.THIN);
        boldStyle.setBorderRight(BorderStyle.THIN);
        return boldStyle;
    }

    public static CellStyle getCellBasicStyleWithBackgroundGreen(Workbook workbook) {
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Times New Roman");
        boldFont.setBold(false);
        boldFont.setFontHeightInPoints((short) 22);

        XSSFColor color2 = new XSSFColor(new byte[]{(byte) 198, (byte) 224, (byte) 180}, null);
        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFillForegroundColor(color2);

        boldStyle.setFont(boldFont);
        boldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 🔵 Create cell style with bold font
        boldStyle.setWrapText(true);
        boldStyle.setFont(boldFont);
        boldStyle.setAlignment(HorizontalAlignment.CENTER);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.THIN);
        boldStyle.setBorderLeft(BorderStyle.THIN);
        boldStyle.setBorderRight(BorderStyle.THIN);
        return boldStyle;
    }

    public static CellStyle getLeftCellStyle(Workbook workbook) {
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Times New Roman");
        boldFont.setBold(true);
        boldFont.setFontHeightInPoints((short) 22);

        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setWrapText(true);
        boldStyle.setFont(boldFont);
        boldStyle.setAlignment(HorizontalAlignment.LEFT);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.THIN);
        boldStyle.setBorderLeft(BorderStyle.THIN);
        boldStyle.setBorderRight(BorderStyle.THIN);
        return boldStyle;
    }

    public static CellStyle getLeftCellStyleBlue(Workbook workbook) {
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Times New Roman");
        boldFont.setBold(true);
        boldFont.setFontHeightInPoints((short) 22);

        XSSFColor color2 = new XSSFColor(new byte[]{(byte) 180, (byte) 198, (byte) 231}, null);
        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFillForegroundColor(color2);

        boldStyle.setFont(boldFont);
        boldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 🔵 Create cell style with bold font
        boldStyle.setAlignment(HorizontalAlignment.LEFT);
        boldStyle.setWrapText(true);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.THIN);
        boldStyle.setBorderLeft(BorderStyle.THIN);
        boldStyle.setBorderRight(BorderStyle.THIN);
        return boldStyle;
    }

    public static CellStyle getBackgroundAndText(Workbook workbook) {
        XSSFFont font = (XSSFFont) workbook.createFont();
        byte[] fontColor = new byte[]{(byte) 0, (byte) 112, (byte) 192}; // Red color
        XSSFColor fontXssfColor = new XSSFColor(fontColor, null);
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 22);
        font.setColor(fontXssfColor);


        XSSFColor color2 = new XSSFColor(new byte[]{(byte) 248, (byte) 203, (byte) 173}, null);
        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFillForegroundColor(color2);

        boldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        boldStyle.setWrapText(true);
        boldStyle.setFont(font);
        boldStyle.setAlignment(HorizontalAlignment.CENTER);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.THIN);
        boldStyle.setBorderLeft(BorderStyle.THIN);
        boldStyle.setBorderRight(BorderStyle.THIN);
        return boldStyle;
    }

    public static CellStyle getBackground(Workbook workbook) {
        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 22);


        XSSFColor color2 = new XSSFColor(new byte[]{(byte) 248, (byte) 203, (byte) 173}, null);
        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFillForegroundColor(color2);

        boldStyle.setWrapText(true);
        boldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        boldStyle.setFont(font);
        boldStyle.setAlignment(HorizontalAlignment.CENTER);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.THIN);
        boldStyle.setBorderLeft(BorderStyle.THIN);
        boldStyle.setBorderRight(BorderStyle.THIN);

        return boldStyle;
    }

    public static CellStyle getBackgroundBlue(Workbook workbook) {
        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 22);


        XSSFColor color2 = new XSSFColor(new byte[]{(byte) 180, (byte) 198, (byte) 231}, null);
        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFillForegroundColor(color2);

        boldStyle.setWrapText(true);
        boldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        boldStyle.setFont(font);
        boldStyle.setAlignment(HorizontalAlignment.CENTER);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.THIN);
        boldStyle.setBorderLeft(BorderStyle.THIN);
        boldStyle.setBorderRight(BorderStyle.THIN);

        return boldStyle;
    }

    public static CellStyle getBackgroundBlueWithoutBorder(Workbook workbook) {
        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 22);


        XSSFColor color2 = new XSSFColor(new byte[]{(byte) 180, (byte) 198, (byte) 231}, null);
        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFillForegroundColor(color2);
        boldStyle.setWrapText(true);

        boldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        boldStyle.setFont(font);
        boldStyle.setAlignment(HorizontalAlignment.CENTER);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.NONE);
        boldStyle.setBorderTop(BorderStyle.NONE);
        boldStyle.setBorderLeft(BorderStyle.NONE);
        boldStyle.setBorderRight(BorderStyle.NONE);

        return boldStyle;
    }

    public static CellStyle getItalicStyleWithRed(Workbook workbook) {
        // 🔴 Create italic bold font (For specific rows & columns)
        Font italicBoldFont = workbook.createFont();
        italicBoldFont.setFontName("Times New Roman");
        italicBoldFont.setBold(true);
        italicBoldFont.setItalic(true);
        italicBoldFont.setFontHeightInPoints((short) 18);
        italicBoldFont.setColor(IndexedColors.RED.getIndex());

        // 🔴 Create italic cell style
        CellStyle italicStyle = workbook.createCellStyle();
        italicStyle.setWrapText(true);
        italicStyle.setFont(italicBoldFont);
        italicStyle.setAlignment(HorizontalAlignment.CENTER);
        italicStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        italicStyle.setBorderBottom(BorderStyle.THIN);
        italicStyle.setBorderTop(BorderStyle.THIN);
        italicStyle.setBorderLeft(BorderStyle.THIN);
        italicStyle.setBorderRight(BorderStyle.THIN);

        return italicStyle;
    }

    public static CellStyle getCellStyleRed(XSSFWorkbook workbook) {
        Font italicBoldFont = workbook.createFont();
        italicBoldFont.setFontName("Times New Roman");
        italicBoldFont.setBold(false);
        italicBoldFont.setItalic(false);
        italicBoldFont.setFontHeightInPoints((short) 18);
        italicBoldFont.setColor(IndexedColors.RED.getIndex());

        // 🔴 Create italic cell style
        CellStyle italicStyle = workbook.createCellStyle();
        italicStyle.setWrapText(true);
        italicStyle.setFont(italicBoldFont);
        italicStyle.setAlignment(HorizontalAlignment.CENTER);
        italicStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        italicStyle.setBorderBottom(BorderStyle.THIN);
        italicStyle.setBorderTop(BorderStyle.THIN);
        italicStyle.setBorderLeft(BorderStyle.THIN);
        italicStyle.setBorderRight(BorderStyle.THIN);

        return italicStyle;
    }

    public static CellStyle getCellStyleRedSize(XSSFWorkbook workbook) {
        Font italicBoldFont = workbook.createFont();
        italicBoldFont.setFontName("Times New Roman");
        italicBoldFont.setBold(false);
        italicBoldFont.setItalic(false);
        italicBoldFont.setFontHeightInPoints((short) 22);
        italicBoldFont.setColor(IndexedColors.RED.getIndex());

        // 🔴 Create italic cell style
        CellStyle italicStyle = workbook.createCellStyle();
        italicStyle.setWrapText(true);
        italicStyle.setFont(italicBoldFont);
        italicStyle.setAlignment(HorizontalAlignment.CENTER);
        italicStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        italicStyle.setBorderBottom(BorderStyle.THIN);
        italicStyle.setBorderTop(BorderStyle.THIN);
        italicStyle.setBorderLeft(BorderStyle.THIN);
        italicStyle.setBorderRight(BorderStyle.THIN);

        return italicStyle;
    }

    public static CellStyle getCellBasicStyleWithBackgroundGreenBolt(XSSFWorkbook workbook) {
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Times New Roman");
        boldFont.setBold(true);
        boldFont.setFontHeightInPoints((short) 22);

        XSSFColor color2 = new XSSFColor(new byte[]{(byte) 198, (byte) 224, (byte) 180}, null);
        // 🔵 Create cell style with bold font
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFillForegroundColor(color2);

        boldStyle.setFont(boldFont);
        boldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 🔵 Create cell style with bold font
        boldStyle.setWrapText(true);
        boldStyle.setFont(boldFont);
        boldStyle.setAlignment(HorizontalAlignment.CENTER);
        boldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        boldStyle.setBorderBottom(BorderStyle.THIN);
        boldStyle.setBorderTop(BorderStyle.THIN);
        boldStyle.setBorderLeft(BorderStyle.THIN);
        boldStyle.setBorderRight(BorderStyle.THIN);
        return boldStyle;
    }
}
