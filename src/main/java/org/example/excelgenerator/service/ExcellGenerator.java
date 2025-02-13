package org.example.excelgenerator.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.excelgenerator.dto.request.ExcelRequest;
import org.example.excelgenerator.fonts.Styles;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

@Service
public class ExcellGenerator {
    public byte[] generateExcel(ExcelRequest excelRequest) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Data");
        //ExcelRequest excelRequest = new ExcelRequest();
        int rowIndex = 0;


        sheet.setColumnWidth(0, 9 * 256);
        sheet.setColumnWidth(1, 100 * 256);
        sheet.setColumnWidth(2, 150 * 256);
        sheet.setColumnWidth(3, 150 * 256);

        Row row1 = sheet.createRow(rowIndex);
        row1.setHeightInPoints(75);

        Cell cell = row1.createCell(0);
        cell.setCellStyle(Styles.getCellStyle(workbook));
        cell.setCellValue("АНДЕРРАЙТЕР ХУЛОСАСИ");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 3));

        rowIndex += 1;

//        InputStream inputStream = new FileInputStream("C:\\Users\\user\\IdeaProjects\\ExcelGenerator\\src\\main\\resources\\static\\img.png");  // Replace with your image path
//        byte[] imageBytes = inputStream.readAllBytes();
//        int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
//        inputStream.close();
//
//        // 🔹 Create a drawing object
//        Drawing<?> drawing = sheet.createDrawingPatriarch();
//
//        // 🔹 Define anchor (Position of image)
//        XSSFClientAnchor anchor = new XSSFClientAnchor();
//        anchor.setCol1(3);  // Column C (index 2)
//        anchor.setRow1(1);  // Start at row 1
//        anchor.setCol2(3);  // Span to next column (optional)
//        anchor.setRow2(5);  // End at row 5
//        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);// Image moves & resizes with cells
//
//        // 🔹 Insert the picture
//        Picture picture = drawing.createPicture(anchor, pictureIdx);
//        picture.resize();

        for (int rowNum = 2; rowNum <= 8; rowNum++) {
            Row row2 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw2cell1 = row2.createCell(0);
                    raw2cell1.setCellStyle(Styles.getCellStyle(workbook));
                    raw2cell1.setCellValue("1");
                    Cell raw2cell2 = row2.createCell(1);
                    raw2cell2.setCellValue("Сана:");
                    raw2cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw2cell3 = row2.createCell(2);
                    raw2cell3.setCellValue("03.02.2025 йил");
                    raw2cell3.setCellStyle(Styles.getCellStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row2.createCell(0);
                    raw3cell1.setCellValue("2");
                    raw3cell1.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell2 = row2.createCell(1);
                    raw3cell2.setCellValue("Хизмат кўрсатувчи БХО:");
                    raw3cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw3cell3 = row2.createCell(2);
                    raw3cell3.setCellValue("Қўқон БХМ");
                    raw3cell3.setCellStyle(Styles.getCellStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row2.createCell(0);
                    raw4cell1.setCellValue("3");
                    raw4cell1.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell2 = row2.createCell(1);
                    raw4cell2.setCellValue("Корхона номи:");
                    raw4cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw4cell3 = row2.createCell(2);
                    raw4cell3.setCellValue("\"SHODLIK TECHNO\" МЧЖ");
                    raw4cell3.setCellStyle(Styles.getCellStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row2.createCell(0);
                    raw5cell1.setCellValue("4");
                    raw5cell1.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw5cell2 = row2.createCell(1);
                    raw5cell2.setCellValue("Уникали:");
                    raw5cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw5cell3 = row2.createCell(2);
                    raw5cell3.setCellValue("01024509");
                    raw5cell3.setCellStyle(Styles.getCellStyle(workbook));
                    break;

                case 6:
                    Cell raw6cell1 = row2.createCell(0);
                    raw6cell1.setCellValue("5");
                    raw6cell1.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw6cell2 = row2.createCell(1);
                    raw6cell2.setCellValue("ID заявки");
                    raw6cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw6cell3 = row2.createCell(2);
                    raw6cell3.setCellValue("2845623");
                    raw6cell3.setCellStyle(Styles.getCellStyle(workbook));
                    break;

                case 7:
                    Cell raw7cell1 = row2.createCell(0);
                    raw7cell1.setCellValue("6");
                    raw7cell1.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw7cell2 = row2.createCell(1);
                    raw7cell2.setCellValue("ИНН:");
                    raw7cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw7cell3 = row2.createCell(2);
                    raw7cell3.setCellValue("306110816");
                    raw7cell3.setCellStyle(Styles.getCellStyle(workbook));
                    break;

                case 8:
                    Cell raw8cell1 = row2.createCell(0);
                    raw8cell1.setCellValue("7");
                    raw8cell1.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw8cell2 = row2.createCell(1);
                    raw8cell2.setCellValue("Андеррайтерга тўлиқ юборилган сана");
                    raw8cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw8cell3 = row2.createCell(2);
                    raw8cell3.setCellValue("03.02.2025 йил");
                    raw8cell3.setCellStyle(Styles.getCellStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row9 = sheet.createRow(rowIndex);
        row9.setHeightInPoints(35);

        Cell raw9cell = row9.createCell(0);
        Cell raw9cell2 = row9.createCell(1);
        Cell raw9cell3 = row9.createCell(2);
        Cell raw9cell4 = row9.createCell(3);
        raw9cell.setCellStyle(Styles.getBackgroundAndText(workbook));
        raw9cell2.setCellStyle(Styles.getBackgroundAndText(workbook));
        raw9cell3.setCellStyle(Styles.getBackgroundAndText(workbook));
        raw9cell4.setCellStyle(Styles.getBackgroundAndText(workbook));
        raw9cell.setCellValue("\"Универсал\" кредит  махсулоти паспортига мослиги");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 3));
        rowIndex += 1;

        Row row10 = sheet.createRow(rowIndex);
        row10.setHeightInPoints(35);

        Cell raw10cell = row10.createCell(0);
        Cell raw10cell2 = row10.createCell(1);
        Cell raw10cell3 = row10.createCell(2);
        Cell raw10cell4 = row10.createCell(3);
        raw10cell.setCellStyle(Styles.getBackground(workbook));
        raw10cell2.setCellStyle(Styles.getBackground(workbook));
        raw10cell3.setCellStyle(Styles.getBackground(workbook));
        raw10cell4.setCellStyle(Styles.getBackground(workbook));
        raw10cell.setCellValue("№");
        raw10cell2.setCellValue(" Керакли маълумотлар/хужжатлар");
        raw10cell3.setCellValue("Кредит маҳсулот паспорти бўйича талаб");
        raw10cell4.setCellValue("Хакикатда хужжатлар ва маълумотларнинг \"Анкета\" дастурида мослиги ");
        rowIndex += 1;

        Row row11 = sheet.createRow(rowIndex);
        row11.setHeightInPoints(35);

        Cell raw11cell = row11.createCell(0);
        Cell raw11cell2 = row11.createCell(1);
        Cell raw11cell3 = row11.createCell(2);
        Cell raw11cell4 = row11.createCell(3);
        raw11cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw11cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw11cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw11cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw11cell.setCellValue("1");
        raw11cell2.setCellValue("Ариза");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        for (int rowNum = 2; rowNum <= 10; rowNum++) {
            Row row12 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw2cell1 = row12.createCell(0);
                    raw2cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw2cell2 = row12.createCell(1);
                    raw2cell2.setCellValue("Мижоз аризаси санаси");
                    raw2cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw2cell3 = row12.createCell(2);
                    raw2cell3.setCellValue("-");
                    raw2cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw2cell4 = row12.createCell(3);
                    raw2cell4.setCellValue(excelRequest.getApplicationDate());
                    raw2cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row12.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row12.createCell(1);
                    raw3cell2.setCellValue("Кирим қилинган сана");
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row12.createCell(2);
                    raw3cell3.setCellValue("-");
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw3cell4 = row12.createCell(3);
                    raw3cell4.setCellValue(excelRequest.getEntryDate());
                    raw3cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row12.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row12.createCell(1);
                    raw4cell2.setCellValue("Мижоз ҳисоб рақами");
                    raw4cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell3 = row12.createCell(2);
                    raw4cell3.setCellValue("Асосий / иккиламчи");
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw4cell4 = row12.createCell(3);
                    raw4cell4.setCellValue(excelRequest.getClientAccount());
                    raw4cell4.setCellStyle(Styles.getCellStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row12.createCell(0);
                    raw5cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw5cell2 = row12.createCell(1);
                    raw5cell2.setCellValue("Кредит мақсади");
                    raw5cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw5cell3 = row12.createCell(2);
                    raw5cell3.setCellValue("Мақсадлилик тамойили мавжуд эмас");
                    raw5cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw5cell4 = row12.createCell(3);
                    raw5cell4.setCellValue(excelRequest.getLoanPurpose());
                    raw5cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 6:
                    Cell raw6cell1 = row12.createCell(0);
                    row12.setHeightInPoints(350);
                    raw6cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw6cell2 = row12.createCell(1);
                    raw6cell2.setCellValue("Кредит миқдори");
                    raw6cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw6cell3 = row12.createCell(2);
                    raw6cell3.setCellValue("\"“Бизнесни ривожлантириш банки” АТБ тизимида асосий ҳисоб \n" +
                            "рақами мавжуд мижозларга - 5 000 000 000 сўмгача;\n" +
                            "\uF0B7“Бизнесни ривожлантириш банки” АТБ тизимида иккиламчи ҳисоб \n" +
                            "рақами мавжуд мижозларга - 1 000 000 000 сўмгача;\n" +
                            "Бунда,тадбиркорлик субектларининг мулкчилик шаклидан келиб чиқиб қуйидагича тақсимланади:\n" +
                            "\uF0B7 600,0 млн.сўмгача - ЯТТларга;\n" +
                            "\uF0B7 1 000,0 млн.сўмгача - микрофирмалар (жами даромади \n" +
                            "охирги 12 ойда 1,0 млрд.сўмгача бўлган тадбиркорлик субектлари);\n" +
                            "\uF0B7 5 000,0 млн.сўмгача - кичик корхоналар (жами даромади \n" +
                            "охирги 12 ойда 1,0 млрд.сўмдан 10,0 млрд.сўмгача бўлган тадбиркорлик субектлари);\n" +
                            "\uF0B7 5 000,0 млн.сўмгача - ўрта тадбиркорлик субектлари (жами даромади \n" +
                            "охирги 12 ойда 10,0 млрд.сўмдан юқори бўлган тадбиркорлик субектлари).\"");

                    raw6cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw6cell4 = row12.createCell(3);
                    raw6cell4.setCellValue(excelRequest.getCreditAmount());
                    raw6cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 7:
                    Cell raw7cell1 = row12.createCell(0);
                    row12.setHeightInPoints(250);
                    raw7cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw7cell2 = row12.createCell(1);
                    raw7cell2.setCellValue("Кредит муддати");
                    raw7cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw7cell3 = row12.createCell(2);
                    raw7cell3.setCellValue("\" “Бизнесни ривожлантириш банки” АТБ тизимида асосий ҳисоб \n" +
                            "рақами мавжуд мижозларга - 36 ойгача;\n" +
                            "\uF0B7 “Бизнесни ривожлантириш банки” АТБ тизимида асосий ҳисоб \n" +
                            "рақами мавжуд бўлмаган мижозларга - 24 ойгача;\n" +
                            "\uF0B7 24 ой муддатгача Бош келишув имзоланган ҳолда \n" +
                            "индивидуал кредит шартномаларига асосан 12 ойгача очиқ кредит линияси орқали (револвер шаклда) кредит ажратилиши мумкин.\"");
                    raw7cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw7cell4 = row12.createCell(3);
                    raw7cell4.setCellValue(excelRequest.getCreditDuration());
                    raw7cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 8:
                    Cell raw8cell1 = row12.createCell(0);
                    raw8cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw8cell2 = row12.createCell(1);
                    raw8cell2.setCellValue("Имтиёзли давр");
                    raw8cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw8cell3 = row12.createCell(2);
                    raw8cell3.setCellValue("6 ойгача");
                    raw8cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw8cell4 = row12.createCell(3);
                    raw8cell4.setCellValue(excelRequest.getGracePeriod());
                    raw8cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 9:
                    Cell raw9cell1 = row12.createCell(0);
                    row12.setHeightInPoints(350);
                    raw9cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raww9cell2 = row12.createCell(1);
                    raww9cell2.setCellValue("Фоизи");
                    raww9cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raww9cell3 = row12.createCell(2);
                    raww9cell3.setCellValue("\" “Бизнесни ривожлантириш банки” АТБ тизимида асосий ҳисоб \n" +
                            "рақами мавжуд бўлганда;\n" +
                            "Кредит муддати 12 ойгача - 27%\n" +
                            "Кредит муддати 12 ойдан 24 ойгача - 28%\n" +
                            "Кредит муддати 24 ойдан 36 ойгача - 30%\n" +
                            "24 ой муддатгача Бош келишув имзолаган ҳолда индивидуал кредит шартномаларига асосан 12 ойгача очиқ кредит линияси орқали (револьвер шаклда) кредит ажратилганда - 27%.\n" +
                            "\n" +
                            "“Бизнесни ривожлантириш банки” АТБ тизимида асосий ҳисоб \n" +
                            "рақами мавжуд бўлмаганда;\n" +
                            "Кредит муддати 12 ойгача - 28%\n" +
                            "Кредит муддати 12 ойдан 24 ойгача - 29%\"");
                    raww9cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raww9cell4 = row12.createCell(3);
                    raww9cell4.setCellValue(excelRequest.getInterestRate());
                    raww9cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 10:
                    Cell raw10cell1 = row12.createCell(0);
                    raw10cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw12cell2 = row12.createCell(1);
                    raw12cell2.setCellValue("Кредитлаш усули");
                    raw12cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw12cell3 = row12.createCell(2);
                    raw12cell3.setCellValue("Очиқ ва опиқ кредит линияси орқали");
                    raw12cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raww10cell4 = row12.createCell(3);
                    raww10cell4.setCellValue(excelRequest.getLendingMethod());
                    raww10cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row21 = sheet.createRow(rowIndex);

        Cell raw21cell = row21.createCell(0);
        Cell raw21cell2 = row21.createCell(1);
        Cell raw21cell3 = row21.createCell(2);
        Cell raw21cell4 = row21.createCell(3);
        raw21cell.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
        raw21cell2.setCellStyle(Styles.getCellStyle(workbook));
        raw21cell3.setCellStyle(Styles.getItalicStyle(workbook));
        raw21cell4.setCellStyle(Styles.getCellStyle(workbook));
        raw21cell2.setCellValue("Кредит валютаси");
        raw21cell3.setCellValue("Миллий валюта (сўм)");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 2, 3));
        rowIndex += 1;

        for (int rowNum = 2; rowNum <= 4; rowNum++) {
            Row row22 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw22cell1 = row22.createCell(0);
                    raw22cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw22cell2 = row22.createCell(1);
                    raw22cell2.setCellValue("Молиялаштириш манбаи");
                    raw22cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw22cell3 = row22.createCell(2);
                    raw22cell3.setCellValue("Банк ўз маблағи ва (ёки) жалб қилинган маблағлари ҳисобидан");
                    raw22cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw22cell4 = row22.createCell(3);
                    raw22cell4.setCellValue(excelRequest.getFundingSource());
                    raw22cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row22.createCell(0);
                    row22.setHeightInPoints(50);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row22.createCell(1);
                    raw3cell2.setCellValue("Кредитни ажратиш шакли");
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row22.createCell(2);
                    raw3cell3.setCellValue("Мижознинг \"Бизнесни ривожлантириш банки\" АТБ тизимида очилган асосий ёки \nиккиламчи хисоб рақамига пул ўтказиб берилади.");
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw23cell4 = row22.createCell(3);
                    raw23cell4.setCellValue(excelRequest.getLoanDisbursementMethod());
                    raw23cell4.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row22.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row22.createCell(1);
                    raw4cell2.setCellValue("Қўшимча шарт");
                    raw4cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell3 = row22.createCell(2);
                    raw4cell3.setCellValue("Агар мавжуд бўлса");
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw24cell4 = row22.createCell(3);
                    raw24cell4.setCellValue(excelRequest.getAdditionalCondition());
                    raw24cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row25 = sheet.createRow(rowIndex);

        Cell raw25cell = row25.createCell(0);
        Cell raw25cell2 = row25.createCell(1);
        Cell raw25cell3 = row25.createCell(2);
        Cell raw25cell4 = row25.createCell(3);
        raw25cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw25cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw25cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw25cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw25cell.setCellValue("2");
        raw25cell2.setCellValue("Қарз олувчининг таъсис ҳужжатлари (устав, гувоҳнома, паспорт нусхалари, имзо наъмунаси)");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        for (int rowNum = 2; rowNum <= 7; rowNum++) {
            Row row26 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw2cell1 = row26.createCell(0);
                    raw2cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw2cell2 = row26.createCell(1);
                    raw2cell2.setCellValue("Юридик манзили (Низом)");
                    raw2cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw2cell3 = row26.createCell(2);
                    raw2cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw2cell4 = row26.createCell(3);
                    raw2cell4.setCellValue(excelRequest.getLegalAddress());
                    raw2cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row26.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row26.createCell(1);
                    raw3cell2.setCellValue("Корхона ташкил топган сана ");
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row26.createCell(2);
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw3cell4 = row26.createCell(3);
                    raw3cell4.setCellValue(excelRequest.getEstablishmentDate());
                    raw3cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row26.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row26.createCell(1);
                    raw4cell2.setCellValue("Таъсисчилар ва уларнинг улуши (stat.uz )");
                    raw4cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell3 = row26.createCell(2);
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw4cell4 = row26.createCell(3);
                    raw4cell4.setCellValue(excelRequest.getFoundersAndShares());
                    raw4cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row26.createCell(0);
                    raw5cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw5cell2 = row26.createCell(1);
                    raw5cell2.setCellValue("Кредит олиш қарори");
                    raw5cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw5cell3 = row26.createCell(2);
                    raw5cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw5cell4 = row26.createCell(3);
                    raw5cell4.setCellValue(excelRequest.getLoanApprovalDecision());
                    raw5cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 6:
                    Cell raw6cell1 = row26.createCell(0);
                    raw6cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw6cell2 = row26.createCell(1);
                    raw6cell2.setCellValue("Низом жамғармаси суммаси");
                    raw6cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw6cell3 = row26.createCell(2);
                    raw6cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw6cell4 = row26.createCell(3);
                    raw6cell4.setCellValue(excelRequest.getCharterCapitalAmount());
                    raw6cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 7:
                    Cell raw7cell1 = row26.createCell(0);
                    raw7cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw7cell2 = row26.createCell(1);
                    raw7cell2.setCellValue("Асосий фаолияти (Низом бўйича)");
                    raw7cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw7cell3 = row26.createCell(2);
                    raw7cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw7cell4 = row26.createCell(3);
                    raw7cell4.setCellValue(excelRequest.getMainActivity());
                    raw7cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row32 = sheet.createRow(rowIndex);

        Cell raw32cell = row32.createCell(0);
        Cell raw32cell2 = row32.createCell(1);
        Cell raw32cell3 = row32.createCell(2);
        Cell raw32cell4 = row32.createCell(3);
        raw32cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw32cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw32cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw32cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw32cell.setCellValue("3");
        raw32cell2.setCellValue("Кредит ахборот тахлилий маркази (КАТМ) маълумотлари");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + 7, 1, 1));

        for (int rowNum = 2; rowNum <= 9; rowNum++) {
            Row row33 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw2cell1 = row33.createCell(0);
                    raw2cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw2cell2 = row33.createCell(1);
                    raw2cell2.setCellValue("Ижобий кредит тарихига эга бўлиши");
                    raw2cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw2cell3 = row33.createCell(2);
                    raw2cell3.setCellValue("Мавжуд амалдаги кредитлари сони");
                    raw2cell3.setCellStyle(Styles.getItalicStyleWithRed(workbook));
                    Cell raw2cell4 = row33.createCell(3);
                    raw2cell4.setCellValue(excelRequest.getActiveLoanCount());
                    raw2cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row33.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row33.createCell(1);
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row33.createCell(2);
                    raw3cell3.setCellValue("Мавжуд амалдаги кредитлари қолдиғи");
                    raw3cell3.setCellStyle(Styles.getItalicStyleWithRed(workbook));
                    Cell raw3cell4 = row33.createCell(3);
                    raw3cell4.setCellValue(excelRequest.getActiveLoanBalance());
                    raw3cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row33.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row33.createCell(1);
                    raw4cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell3 = row33.createCell(2);
                    raw4cell3.setCellValue("Мавжуд кредитлари бўйича муддати ўтган асосий/ фоиз қарздорлиги тўғрисида");
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw4cell4 = row33.createCell(3);
                    raw4cell4.setCellValue(excelRequest.getOverduePrincipalAndInterest());
                    raw4cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row33.createCell(0);
                    raw5cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw5cell2 = row33.createCell(1);
                    raw5cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw5cell3 = row33.createCell(2);
                    row33.setHeightInPoints(50);
                    raw5cell3.setCellValue("\"Мижознинг суд жараёнидаги кредит қолдиқлари ва балансдан ташқари \n" +
                            "ҳисобвараққа ўтказилган кредит асосий қарзи ва фоизлари мавжуд бўлмаслиги\"");
                    raw5cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw5cell4 = row33.createCell(3);
                    raw5cell4.setCellValue(excelRequest.getNoLegalProceedingsOrOffBalanceLoans());
                    raw5cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 6:
                    Cell raw6cell1 = row33.createCell(0);
                    raw6cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw6cell2 = row33.createCell(1);
                    raw6cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw6cell3 = row33.createCell(2);
                    raw6cell3.setCellValue("KATM бали 200 баллдан юқори бўлиши");
                    raw6cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw6cell4 = row33.createCell(3);
                    raw6cell4.setCellValue(excelRequest.getKatmScoreAbove200());
                    raw6cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 7:
                    Cell raw7cell1 = row33.createCell(0);
                    raw7cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw7cell2 = row33.createCell(1);
                    raw7cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw7cell3 = row33.createCell(2);
                    row33.setHeightInPoints(50);
                    raw7cell3.setCellValue("КАТМ мижознинг барча тижорат банклари тизимида “қониқарсиз”, “шубҳали” ва \n“умидсиз” тоифаларида таснифланган амалдаги кредитлари мавжуд бўлмаслиги;");
                    raw7cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw7cell4 = row33.createCell(3);
                    raw7cell4.setCellValue(excelRequest.getNoUnsatisfactoryLoansInAllBanks());
                    raw7cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 8:
                    Cell raw8cell1 = row33.createCell(0);
                    raw8cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw8cell2 = row33.createCell(1);
                    raw8cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw8cell3 = row33.createCell(2);
                    row33.setHeightInPoints(80);
                    raw8cell3.setCellValue("\"Ўзаро алоқадор тадбиркорлик субъектларининг номи\n" +
                            "(orginfo.uz сайти орқали ўрганилди)\"\n");
                    raw8cell3.setCellStyle(Styles.getItalicStyleWithRed(workbook));
                    Cell raw8cell4 = row33.createCell(3);
                    raw8cell4.setCellValue(excelRequest.getRelatedBusinessEntities());
                    raw8cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 9:
                    Cell raw1cell1 = row33.createCell(0);
                    raw1cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw1cell2 = row33.createCell(1);
                    raw1cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw1cell3 = row33.createCell(2);
                    raw1cell3.setCellValue("Ўзаро алоқадор тадбиркорлик субъектларининг “БРБ ” АТБ муддати ўтган \nқарздорликлари мавжуд бўлмаслиги");
                    raw1cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw1cell4 = row33.createCell(3);
                    raw1cell4.setCellValue(excelRequest.getNoOverdueDebtsInBRB());
                    raw1cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row41 = sheet.createRow(rowIndex);

        Cell raw41cell = row41.createCell(0);
        Cell raw41cell2 = row41.createCell(1);
        Cell raw41cell3 = row41.createCell(2);
        Cell raw41cell4 = row41.createCell(3);
        raw41cell.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
        raw41cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw41cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw41cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw41cell2.setCellValue("Кредитлар бўйича охирги 24 ой ичида:");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        for (int rowNum = 2; rowNum <= 5; rowNum++) {
            Row row42 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw22cell1 = row42.createCell(0);
                    raw22cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw22cell2 = row42.createCell(1);
                    raw22cell2.setCellValue("30 кундан ошган муддати ўтган қарздорликлари");
                    raw22cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw22cell3 = row42.createCell(2);
                    raw22cell3.setCellValue("10 тадан кўп бўлмаслиги");
                    raw22cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw22cell4 = row42.createCell(3);
                    raw22cell4.setCellValue(excelRequest.getOverdueMoreThan30Days());
                    raw22cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row42.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row42.createCell(1);
                    raw3cell2.setCellValue("60 кундан ошган муддати ўтган қарздорликлари ");
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row42.createCell(2);
                    raw3cell3.setCellValue("2 тадан кўп бўлмаслиги");
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw3cell4 = row42.createCell(3);
                    raw3cell4.setCellValue(excelRequest.getOverdueMoreThan60Days());
                    raw3cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row42.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row42.createCell(1);
                    raw4cell2.setCellValue("90 кундан ошган муддати ўтган қарздорликлари");
                    raw4cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell3 = row42.createCell(2);
                    raw4cell3.setCellValue("1 тадан кўп бўлмаслиги;");
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw4cell4 = row42.createCell(3);
                    raw4cell4.setCellValue(excelRequest.getOverdueMoreThan90Days());
                    raw4cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row42.createCell(0);
                    raw5cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw5cell2 = row42.createCell(1);
                    raw5cell2.setCellValue("Охирги 12 ой ичида 90 кундан ошган муддати \nўтган қарздорликлари");
                    raw5cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw5cell3 = row42.createCell(2);
                    raw5cell3.setCellValue("умуман бўлмаслиги;");
                    raw5cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw5cell4 = row42.createCell(3);
                    raw5cell4.setCellValue(excelRequest.getOverdueMoreThan90DaysLast12Months());
                    raw5cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row46 = sheet.createRow(rowIndex);

        Cell raw46cell = row46.createCell(0);
        Cell raw46cell2 = row46.createCell(1);
        Cell raw46cell3 = row46.createCell(2);
        Cell raw46cell4 = row46.createCell(3);
        raw46cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw46cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw46cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw46cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw46cell.setCellValue("4");
        raw46cell2.setCellValue("Мавжуд кредитлари тўғрисида маълумот");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        for (int rowNum = 2; rowNum <= 8; rowNum++) {
            Row row47 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw2cell1 = row47.createCell(0);
                    raw2cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw2cell2 = row47.createCell(1);
                    raw2cell2.setCellValue("Кредитнинг шартнома суммаси");
                    raw2cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw2cell3 = row47.createCell(2);
                    raw2cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw2cell4 = row47.createCell(3);
                    raw2cell4.setCellValue(excelRequest.getContractAmount());
                    raw2cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row47.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row47.createCell(1);
                    raw3cell2.setCellValue("Кредит қолдиғи");
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row47.createCell(2);
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw3cell4 = row47.createCell(3);
                    raw3cell4.setCellValue(excelRequest.getRemainingCredit());
                    raw3cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row47.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row47.createCell(1);
                    raw4cell2.setCellValue("Мақсади");
                    raw4cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell3 = row47.createCell(2);
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw4cell4 = row47.createCell(3);
                    raw4cell4.setCellValue(excelRequest.getPurpose());
                    raw4cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row47.createCell(0);
                    raw5cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw5cell2 = row47.createCell(1);
                    raw5cell2.setCellValue("Муддати");
                    raw5cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw5cell3 = row47.createCell(2);
                    raw5cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw5cell4 = row47.createCell(3);
                    raw5cell4.setCellValue(excelRequest.getDuration());
                    raw5cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 6:
                    Cell raw6cell1 = row47.createCell(0);
                    raw6cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw6cell2 = row47.createCell(1);
                    raw6cell2.setCellValue("Муддати ўтган график суммаси");
                    raw6cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw6cell3 = row47.createCell(2);
                    raw6cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw6cell4 = row47.createCell(3);
                    raw6cell4.setCellValue(excelRequest.getOverdueScheduledAmount());
                    raw6cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 7:
                    Cell raw7cell1 = row47.createCell(0);
                    raw7cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw7cell2 = row47.createCell(1);
                    raw7cell2.setCellValue("Муддати ўтган фоиз суммаси");
                    raw7cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw7cell3 = row47.createCell(2);
                    raw7cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw7cell4 = row47.createCell(3);
                    raw7cell4.setCellValue(excelRequest.getOverdueInterestAmount());
                    raw7cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 8:
                    Cell raw8cell1 = row47.createCell(0);
                    raw8cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw8cell2 = row47.createCell(1);
                    raw8cell2.setCellValue("Мавжуд таъминотлари");
                    raw8cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw8cell3 = row47.createCell(2);
                    row47.setHeightInPoints(80);
                    raw8cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw8cell4 = row47.createCell(3);
                    raw8cell4.setCellValue(excelRequest.getAvailableCollateral());
                    raw8cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row54 = sheet.createRow(rowIndex);

        Cell raw54cell = row54.createCell(0);
        Cell raw54cell2 = row54.createCell(1);
        Cell raw54cell3 = row54.createCell(2);
        Cell raw54cell4 = row54.createCell(3);
        raw54cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw54cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw54cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw54cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw54cell.setCellValue("5");
        raw54cell2.setCellValue("Молиявий натижалари тўғрисида ва хисоб рақам айланмалари тўғрисида");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + 2, 1, 1));

        for (int rowNum = 2; rowNum <= 5; rowNum++) {
            Row row55 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw22cell1 = row55.createCell(0);
                    row55.setHeightInPoints(100);
                    raw22cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw22cell2 = row55.createCell(1);
                    raw22cell2.setCellValue("\"Асосий ва иккиламчи хисоб рақам орқали \nайланмалар тўғрисида  \n" +
                            "(жами ҳисоб рақамларида, жумладан бошқа \nбанкдаги)");
                    raw22cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw22cell3 = row55.createCell(2);
                    raw22cell3.setCellValue("21.08.2023 - 31.12.2023");
                    raw22cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw22cell4 = row55.createCell(3);
                    raw22cell4.setCellValue(excelRequest.getPeriodOne());
                    raw22cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row55.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row55.createCell(1);
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row55.createCell(2);
                    raw3cell3.setCellValue("\"Охирги 12 ойда пул айланмаларига эга бўлиши \n" +
                            "01.01.2024 - 31.12.2024");
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw3cell4 = row55.createCell(3);
                    raw3cell4.setCellValue(excelRequest.getLast12MonthsTurnover());
                    raw3cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row55.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row55.createCell(1);
                    raw4cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell3 = row55.createCell(2);
                    raw4cell3.setCellValue("01.01.2025 - 29.01.2025");
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw4cell4 = row55.createCell(3);
                    raw4cell4.setCellValue(excelRequest.getPeriodTwo());
                    raw4cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row55.createCell(0);
                    raw5cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw5cell2 = row55.createCell(1);
                    raw5cell2.setCellValue("Айланмалари (Ф-2 010 сатр)");
                    raw5cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw5cell3 = row55.createCell(2);
                    raw5cell3.setCellValue("2024-йиллик баланс бўйича");
                    raw5cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw5cell4 = row55.createCell(3);
                    raw5cell4.setCellValue(excelRequest.getAnnualBalance2024());
                    raw5cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row59 = sheet.createRow(rowIndex);

        Cell raw59cell = row59.createCell(0);
        Cell raw59cell2 = row59.createCell(1);
        Cell raw59cell3 = row59.createCell(2);
        Cell raw59cell4 = row59.createCell(3);
        raw59cell.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
        raw59cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw59cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw59cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw59cell2.setCellValue("1,0 млрд сўмгача бўлган лойиҳалар учун талаблар");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        for (int rowNum = 2; rowNum <= 5; rowNum++) {
            Row row60 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw22cell1 = row60.createCell(0);
                    raw22cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw22cell2 = row60.createCell(1);
                    raw22cell2.setCellValue("Фойда ёки зарар (Ф-2 270 сатр)");
                    raw22cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw22cell3 = row60.createCell(2);
                    raw22cell3.setCellValue("Молиявий натижалар тўғрисида ҳисобот (2-сон шакл) охирги ҳисобот даври билан зарар \nбилан  якунланмаган бўлиши");
                    raw22cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw2cell4 = row60.createCell(3);
                    raw2cell4.setCellValue(excelRequest.getProfitOrLoss());
                    raw2cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row60.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row60.createCell(1);
                    raw3cell2.setCellValue("Ўз айланма маблағлари суммаси (тахлил)");
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row60.createCell(2);
                    raw3cell3.setCellValue("Ўз айланма маблағлари мавжудлиги манфий кўрсаткичда бўлмаслиги;");
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw3cell4 = row60.createCell(3);
                    raw3cell4.setCellValue(excelRequest.getOwnWorkingCapital());
                    raw3cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row60.createCell(0);
                    row60.setHeightInPoints(60);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row60.createCell(1);
                    raw4cell2.setCellValue("mib.uz");
                    raw4cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell3 = row60.createCell(2);
                    raw4cell3.setCellValue("Мижознинг Мажбурий ижро бюроси томонидан очилган ижро иши бўйича тўланиши \nлозим бўлган маблағ бўлмаслиги;");
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw4cell4 = row60.createCell(3);
                    raw4cell4.setCellValue(excelRequest.getMibUz());
                    raw4cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row60.createCell(0);
                    raw5cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw5cell2 = row60.createCell(1);
                    raw5cell2.setCellValue("2-сонли картотека  қарздорлиги");
                    raw5cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw5cell3 = row60.createCell(2);
                    raw5cell3.setCellValue("2-сонли картотека ҳисобварағида қарздорлик мавжуд бўлмаслиги;");
                    raw5cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw5cell4 = row60.createCell(3);
                    raw5cell4.setCellValue(excelRequest.getSecondRegistryDebt());
                    raw5cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row64 = sheet.createRow(rowIndex);

        Cell raw64cell = row64.createCell(0);
        Cell raw64cell2 = row64.createCell(1);
        Cell raw64cell3 = row64.createCell(2);
        Cell raw64cell4 = row64.createCell(3);
        raw64cell.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
        raw64cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw64cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw64cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw64cell2.setCellValue("1,0 млрд сўмдан юқори бўлган лойиҳалар учун қўшимча талаблар");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        for (int rowNum = 2; rowNum <= 4; rowNum++) {
            Row row65 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw22cell1 = row65.createCell(0);
                    raw22cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw22cell2 = row65.createCell(1);
                    raw22cell2.setCellValue("Мижоз фаолиятидан тушумга");
                    raw22cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw22cell3 = row65.createCell(2);
                    raw22cell3.setCellValue("Мижоз фаолиятидан тушумга эга бўлиши");
                    raw22cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw2cell4 = row65.createCell(3);
                    raw2cell4.setCellValue(excelRequest.getClientRevenue());
                    raw2cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row65.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row65.createCell(1);
                    row65.setHeightInPoints(90);
                    raw3cell2.setCellValue("\"Бизнесни ривожлантириш банки\" АТБ тизимида \nҳисоб рақами мавжудлиги");
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row65.createCell(2);
                    raw3cell3.setCellValue("\"Бизнесни ривожлантириш банки\" АТБ тизимида фақат асосий ҳисоб рақами мавжуд \nбўлиши");
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw3cell4 = row65.createCell(3);
                    raw3cell4.setCellValue(excelRequest.getHasAccountInBRBBank());
                    raw3cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row65.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row65.createCell(1);
                    row65.setHeightInPoints(60);
                    raw4cell2.setCellValue("Қарз олувчи кредит юкламаси");
                    raw4cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell3 = row65.createCell(2);
                    raw4cell3.setCellValue("Қарз олувчи кредит юкламаси 100% дан кўп бўлмаслиги");
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw4cell4 = row65.createCell(3);
                    raw4cell4.setCellValue(excelRequest.getBorrowerCreditLoad());
                    raw4cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row68 = sheet.createRow(rowIndex);

        Cell raw68cell = row68.createCell(0);
        Cell raw68cell2 = row68.createCell(1);
        Cell raw68cell3 = row68.createCell(2);
        Cell raw68cell4 = row68.createCell(3);
        raw68cell.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
        raw68cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw68cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw68cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw68cell2.setCellValue("500,0 млн сўмдан юқори бўлган лойиҳалар учун қўшимча талаблар (гаровсиз)");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + 1, 2, 2));

        for (int rowNum = 2; rowNum <= 3; rowNum++) {
            Row row69 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw22cell1 = row69.createCell(0);
                    raw22cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    row69.setHeightInPoints(120);
                    Cell raw22cell2 = row69.createCell(1);
                    raw22cell2.setCellValue("Қарз юки кўрсаткичи");
                    raw22cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw22cell3 = row69.createCell(2);
                    raw22cell3.setCellValue("\"Ваколат доирасида Кредит қўмитаси қарорига асосан доимий пул оқимига ва ижобий \nкредит тарихига эга мижозларга кредит суммасининг 125% миқдорида суғурта полиси \nёки учинчи шахс кафиллиги билан кредит ажратилишига рухсат этилиши мумкин \n(кредит миқдори 500.0 млн. сўмдан юқори бўлган лойиҳалар учун).\n" +
                            " Бунда, доимий пул оқимига эга мижозлар дейилганда жорий ва янги кредитлари билан \nбирга ҳисобланганда қарз юки 50% дан (50% ҳам киради) баланд бўлмаслиги ва сўнги 12 \nойда ҳисоб рақамида узликсиз тушуми мавжуд бўлиши (асосий фаолияти мавсумий \nбўлганда, тушуми узликсизлигига ўрнатилган талаб бундан мустасно) лозим.\"\n" +
                            "\n");
                    raw22cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw2cell4 = row69.createCell(3);
                    raw2cell4.setCellValue(excelRequest.getDebtLoadIndicator());
                    raw2cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row69.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row69.createCell(1);
                    row69.setHeightInPoints(120);
                    raw3cell2.setCellValue("Сўнги 12 ойда ҳисоб рақамида узликсиз тушуми \nмавжуд бўлиши");
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row69.createCell(2);
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    Cell raw3cell4 = row69.createCell(3);
                    raw3cell4.setCellValue(excelRequest.getUninterruptedAccountReceiptsLast12Months());
                    raw3cell4.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row71 = sheet.createRow(rowIndex);

        Cell raw71cell = row71.createCell(0);
        Cell raw71cell2 = row71.createCell(1);
        Cell raw71cell3 = row71.createCell(2);
        Cell raw71cell4 = row71.createCell(3);
        raw71cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw71cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw71cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw71cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw71cell.setCellValue("6");
        raw71cell2.setCellValue("Тақдим этилаётган таъминот маълумотлари");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + 10, 2, 2));

        for (int rowNum = 2; rowNum <= 12; rowNum++) {
            Row row72 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw2cell1 = row72.createCell(0);
                    raw2cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw2cell2 = row72.createCell(1);
                    raw2cell2.setCellValue("Манзили");
                    row72.setHeightInPoints(65);
                    raw2cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw2cell3 = row72.createCell(2);
                    raw2cell3.setCellValue("\"500 млн.сўмгача  125% номулкий - Қарз юки 20% гача;\n" +
                            "- 1 000 млн.сўмгача 75 % мулкий, 50% номулкий - Қарз юки 30% гача;\n" +
                            "- 2 000 млн.сўмгача 80 % мулкий, 45% номулкий - Қарз юки 30% гача;\n" +
                            "- Бошқа холларда 100% мулкий, 25% номулкий таъминотлар тақдим \nэтилиши лозим;\n" +
                            "Қуйидаги ҳолларда фақат 125% мулкий таъминот талаб этилади:\n" +
                            "- 1,0 млрд.сўмгача кредит ажратишда (Мижознинг қарз юки ёки \nбиргаликдаги қарз юки 100% дан юқори бўлганда;\n" +
                            "- 90 кундан ошган муддати ўтган қарздорликка 1 марта йўл қойган \nхолларда;\n" +
                            "- “БРБ” АТБ тизимида асосий ҳисобрақами мавжуд бўлмаган мижозларга \nкредит ажратишда;\n" +
                            "- 30 кундан ошган муддати ўтган қарздорликка 10 мартадан кўп бўлган \nхолларда;\n" +
                            "- Ваколат доирасида Кредит қўмитаси қарорига асосан доимий пул оқимига \nва ижобий кредит тарихига эга мижозларга кредит суммасининг 125% \nмиқдорида суғурта полиси ёки учинчи шахс кафиллиги билан кредит \nажратилишига рухсат этилиши мумкин (кредит миқдори 500.0 млн. \nсўмдан юқори бўлган лойиҳалар учун).\n" +
                            " Бунда, доимий пул оқимига эга мижозлар дейилганда жорий ва янги \nкредитлари билан бирга ҳисобланганда қарз юки 50% дан (50% ҳам киради) \nбаланд бўлмаслиги ва сўнги 12 ойда ҳисоб рақамида узликсиз тушуми \nмавжуд бўлиши (асосий фаолияти мавсумий бўлганда, тушуми \nузликсизлигига ўрнатилган талаб бундан мустасно) лозим.\n" +
                            "- Охирги 12 ой ичида мижознинг 50% дан юқори улушга эга таъсисчиси \nўзгарганда (яқин қариндошлар ўртасида ўзгариш бундан мустасно), камида \n125% фоиз ликвидли мол-мулк гарови тақдим этилиш лозим.\"\n");
                    raw2cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row72.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row72.createCell(1);
                    raw3cell2.setCellValue("Номи");
                    row72.setHeightInPoints(65);
                    raw3cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw3cell3 = row72.createCell(2);
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row72.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row72.createCell(1);
                    row72.setHeightInPoints(65);
                    raw4cell2.setCellValue("Гаров мулки эгаси");
                    raw4cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw4cell3 = row72.createCell(2);
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row72.createCell(0);
                    raw5cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw5cell2 = row72.createCell(1);
                    row72.setHeightInPoints(65);
                    raw5cell2.setCellValue("Тегишлилиги тўғрисида хужжат");
                    raw5cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw5cell3 = row72.createCell(2);
                    raw5cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 6:
                    Cell raw6cell1 = row72.createCell(0);
                    raw6cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw6cell2 = row72.createCell(1);
                    row72.setHeightInPoints(65);
                    raw6cell2.setCellValue("Рўйхатдан ўтганлиги юзасидан кадастр кўчирмаси");
                    raw6cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw6cell3 = row72.createCell(2);
                    raw6cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 7:
                    Cell raw7cell1 = row72.createCell(0);
                    raw7cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw7cell2 = row72.createCell(1);
                    row72.setHeightInPoints(65);
                    raw7cell2.setCellValue("Рўйҳатда ҳеч ким турмаслиги");
                    raw7cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw7cell3 = row72.createCell(2);
                    raw7cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 8:
                    Cell raw8cell1 = row72.createCell(0);
                    raw8cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw8cell2 = row72.createCell(1);
                    raw8cell2.setCellValue("Таъқиқ мавжудлиги бўйича маълумотнома");
                    raw8cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw8cell3 = row72.createCell(2);
                    row72.setHeightInPoints(65);
                    raw8cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 9:
                    Cell raw9cell1 = row72.createCell(0);
                    raw9cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raww9cell2 = row72.createCell(1);
                    raww9cell2.setCellValue("Мустақил баҳоловчи ташкилот нархи");
                    raww9cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raww9cell3 = row72.createCell(2);
                    row72.setHeightInPoints(65);
                    raww9cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 10:
                    Cell raw10cell1 = row72.createCell(0);
                    raw10cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raww10cell2 = row72.createCell(1);
                    raww10cell2.setCellValue("Е-баҳолаш нархи (Эксперт-2)");
                    raww10cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raww10cell3 = row72.createCell(2);
                    row72.setHeightInPoints(65);
                    raww10cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 11:
                    Cell raw11cell1 = row72.createCell(0);
                    raw11cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raww11cell2 = row72.createCell(1);
                    raww11cell2.setCellValue("Банк баҳолаш далолатномаси нархи");
                    raww11cell2.setCellStyle(Styles.getCellBasicStyleWithBackgroundGreen(workbook));
                    Cell raww11cell3 = row72.createCell(2);
                    row72.setHeightInPoints(65);
                    raww11cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 12:
                    Cell raw12cell1 = row72.createCell(0);
                    raw12cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw12cell2 = row72.createCell(1);
                    raw12cell2.setCellValue("Гаров эгасининг розилиги (паспорт)");
                    raw12cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw12cell3 = row72.createCell(2);
                    row72.setHeightInPoints(65);
                    raw12cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row81 = sheet.createRow(rowIndex);

        Cell raw81cell = row81.createCell(0);
        Cell raw81cell2 = row81.createCell(1);
        Cell raw81cell3 = row81.createCell(2);
        Cell raw81cell4 = row81.createCell(3);
        raw81cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw81cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw81cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw81cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw81cell.setCellValue("6");
        raw81cell2.setCellValue("Тақдим этилаётган таъминот маълумотлари");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + 10, 2, 2));

        for (int rowNum = 2; rowNum <= 12; rowNum++) {
            Row row82 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw2cell1 = row82.createCell(0);
                    raw2cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw2cell2 = row82.createCell(1);
                    raw2cell2.setCellValue("Манзили");
                    row82.setHeightInPoints(65);
                    raw2cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw2cell3 = row82.createCell(2);
                    raw2cell3.setCellValue("\"500 млн.сўмгача  125% номулкий - Қарз юки 20% гача;\n" +
                            "- 1 000 млн.сўмгача 75 % мулкий, 50% номулкий - Қарз юки 30% гача;\n" +
                            "- 2 000 млн.сўмгача 80 % мулкий, 45% номулкий - Қарз юки 30% гача;\n" +
                            "- Бошқа холларда 100% мулкий, 25% номулкий таъминотлар тақдим \nэтилиши лозим;\n" +
                            "Қуйидаги ҳолларда фақат 125% мулкий таъминот талаб этилади:\n" +
                            "- 1,0 млрд.сўмгача кредит ажратишда (Мижознинг қарз юки ёки \nбиргаликдаги қарз юки 100% дан юқори бўлганда;\n" +
                            "- 90 кундан ошган муддати ўтган қарздорликка 1 марта йўл қойган \nхолларда;\n" +
                            "- “БРБ” АТБ тизимида асосий ҳисобрақами мавжуд бўлмаган мижозларга \nкредит ажратишда;\n" +
                            "- 30 кундан ошган муддати ўтган қарздорликка 10 мартадан кўп бўлган \nхолларда;\n" +
                            "- Ваколат доирасида Кредит қўмитаси қарорига асосан доимий пул оқимига \nва ижобий кредит тарихига эга мижозларга кредит суммасининг 125% \nмиқдорида суғурта полиси ёки учинчи шахс кафиллиги билан кредит \nажратилишига рухсат этилиши мумкин (кредит миқдори 500.0 млн. \nсўмдан юқори бўлган лойиҳалар учун).\n" +
                            " Бунда, доимий пул оқимига эга мижозлар дейилганда жорий ва янги \nкредитлари билан бирга ҳисобланганда қарз юки 50% дан (50% ҳам киради) \nбаланд бўлмаслиги ва сўнги 12 ойда ҳисоб рақамида узликсиз тушуми \nмавжуд бўлиши (асосий фаолияти мавсумий бўлганда, тушуми \nузликсизлигига ўрнатилган талаб бундан мустасно) лозим.\n" +
                            "- Охирги 12 ой ичида мижознинг 50% дан юқори улушга эга таъсисчиси \nўзгарганда (яқин қариндошлар ўртасида ўзгариш бундан мустасно), камида \n125% фоиз ликвидли мол-мулк гарови тақдим этилиш лозим.\"\n");
                    raw2cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row82.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row82.createCell(1);
                    raw3cell2.setCellValue("Номи");
                    row82.setHeightInPoints(65);
                    raw3cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw3cell3 = row82.createCell(2);
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row82.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row82.createCell(1);
                    row82.setHeightInPoints(65);
                    raw4cell2.setCellValue("Гаров мулки эгаси");
                    raw4cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw4cell3 = row82.createCell(2);
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row82.createCell(0);
                    raw5cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw5cell2 = row82.createCell(1);
                    row82.setHeightInPoints(65);
                    raw5cell2.setCellValue("Тегишлилиги тўғрисида хужжат");
                    raw5cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw5cell3 = row82.createCell(2);
                    raw5cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 6:
                    Cell raw6cell1 = row82.createCell(0);
                    raw6cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw6cell2 = row82.createCell(1);
                    row82.setHeightInPoints(65);
                    raw6cell2.setCellValue("Рўйхатдан ўтганлиги юзасидан кадастр кўчирмаси");
                    raw6cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw6cell3 = row82.createCell(2);
                    raw6cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 7:
                    Cell raw7cell1 = row82.createCell(0);
                    raw7cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw7cell2 = row82.createCell(1);
                    row82.setHeightInPoints(65);
                    raw7cell2.setCellValue("Рўйҳатда ҳеч ким турмаслиги");
                    raw7cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw7cell3 = row82.createCell(2);
                    raw7cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 8:
                    Cell raw8cell1 = row82.createCell(0);
                    raw8cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw8cell2 = row82.createCell(1);
                    raw8cell2.setCellValue("Таъқиқ мавжудлиги бўйича маълумотнома");
                    raw8cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw8cell3 = row82.createCell(2);
                    row82.setHeightInPoints(65);
                    raw8cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 9:
                    Cell raw9cell1 = row82.createCell(0);
                    raw9cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raww9cell2 = row82.createCell(1);
                    raww9cell2.setCellValue("Мустақил баҳоловчи ташкилот нархи");
                    raww9cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raww9cell3 = row82.createCell(2);
                    row82.setHeightInPoints(65);
                    raww9cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 10:
                    Cell raw10cell1 = row82.createCell(0);
                    raw10cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raww10cell2 = row82.createCell(1);
                    raww10cell2.setCellValue("Е-баҳолаш нархи (Эксперт-2)");
                    raww10cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raww10cell3 = row82.createCell(2);
                    row82.setHeightInPoints(65);
                    raww10cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 11:
                    Cell raw11cell1 = row82.createCell(0);
                    raw11cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raww11cell2 = row82.createCell(1);
                    raww11cell2.setCellValue("Банк баҳолаш далолатномаси нархи");
                    raww11cell2.setCellStyle(Styles.getCellBasicStyleWithBackgroundGreen(workbook));
                    Cell raww11cell3 = row82.createCell(2);
                    row82.setHeightInPoints(65);
                    raww11cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 12:
                    Cell raw12cell1 = row82.createCell(0);
                    raw12cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw12cell2 = row82.createCell(1);
                    raw12cell2.setCellValue("Гаров эгасининг розилиги (паспорт)");
                    raw12cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw12cell3 = row82.createCell(2);
                    row82.setHeightInPoints(65);
                    raw12cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row83 = sheet.createRow(rowIndex);

        Cell raw83cell = row83.createCell(0);
        Cell raw83cell2 = row83.createCell(1);
        Cell raw83cell3 = row83.createCell(2);
        Cell raw83cell4 = row83.createCell(3);
        raw83cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw83cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw83cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw83cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        raw83cell.setCellValue("6");
        raw83cell2.setCellValue("Тақдим этилаётган таъминот маълумотлари");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + 10, 2, 2));

        for (int rowNum = 2; rowNum <= 12; rowNum++) {
            Row row84 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw2cell1 = row84.createCell(0);
                    raw2cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw2cell2 = row84.createCell(1);
                    raw2cell2.setCellValue("Манзили");
                    row84.setHeightInPoints(65);
                    raw2cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw2cell3 = row84.createCell(2);
                    raw2cell3.setCellValue("\"500 млн.сўмгача  125% номулкий - Қарз юки 20% гача;\n" +
                            "- 1 000 млн.сўмгача 75 % мулкий, 50% номулкий - Қарз юки 30% гача;\n" +
                            "- 2 000 млн.сўмгача 80 % мулкий, 45% номулкий - Қарз юки 30% гача;\n" +
                            "- Бошқа холларда 100% мулкий, 25% номулкий таъминотлар тақдим \nэтилиши лозим;\n" +
                            "Қуйидаги ҳолларда фақат 125% мулкий таъминот талаб этилади:\n" +
                            "- 1,0 млрд.сўмгача кредит ажратишда (Мижознинг қарз юки ёки \nбиргаликдаги қарз юки 100% дан юқори бўлганда;\n" +
                            "- 90 кундан ошган муддати ўтган қарздорликка 1 марта йўл қойган \nхолларда;\n" +
                            "- “БРБ” АТБ тизимида асосий ҳисобрақами мавжуд бўлмаган мижозларга \nкредит ажратишда;\n" +
                            "- 30 кундан ошган муддати ўтган қарздорликка 10 мартадан кўп бўлган \nхолларда;\n" +
                            "- Ваколат доирасида Кредит қўмитаси қарорига асосан доимий пул оқимига \nва ижобий кредит тарихига эга мижозларга кредит суммасининг 125% \nмиқдорида суғурта полиси ёки учинчи шахс кафиллиги билан кредит \nажратилишига рухсат этилиши мумкин (кредит миқдори 500.0 млн. \nсўмдан юқори бўлган лойиҳалар учун).\n" +
                            " Бунда, доимий пул оқимига эга мижозлар дейилганда жорий ва янги \nкредитлари билан бирга ҳисобланганда қарз юки 50% дан (50% ҳам киради) \nбаланд бўлмаслиги ва сўнги 12 ойда ҳисоб рақамида узликсиз тушуми \nмавжуд бўлиши (асосий фаолияти мавсумий бўлганда, тушуми \nузликсизлигига ўрнатилган талаб бундан мустасно) лозим.\n" +
                            "- Охирги 12 ой ичида мижознинг 50% дан юқори улушга эга таъсисчиси \nўзгарганда (яқин қариндошлар ўртасида ўзгариш бундан мустасно), камида \n125% фоиз ликвидли мол-мулк гарови тақдим этилиш лозим.\"\n");
                    raw2cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row84.createCell(0);
                    raw3cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw3cell2 = row84.createCell(1);
                    raw3cell2.setCellValue("Номи");
                    row84.setHeightInPoints(65);
                    raw3cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw3cell3 = row84.createCell(2);
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row84.createCell(0);
                    raw4cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw4cell2 = row84.createCell(1);
                    row84.setHeightInPoints(65);
                    raw4cell2.setCellValue("Гаров мулки эгаси");
                    raw4cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw4cell3 = row84.createCell(2);
                    raw4cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 5:
                    Cell raw5cell1 = row84.createCell(0);
                    raw5cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw5cell2 = row84.createCell(1);
                    row84.setHeightInPoints(65);
                    raw5cell2.setCellValue("Тегишлилиги тўғрисида хужжат");
                    raw5cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw5cell3 = row84.createCell(2);
                    raw5cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 6:
                    Cell raw6cell1 = row84.createCell(0);
                    raw6cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw6cell2 = row84.createCell(1);
                    row84.setHeightInPoints(65);
                    raw6cell2.setCellValue("Рўйхатдан ўтганлиги юзасидан кадастр кўчирмаси");
                    raw6cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw6cell3 = row84.createCell(2);
                    raw6cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 7:
                    Cell raw7cell1 = row84.createCell(0);
                    raw7cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw7cell2 = row84.createCell(1);
                    row84.setHeightInPoints(65);
                    raw7cell2.setCellValue("Рўйҳатда ҳеч ким турмаслиги");
                    raw7cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw7cell3 = row84.createCell(2);
                    raw7cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 8:
                    Cell raw8cell1 = row84.createCell(0);
                    raw8cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw8cell2 = row84.createCell(1);
                    raw8cell2.setCellValue("Таъқиқ мавжудлиги бўйича маълумотнома");
                    raw8cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw8cell3 = row84.createCell(2);
                    row84.setHeightInPoints(65);
                    raw8cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 9:
                    Cell raw9cell1 = row84.createCell(0);
                    raw9cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raww9cell2 = row84.createCell(1);
                    raww9cell2.setCellValue("Мустақил баҳоловчи ташкилот нархи");
                    raww9cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raww9cell3 = row84.createCell(2);
                    row84.setHeightInPoints(65);
                    raww9cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 10:
                    Cell raw10cell1 = row84.createCell(0);
                    raw10cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raww10cell2 = row84.createCell(1);
                    raww10cell2.setCellValue("Е-баҳолаш нархи (Эксперт-2)");
                    raww10cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raww10cell3 = row84.createCell(2);
                    row84.setHeightInPoints(65);
                    raww10cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 11:
                    Cell raw11cell1 = row84.createCell(0);
                    raw11cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raww11cell2 = row84.createCell(1);
                    raww11cell2.setCellValue("Банк баҳолаш далолатномаси нархи");
                    raww11cell2.setCellStyle(Styles.getCellBasicStyleWithBackgroundGreen(workbook));
                    Cell raww11cell3 = row84.createCell(2);
                    row84.setHeightInPoints(65);
                    raww11cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                case 12:
                    Cell raw12cell1 = row84.createCell(0);
                    raw12cell1.setCellStyle(Styles.getBackgroundBlueWithoutBorder(workbook));
                    Cell raw12cell2 = row84.createCell(1);
                    raw12cell2.setCellValue("Гаров эгасининг розилиги (паспорт)");
                    raw12cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw12cell3 = row84.createCell(2);
                    row84.setHeightInPoints(65);
                    raw12cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row107 = sheet.createRow(rowIndex);

        Cell row107cell = row107.createCell(0);
        Cell row107cell2 = row107.createCell(1);
        Cell row107cell3 = row107.createCell(2);
        Cell row107cell4 = row107.createCell(3);
        row107cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        row107cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        row107cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        row107cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        row107cell.setCellValue("9");
        row107cell2.setCellValue("Кредит қайтмаслик юзасидан суғурта ташкилотининг суғурта полиси тақдим этилаётганда");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        sheet.addMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 5, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 5, 1, 1));

        for (int rowNum = 2; rowNum <= 7; rowNum++) {
            Row row108 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw22cell1 = row108.createCell(0);
                    raw22cell1.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw22cell2 = row108.createCell(1);
                    raw22cell2.setCellValue("Суғурта ташкилоти номи");
                    raw22cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw22cell3 = row108.createCell(2);
                    break;

                case 3:
                    Cell raw3cell1 = row108.createCell(0);
                    raw3cell1.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell2 = row108.createCell(1);
                    raw3cell2.setCellValue("Молиявий барқарорлиги ");
                    raw3cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw3cell3 = row108.createCell(2);
                    raw3cell3.setCellValue("\"Риск\" департаменти маълумотига асосан");
                    raw3cell3.setCellStyle(Styles.getItalicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row108.createCell(0);
                    raw4cell1.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell2 = row108.createCell(1);
                    raw4cell2.setCellValue("Суғурта суммаси");
                    raw4cell2.setCellStyle(Styles.getCellStyle(workbook));
                    Cell raw4cell3 = row108.createCell(2);
                    break;
                case 5:
                    Cell raw5cell1 = row108.createCell(0);
                    raw5cell1.setCellValue("7");
                    raw5cell1.setCellStyle(Styles.getBackground(workbook));
                    Cell raw5cell2 = row108.createCell(1);
                    raw5cell2.setCellValue("Жами таъминотлар суммаси");
                    raw5cell2.setCellStyle(Styles.getBackground(workbook));
                    Cell raw5cell3 = row108.createCell(2);
                    raw5cell3.setCellValue("Жами таъминот");
                    raw5cell3.setCellStyle(Styles.getBackground(workbook));
                    break;

                case 6:
                    Cell raw6cell1 = row108.createCell(0);
                    raw6cell1.setCellStyle(Styles.getBackground(workbook));
                    Cell raw6cell2 = row108.createCell(1);
                    raw6cell2.setCellStyle(Styles.getBackground(workbook));
                    Cell raw6cell3 = row108.createCell(2);
                    raw6cell3.setCellValue("Мулкий таъминот");
                    raw6cell3.setCellStyle(Styles.getBackground(workbook));
                    break;

                case 7:
                    Cell raw7cell1 = row108.createCell(0);
                    raw7cell1.setCellStyle(Styles.getBackground(workbook));
                    Cell raw7cell2 = row108.createCell(1);
                    raw7cell2.setCellStyle(Styles.getBackground(workbook));
                    Cell raw7cell3 = row108.createCell(2);
                    raw7cell3.setCellValue("Номулкий таъминот");
                    raw7cell3.setCellStyle(Styles.getBackground(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row114 = sheet.createRow(rowIndex);

        Cell row114cell = row114.createCell(0);
        Cell row114cell2 = row114.createCell(1);
        Cell row114cell3 = row114.createCell(2);
        Cell row114cell4 = row114.createCell(3);
        row114cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        row114cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        row114cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        row114cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        row114cell.setCellValue("8");
        row114cell2.setCellValue("Банк томонидан берилган хужжатлар");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        for (int rowNum = 2; rowNum <= 4; rowNum++) {
            Row row115 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw22cell1 = row115.createCell(0);
                    raw22cell1.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw22cell2 = row115.createCell(1);
                    raw22cell2.setCellValue("Банк хулоса");
                    raw22cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw22cell3 = row115.createCell(2);
                    raw22cell3.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 3:
                    Cell raw3cell1 = row115.createCell(0);
                    raw3cell1.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw3cell2 = row115.createCell(1);
                    raw3cell2.setCellValue("Юрист хулоса");
                    raw3cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw3cell3 = row115.createCell(2);
                    raw3cell3.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;

                case 4:
                    Cell raw4cell1 = row115.createCell(0);
                    raw4cell1.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw4cell2 = row115.createCell(1);
                    raw4cell2.setCellValue("Қўмитага хат");
                    raw4cell2.setCellStyle(Styles.getCellBasicStyle(workbook));
                    Cell raw4cell3 = row115.createCell(2);
                    raw4cell3.setCellStyle(Styles.getCellBasicStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        Row row118 = sheet.createRow(rowIndex);

        Cell row118cell = row118.createCell(0);
        Cell row118cell2 = row118.createCell(1);
        Cell row118cell3 = row118.createCell(2);
        Cell row118cell4 = row118.createCell(3);
        row118cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        row118cell2.setCellStyle(Styles.getBackgroundBlue(workbook));
        row118cell3.setCellStyle(Styles.getBackgroundBlue(workbook));
        row118cell4.setCellStyle(Styles.getBackgroundBlue(workbook));
        row118cell.setCellValue("9");
        row118cell2.setCellValue("Андеррайтер ХУЛОСАСИ");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 1;

        Row row119 = sheet.createRow(rowIndex);

        Cell row119cell = row119.createCell(0);
        Cell row119cell2 = row119.createCell(1);
        Cell row119cell3 = row119.createCell(2);
        Cell row119cell4 = row119.createCell(3);
        row119cell.setCellStyle(Styles.getBackgroundBlue(workbook));
        row119cell2.setCellStyle(Styles.getLeftCellStyleBlue(workbook));
        row119cell3.setCellStyle(Styles.getLeftCellStyleBlue(workbook));
        row119cell4.setCellStyle(Styles.getLeftCellStyleBlue(workbook));
        row119cell.setCellValue("1.");
        row119.setHeightInPoints(400);
        row119cell2.setCellValue("         1. \"SHODLIK TECHNO\" МЧЖга 6 ой имтиёли давр билан, 36 ой муддатга, йиллик 30,0 фоиз устама тўлаш шартлари асосида 3 500 000 000,0 сўм миқдорида кредит маблағлари ажратиш  \"Universal\" кредит маҳсулоти паспорти талабларига мос ҳисобланади.\n" +
                "\n" +
                "        2. Мазкур лойиҳа Кредит қўмитаси ваколати доирасида ҳисобланиб, шунингдек, мижознинг бугунги кундаги мажбуриятлари 3,0 млрд сўмлигини инобатга олиб кредит ажратиш масаласини кўриб чиқиш ва якуний қарор қабул қилиш учун Кредит қўмитаси муҳокамасига киритилмоқда.\n" +
                "\n" +
                "        3. Лойиҳа юзасидан Риск менджмент департаменти ҳулосаси илова қилинади. \n" +
                "\n" +
                "        4. Тақдим этилган ҳужжатларни ҳаққонийлигига, гаров мулкларини тўғри баҳоланиши ва кредит йиғмажилдидаги ҳужжатларни қонуний тарзда расмийлаштирилишига БХМ/БХО раҳбари, баҳолаш \nкомиссияси ва тегишли масъул ходимлар жавобгар ҳисобланади.\"\t\t\n");
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
        rowIndex += 2;

        for (int rowNum = 2; rowNum <= 8; rowNum++) {
            Row row120 = sheet.createRow(rowIndex);

            switch (rowNum) {
                case 2:
                    Cell raw22cell1 = row120.createCell(0);
                    raw22cell1.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw22cell2 = row120.createCell(1);
                    row120.setHeightInPoints(60);
                    raw22cell2.setCellValue("Юридик шахслар андеррайтинги бошқармаси \nетакчи менежери");
                    raw22cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw22cell3 = row120.createCell(2);
                    raw22cell3.setCellStyle(Styles.getLeftCellStyle(workbook));
                    break;

                case 3:
                    row120.setHeightInPoints(60);
                    break;

                case 4:
                    Cell raw3cell1 = row120.createCell(0);
                    raw3cell1.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw3cell2 = row120.createCell(1);
                    raw3cell2.setCellValue("Келишилди:");
                    raw3cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw3cell3 = row120.createCell(2);
                    raw3cell3.setCellStyle(Styles.getLeftCellStyle(workbook));
                    break;

                case 5:
                    row120.setHeightInPoints(60);
                    break;

                case 6:
                    Cell raw4cell1 = row120.createCell(0);
                    row120.setHeightInPoints(60);
                    raw4cell1.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw4cell2 = row120.createCell(1);
                    raw4cell2.setCellValue("Юридик шахслар андеррайтинги бошқармаси \nбошлиғи");
                    raw4cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw4cell3 = row120.createCell(2);
                    raw4cell3.setCellStyle(Styles.getLeftCellStyle(workbook));
                    break;

                case 7:
                    row120.setHeightInPoints(60);
                    break;

                case 8:
                    Cell raw5cell1 = row120.createCell(0);
                    raw5cell1.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw5cell2 = row120.createCell(1);
                    row120.setHeightInPoints(70);
                    raw5cell2.setCellValue("Кредитларни маъқуллаш ва лойиҳаларни \n" +
                            "молиялаштириш департаменти директори\n");
                    raw5cell2.setCellStyle(Styles.getLeftCellStyle(workbook));
                    Cell raw5cell3 = row120.createCell(2);
                    raw5cell3.setCellStyle(Styles.getLeftCellStyle(workbook));
                    break;
            }

            rowIndex += 1;
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();

        return outputStream.toByteArray();
    }
}
