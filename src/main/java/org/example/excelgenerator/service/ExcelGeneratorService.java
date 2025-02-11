package org.example.excelgenerator.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.excelgenerator.dto.request.ExcelRequest;
import org.example.excelgenerator.fonts.Styles;
import org.example.excelgenerator.helper.StaticData;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@Service
public class ExcelGeneratorService {
    public byte[] generateExcel(ExcelRequest request) throws IOException {
        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Data");

        CellStyle boldStyle = Styles.getCellStyle(workbook);

        CellStyle italicStyle = Styles.getItalicStyle(workbook);

        CellStyle styleWithoutBorder = Styles.getCellStyle(workbook);

        String[][] staticData1 = StaticData.staticData1;

        // 🟢 Fill static columns (A, B, C)
        for (int rowIndex = 0; rowIndex < staticData1.length; rowIndex++) {
            Row row = sheet.createRow(rowIndex);

            for (int colIndex = 0; colIndex < 4; colIndex++) {
                if(rowIndex == 0) {
                    row.setHeightInPoints(73.5f);
                }
                if(rowIndex == 15) {
                    row.setHeightInPoints(327f);
                }
                if(rowIndex == 16) {
                    row.setHeightInPoints(180f);
                }
                if(rowIndex == 18) {
                    row.setHeightInPoints(300f);
                }
                if(rowIndex == 23) {
                    row.setHeightInPoints(48f);
                }
                else {
                    row.setHeightInPoints(32.25f);
                }

                Cell cell = row.createCell(colIndex);
                cell.setCellValue(staticData1[rowIndex][colIndex]);

                if (rowIndex == 0) {
                    cell.setCellStyle(styleWithoutBorder);
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
                    break;
                }

                if(rowIndex == 1) {
                    cell.setCellStyle(Styles.getCellStyle(workbook));
                    if(colIndex == 2) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        break;
                    }
                    continue;
                }

                if(rowIndex > 1 && rowIndex < 7) {
                    if(colIndex == 1) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    if(colIndex == 2) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        break;
                    }
                }

                if(rowIndex == 7) {
                    if(colIndex == 1) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    if(colIndex == 2) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    if(colIndex == 3) {
                        cell.setCellStyle(Styles.getBottomBorder(workbook));
                        continue;
                    }

                    cell.setCellStyle(Styles.getCellStyle(workbook));
                    continue;
                }

                if(rowIndex == 8) {
                    cell.setCellStyle(Styles.getCellStyle(workbook));
                    sheet.addMergedRegion(new CellRangeAddress(8, 8, 0, 3));
                    break;
                }

                if(rowIndex == 9) {
                    if(colIndex == 3) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    cell.setCellStyle(Styles.getCellStyle(workbook));
                    continue;
                }

                if(rowIndex == 10) {
                    if(colIndex == 0) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    sheet.addMergedRegion(new CellRangeAddress(10, 10, 1, 3));
                    cell.setCellStyle(Styles.getCellStyle(workbook));
                    break;
                }

                if (rowIndex >= 11 && rowIndex <= 19) {
                    if(colIndex == 1) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    cell.setCellStyle(italicStyle);
                    continue;
                }

                if(rowIndex == 20) {
                    if(colIndex <= 1) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    sheet.addMergedRegion(new CellRangeAddress(20, 20, 2, 3));
                    cell.setCellStyle(Styles.getItalicStyle(workbook));
                    break;
                }

                if (rowIndex >= 21 && rowIndex <= 23) {
                    if(colIndex <= 1) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    cell.setCellStyle(italicStyle);
                    continue;
                }

                if(rowIndex == 24) {
                    if(colIndex == 0) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    sheet.addMergedRegion(new CellRangeAddress(24, 24, 1, 3));
                    cell.setCellStyle(Styles.getCellStyle(workbook));
                    break;
                }

                if(rowIndex == 31) {
                    if(colIndex == 0) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    sheet.addMergedRegion(new CellRangeAddress(31, 31, 1, 3));
                    cell.setCellStyle(Styles.getCellStyle(workbook));
                    break;
                }

                if(rowIndex == 39) {
                    if(colIndex == 0) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    sheet.addMergedRegion(new CellRangeAddress(39, 39, 1, 3));
                    cell.setCellStyle(Styles.getCellStyle(workbook));
                    break;
                }

                if(rowIndex == 44) {
                    if(colIndex == 0) {
                        cell.setCellStyle(Styles.getCellStyle(workbook));
                        continue;
                    }
                    sheet.addMergedRegion(new CellRangeAddress(44, 44, 1, 3));
                    cell.setCellStyle(Styles.getCellStyle(workbook));
                    break;
                }


                cell.setCellStyle(boldStyle);
            }
        }

        sheet.setColumnWidth(0, (int) (8.43 * 256));
        sheet.setColumnWidth(1, (int) (103.29 * 256));
        sheet.setColumnWidth(2, (int) (149.71 * 256));
        sheet.setColumnWidth(3, (int) (147 * 256));

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();

        return outputStream.toByteArray();
    }
}
