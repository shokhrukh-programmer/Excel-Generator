//package org.example.excelgenerator.service;
//
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.example.excelgenerator.dto.request.ExcelRequest;
//import org.example.excelgenerator.fonts.Styles;
//import org.example.excelgenerator.helper.StaticData;
//import org.springframework.stereotype.Service;
//
//import java.io.ByteArrayOutputStream;
//import java.io.IOException;
//
//@Service
//public class ExcelGeneratorService {
//    public byte[] generateExcel(ExcelRequest request) throws IOException {
//        Workbook workbook = new XSSFWorkbook();
//
//        Sheet sheet = workbook.createSheet("Data");
//
//        CellStyle boldStyle = Styles.getCellStyle(workbook);
//
//        CellStyle italicStyle = Styles.getItalicStyle(workbook);
//
//        CellStyle styleWithoutBorder = Styles.getCellStyle(workbook);
//
//        String[][] staticData1 = StaticData.staticData1;
//        String[] dynamic = request.toArray();
//        int i = 0;
//        // 🟢 Fill static columns (A, B, C)
//        for (int rowIndex = 0; rowIndex < staticData1.length; rowIndex++) {
//            Row row = sheet.createRow(rowIndex);
//
//            if(rowIndex == 0  || rowIndex == 52 || rowIndex == 68 || rowIndex == 69 || rowIndex == 44 ) {
//                row.setHeightInPoints(73.5f);
//            }
//            else if(rowIndex == 15 || rowIndex == 118) {
//                row.setHeightInPoints(327f);
//            }
//            else if(rowIndex == 16) {
//                row.setHeightInPoints(180f);
//            }
//            else if(rowIndex == 18) {
//                row.setHeightInPoints(300f);
//            } else if(rowIndex == 22 || rowIndex == 27 || rowIndex == 35 || rowIndex == 37 || rowIndex == 38 || rowIndex == 39 ||
//            rowIndex == 55 || rowIndex == 56 || rowIndex == 54 || rowIndex == 59 || rowIndex == 61 || rowIndex == 65 || rowIndex == 66 ) {
//                row.setHeightInPoints(60f);
//            } else if(rowIndex >= 71 && rowIndex <= 81 || rowIndex >= 83 && rowIndex <= 93 || rowIndex >= 95 && rowIndex <= 105 || rowIndex == 120 || rowIndex == 124
//                    || rowIndex == 127) {
//                row.setHeightInPoints(60f);
//            }
//            else {
//                row.setHeightInPoints(32.25f);
//            }
//            for (int colIndex = 0; colIndex < 4; colIndex++) {
//
//
//                Cell cell = row.createCell(colIndex);
////|| rowIndex == 127
//                if(colIndex != 3) {
//                    cell.setCellValue(staticData1[rowIndex][colIndex]);
//                } else {
//                    if(rowIndex >= 11 && rowIndex <= 19 || rowIndex >= 21 && rowIndex <= 23 || rowIndex >= 25 && rowIndex <= 30 || rowIndex >= 32 && rowIndex <= 39
//                    || rowIndex >= 41 && rowIndex <= 44 || rowIndex >= 46 && rowIndex <= 52 || rowIndex >= 54 && rowIndex <= 57 || rowIndex >= 59 && rowIndex <= 62
////                            || rowIndex >= 64 && rowIndex <= 66 || rowIndex >= 68 && rowIndex <= 69  || rowIndex >= 71 && rowIndex <= 81 || rowIndex >= 83 && rowIndex <= 93
////                            || rowIndex >= 95 && rowIndex <= 105 || rowIndex >= 107 && rowIndex <= 112 || rowIndex >= 114 && rowIndex <= 116 || rowIndex == 120 || rowIndex == 122
//                    || rowIndex == 124 ) {
//                        cell.setCellValue(dynamic[i++]);
//                    }
//                }
//
//                if (rowIndex == 0) {
//                    cell.setCellStyle(styleWithoutBorder);
//                    sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
//                    break;
//                }
//
//                if(rowIndex == 1) {
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        break;
//                    }
//                    continue;
//                }
//
//                if(rowIndex > 1 && rowIndex < 7) {
//                    if(colIndex == 1) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        break;
//                    }
//                }
//
//                if(rowIndex == 7) {
//                    if(colIndex == 1) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    if(colIndex == 3) {
//                        cell.setCellStyle(Styles.getBottomBorder(workbook));
//                        continue;
//                    }
//
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 8) {
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    sheet.addMergedRegion(new CellRangeAddress(8, 8, 0, 3));
//                    break;
//                }
//
//                if(rowIndex == 9) {
//                    if(colIndex == 3) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 10) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(10, 10, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if (rowIndex >= 11 && rowIndex <= 19) {
//                    if(colIndex == 1) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(italicStyle);
//                    continue;
//                }
//
//                if(rowIndex == 20) {
//                    if(colIndex <= 1) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(20, 20, 2, 3));
//                    cell.setCellStyle(Styles.getItalicStyle(workbook));
//                    break;
//                }
//
//                if (rowIndex >= 21 && rowIndex <= 23) {
//                    if(colIndex <= 1) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(italicStyle);
//                    continue;
//                }
//
//                if(rowIndex == 24) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(24, 24, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//
//                if(rowIndex == 31) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(31, 31, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 32 && rowIndex <= 39) {
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getItalicStyle(workbook));
//                        continue;
//                    }
//
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 40) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(40, 40, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 41 && rowIndex <= 44) {
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getItalicStyle(workbook));
//                        continue;
//                    }
//
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 45) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(45, 45, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 46 && rowIndex <= 52) {
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 53) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(53, 53, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 54 && rowIndex <= 57) {
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getItalicStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 58) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(58, 58, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 59 && rowIndex <= 62) {
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getItalicStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 63) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(63, 63, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 64 && rowIndex <= 66) {
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getItalicStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 67) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(67, 67, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 68 && rowIndex <= 69) {
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getItalicStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 70) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(70, 70, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 71 && rowIndex <= 81) {
//                    if(colIndex == 1) {
//                        cell.setCellStyle(Styles.getCellBasicStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 82) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(82, 82, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 83 && rowIndex <= 93) {
//                    if(colIndex == 1) {
//                        cell.setCellStyle(Styles.getCellBasicStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 94) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(94, 94, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 95 && rowIndex <= 105) {
//                    if(colIndex == 1) {
//                        cell.setCellStyle(Styles.getCellBasicStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 106) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(106, 106, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 107 && rowIndex <= 109) {
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getItalicStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex >= 110 && rowIndex <= 112) {
//                    if(colIndex == 2) {
//                        cell.setCellStyle(Styles.getItalicStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 113) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(113, 113, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex >= 114 && rowIndex <= 116) {
//                    if(colIndex == 1) {
//                        cell.setCellStyle(Styles.getCellBasicStyle(workbook));
//                        continue;
//                    }
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    continue;
//                }
//
//                if(rowIndex == 117) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(117, 117, 1, 3));
//                    cell.setCellStyle(Styles.getCellStyle(workbook));
//                    break;
//                }
//
//                if(rowIndex == 118) {
//                    if(colIndex == 0) {
//                        cell.setCellStyle(Styles.getCellStyle(workbook));
//                        continue;
//                    }
//                    sheet.addMergedRegion(new CellRangeAddress(118, 118, 1, 3));
//                    cell.setCellStyle(Styles.getLeftCellStyle(workbook));
//                    break;
//                }
//
//                cell.setCellStyle(boldStyle);
//            }
//        }
//
//        sheet.setColumnWidth(0, (int) (8.43 * 256));
//        sheet.setColumnWidth(1, (int) (150.29 * 256));
//        sheet.setColumnWidth(2, (int) (200.71 * 256));
//        sheet.setColumnWidth(3, (int) (147 * 256));
//
//        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
//        workbook.write(outputStream);
//        workbook.close();
//
//        return outputStream.toByteArray();
//    }
//}
