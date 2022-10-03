package com.example.translator;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.hc.core5.http.ParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Component;
import org.springframework.web.context.annotation.SessionScope;

import java.io.IOException;
import java.io.InputStream;
import java.io.Serializable;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;


// TODO: Добавить возможность выбора колонки для перевода
//@Component
//@SessionScope
//@ToString
//@Scope(value = "session",proxyMode = ScopedProxyMode.TARGET_CLASS)
@Getter
@Setter
public class ExcelReader implements Serializable {
    public int numberOfRows = 0;
    public int numberOfReadyRows = 0;

//    @Resource(name = "sessionScopedBean2")
//    private ExcelWriter writer;
//    @Autowired
//    private ExcelController controller;

    ExcelWriter excelWriter = new ExcelWriter();

    public byte[] getTranslatedFile(InputStream inputStream, String uuid, int columnIndex) throws IOException {
        excelWriter.init();
        return Files.readAllBytes(Path.of(readTest(inputStream, uuid, columnIndex)));
    }

    public String readTest(InputStream inputStream, String uuid, int columnIndex) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
            long startTime = System.currentTimeMillis();

            int numberOfSheets = workbook.getNumberOfSheets();

            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                for (Row row : sheet) {
                    Cell cell = row.getCell(columnIndex);
                    if (cell != null && cell.getCellType() != CellType.BLANK) {
                        setNumberOfRows(getNumberOfRows() + 1);
                        System.out.println(this);
                    }
                }
//                ExcelController.numberOfStandards += workbook.getSheetAt(i).getPhysicalNumberOfRows();
            }

            for (int i = 0; i < numberOfSheets; i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);

                for (Row row : sheet) {
                    Cell cell = row.getCell(columnIndex);
                    if (cell != null && cell.getCellType() != CellType.BLANK) {
                        excelWriter.fillData(cell.toString());
//                        Thread.sleep(500);
                        setNumberOfReadyRows(getNumberOfReadyRows() + 1);
                    }
//                    String standardName = readStandardNameFromCell(row);
//
//                    if (standardName != null && !standardName.isEmpty()) {
//                        ExcelController.numberOfReadyRows++;
//                        log.debug("Считана ячейка с названием стандарта: {}", standardName);
//                        String standardId = StringModifier.getStandardIdFromStandardName(standardName);
//                        Parser p = parserMap.get(StringModifier.getStandardIdFromStandardName(standardId));
//                        if (p == null) {
//                            createNewParser(standardId);
//                        }
//                        p = parserMap.get(StringModifier.getStandardIdFromStandardName(standardId));
//                        if (p != null) {
//                            List<Standard> standards = new ArrayList<Standard>(p.getStandardWithStory(standardName));
//                            for (Standard standard : standards) {
//                                writer.fillData(standardName, standard);
//                                log.debug("Ячейка со стандартом: {} была успешно считана", standardName);
//                            }
//                        } else {
//                            log.error("Произошла ошибка во время заполнения стандарта: {}", standardName);
//                            writer.fillError(standardName);
//                        }
//                    }
                }
            }

            long endTime = System.currentTimeMillis();
            System.out.println("Total execution time: " + (endTime-startTime) + "ms");

            return excelWriter.getFile(uuid);
//            return "assd";
        } catch (IOException e) {
            throw new IOException(e);
        } catch (InterruptedException e) {
            e.printStackTrace();
        } catch (URISyntaxException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        } finally {
//            for (Parser p : parserMap.values()) {
//                p.stop();
//            }
        }
        return uuid;
    }

    public void createNewParser(String parserType) {
        switch(parserType) {
//            case ("IEC") -> parserMap.put("IEC", new IECParser());
//            case ("ISO") -> parserMap.put("ISO", new ISOParser());
//            case ("IAEA") -> parserMap.put("IAEA", new IAEAParser());
//            case ("IEEE") -> parserMap.put("IEEE", new IEEEParser());
//            case ("MSZ") -> parserMap.put("MSZ", new MSZParser());
        }
    }

    private String readStandardNameFromCell(Row row) {
        Iterator<Cell> cellIterator = row.cellIterator();
        String standardId = cellIterator.hasNext() ? cellIterator.next().toString() : null;
        String standardYear = cellIterator.hasNext() ? cellIterator.next().toString() : null;
        String type = "";
//        if (standardId != null) {
//            type = StringModifier.getStandardIdFromStandardName(standardId);
//        }

        if (standardId == null) {
            return null;
        }

        if (standardYear == null) {
            return standardId;
        }

        if (type.equals("IEEE") && !standardId.contains("-") && !standardYear.isEmpty()) {
            return standardId + "-" + standardYear;
        } else if (!type.equals("IEEE") && !standardId.contains(":") && !standardYear.isEmpty()) {
            return standardId + ":" + standardYear;
        }
        return standardId;
    }
}
