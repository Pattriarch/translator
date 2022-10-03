package com.example.translator;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.client5.http.impl.classic.CloseableHttpClient;
import org.apache.hc.client5.http.impl.classic.CloseableHttpResponse;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.core5.http.ParseException;
import org.apache.hc.core5.http.io.entity.EntityUtils;
import org.apache.hc.core5.http.io.entity.StringEntity;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URLEncoder;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

import static java.util.Map.entry;

// TODO: Попробовать ускорить работу XLSX.XLSXWriter's

@Slf4j
//@Component
//@SessionScope
public class ExcelWriter {
    Map<String, Sheet> sheetMap = new HashMap<>();
    Map<String, Row> headerRowsMap = new HashMap<>();
    Map<String, String[]> columnHeadingsMap = new HashMap<>();
    Map<String, Integer> indexMap = new HashMap<>();
    Workbook workbook;
    Font headerFont;
    CellStyle headerStyle;

    String[] translated = new String[]{
            "Старое название",
            "Новое название"
    };

    Map<String, String[]> zakupkiHeadings = Map.ofEntries(
            entry("translated", translated)
    );

    List<String> zakupkiNames = List.of(
            "translated"
    );

    public void init() {
        try {
            workbook = new XSSFWorkbook();

            createSheets();
            createHeadings();
            setHeaderFont();
            setHeaderStyle();
            createRows();
            buildErrorSheet();

            for (String standardName : zakupkiNames) {
                createCellsForRow(columnHeadingsMap.get(standardName), headerRowsMap.get(standardName));
            }

            freezePanes();

            sheetMap.get("translated").setColumnWidth(0, 15000);
            sheetMap.get("translated").setColumnWidth(1, 15000);
//            headerRowsMap.get("gov")

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void createSheets() {
        log.debug("Создание листов в excel-файле...");
        for (String standardName : zakupkiNames) {
            sheetMap.put(standardName, workbook.createSheet(standardName));
        }
    }

    public void createHeadings() {
        log.debug("Создание заголовков в excel-файле...");
        for (String standardName : zakupkiNames) {
            columnHeadingsMap.put(standardName, zakupkiHeadings.get(standardName));
        }
    }

    public void setHeaderFont() {
        log.debug("Установка шрифта в excel-файле...");
        headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.BLACK.index);
    }

    public void setHeaderStyle() {
        log.debug("Установка стиля для заголовков в excel-файле...");
        headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);
    }

    public void createRows() {
        log.debug("Создание строк в excel-файле для каждого листка...");
        for (String standardName : zakupkiNames) {
            Row row = sheetMap.get(standardName).createRow(0);
            row.setHeight((short)(624));
            headerRowsMap.put(standardName, row);
            indexMap.put(standardName, 1);
        }
    }

    public void buildErrorSheet() {
        log.debug("Создание листка с ошибками в excel-файле...");
        sheetMap.put("Errors", workbook.createSheet("Errors"));
        headerRowsMap.put("Errors", sheetMap.get("Errors").createRow(0));

        Row row = headerRowsMap.get("Errors");
        row.createCell(0).setCellValue("Below standards with which an error occurred during parsing");

        indexMap.put("Errors", 1);

        Sheet errorsSheet = sheetMap.get("Errors");
        errorsSheet.autoSizeColumn(0);
    }

    public void createCellsForRow(String[] columnHeadings, Row row) {
        log.debug("Установка заголовков в строке...");
        for (int i = 0, j = 0; i < columnHeadings.length; i++) {
//            if (params.get(columnHeadings[i].substring(0, 1) + columnHeadings[i].substring(1).replaceAll(" ", "")) != null &&
//                    params.get(columnHeadings[i].substring(0, 1) + columnHeadings[i].substring(1).replaceAll(" ", ""))) {
                Cell cell = row.createCell(j);
                cell.setCellValue(columnHeadings[i]);
                cell.setCellStyle(headerStyle);
                j++;
//            }
        }
    }

    public void freezePanes() {
        log.debug("Замораживаем строку с названиями каждого столбца...");
        for (Sheet value : sheetMap.values()) {
            value.createFreezePane(0, 1);
        }
    }

//    private void autoSizeColumns(String[] columnHeadings, Sheet sheet) {
//        log.debug("Установка авто-сайза колонок в excel-файле...");
//        for (int i = 0; i < columnHeadings.length; i++) {
//            sheet.autoSizeColumn(i);
//        }
//    }

    public int getCellIndex(String method, String standardName) {
        log.debug("Получение ячейки для метода {} и стандарта {}...", method, standardName);
        Row row = headerRowsMap.get(standardName);
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            if (cell.toString().replaceAll("\\s+","").equalsIgnoreCase(method.substring(3))) {
                return cell.getColumnIndex();
            }
        }
        return 0;
    }

    public CellStyle getCommonStyle() {
        CellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setWrapText(true);
        return style;
    }

    public CellStyle getDataStyle() {
        CellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);
        return style;
    }

    @Getter
    @Setter
    @NoArgsConstructor
    static class TranslateResponse {
        private String status;
        private String translatedText;
    }

    List<String> noTranslateWords = List.of(
            "i.", "ii.", "iii.", "iv.", "v.", "vi.", "vii.", "viii.", "ix.", "x.",
            "xi.", "xii.", "xiii.", "xiv.", "xv.", "xvi.", "xvii.", "xviii.", "xix.", "xx.",
            "•"
    );

    public void fillData(String data) throws URISyntaxException, IOException, InterruptedException, ParseException {
        String oldData = data;
        data = data.trim().replaceAll("\n", "").replaceAll(" +", " ");
        data = String.join(" ", Arrays.stream(data.split(" ")).map(string -> {
            if (string.isEmpty()) return string;
            if (
                noTranslateWords.contains(string.toLowerCase())                                   ||
                noTranslateWords.contains(string.toLowerCase().substring(0, string.length() - 1)) ||
                string.toLowerCase().replaceAll("[a-z]{1}[.)]", "").equals("")
            ) {
                string = " «_" + string.trim() + "_» ";
            }
            return string;
        }).toList());
//        System.out.println(data);

        URI uri = new URI("https://script.google.com/macros/s/AKfycbz6bWkw6DA9kVXYA74s6oShQd5UhnJTFXjnRw1y-xczzMJC3KeZtEsF5ofmch-GG3M1-A/exec?text=" + URLEncoder.encode(data, Charset.defaultCharset()));

        CloseableHttpClient client = HttpClients.createDefault();
        HttpPost httpPost = new HttpPost(uri);

        String payload = "{\"text\":"+data.replaceAll("\n", "")+"}";
        StringEntity entity = new StringEntity(payload);
        httpPost.setEntity(entity);
        httpPost.setHeader("Accept", "application/json");
        httpPost.setHeader("Content-type", "application/json");

        CloseableHttpResponse response = client.execute(httpPost);
        String responseBody = EntityUtils.toString(response.getEntity(), StandardCharsets.UTF_8);
//        System.out.println(responseBody);

        ObjectMapper objectMapper = new ObjectMapper();
        TranslateResponse response2 = objectMapper.readValue(responseBody, new TypeReference<>() {});
        response2.setTranslatedText(response2.getTranslatedText().replaceAll("(`\\s*)_", "`_").replaceAll("(_\\s*)`", "_`").replaceAll("(«\\s*)_", "`_").replaceAll("(_\\s*)»", "_`"));
//        response2.setTranslatedText(response2.getTranslatedText().replaceAll("[«»]", ""));
//        System.out.println(response2.getTranslatedText());
        client.close();
        String sheetName = "translated";
        Row row = sheetMap.get(sheetName).createRow(indexMap.get(sheetName));
        indexMap.put(sheetName, indexMap.get(sheetName) + 1);

        Cell originalNameCell = row.createCell(0);
        originalNameCell.setCellValue(oldData);
        originalNameCell.setCellStyle(getCommonStyle());

        response2.setTranslatedText(String.join(" ", Arrays.stream(response2.getTranslatedText().split(" ")).map((string) -> {
            if (string.contains("`_")) {
                string = string.replaceAll("`_", "\n    ");
            }
            if (string.contains("_`")) {
                string = string.replaceAll("_`", "");
            }
            return string;
        }).toList()));

        Cell translatedNameCell = row.createCell(1);
//        response2.setTranslatedText(response2.getTranslatedText());
        translatedNameCell.setCellValue(response2.getTranslatedText());
        translatedNameCell.setCellStyle(getCommonStyle());

//        Cell objectOfPurchaseCell = row.createCell(1);
//        objectOfPurchaseCell.setCellValue(zakupka.getObjectOfPurchase());
//        objectOfPurchaseCell.setCellStyle(getCommonStyle());
//
//        Cell customerCell = row.createCell(2);
//        customerCell.setCellValue(zakupka.getCustomer());
//        customerCell.setCellStyle(getCommonStyle());
//
//        Cell placementDateCell = row.createCell(3);
//        placementDateCell.setCellValue(zakupka.getPlacementDate());
//        placementDateCell.setCellStyle(getDataStyle());
//
//        Cell updateDateCell = row.createCell(4);
//        updateDateCell.setCellValue(zakupka.getUpdateDate());
//        updateDateCell.setCellStyle(getDataStyle());
//
//        Cell priceCell = row.createCell(5);
//        priceCell.setCellValue(zakupka.getPrice());
//        priceCell.setCellStyle(getCommonStyle());
//
//        Cell searchPatternCell = row.createCell(6);
//        searchPatternCell.setCellValue(zakupka.getSearchPattern());
//        searchPatternCell.setCellStyle(getCommonStyle());
//
//        // ссылка на закупку
//        Hyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
//        link.setAddress(zakupka.getUrlOfZakupka());
//        System.out.println("ADDR: " + zakupka.getUrlOfZakupka());
//
//        Cell cell = row.createCell(7);
//        cell.setCellValue("Ссылка на закупку");
//        cell.setHyperlink(link);
//
//        CellStyle hLinkStyle = workbook.createCellStyle();
//        final Font hLinkFont = workbook.createFont();
//        hLinkFont.setFontName("Ariel");
//        hLinkFont.setUnderline(Font.U_SINGLE);
//        hLinkFont.setColor(IndexedColors.BLUE.getIndex() );
//        hLinkStyle.setFont(hLinkFont);
//        hLinkStyle.setVerticalAlignment(VerticalAlignment.TOP);
//
//        cell.setCellStyle(hLinkStyle);
//
//        for (int i = 8; i < 8 + zakupka.getListOfDocumentLinks().size(); i++) {
//            link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
//            link.setAddress(zakupka.getListOfDocumentLinks().get(i - 8));
//            cell = row.createCell(i);
//            cell.setHyperlink(link);
//                cell.setCellValue(zakupka.getListOfDocumentNames().get(i - 8));
//                sheetMap.get(zakupkaName).autoSizeColumn(i);
//            cell.setCellStyle(hLinkStyle);
//        }
    }

    synchronized public void fillError(String standardName) {
        log.error("Заполнение ошибочного стандарта под названием {}...", standardName);
        int index = indexMap.get("Errors");
        Row row = sheetMap.get("Errors").createRow(index);
        indexMap.put("Errors", index + 1);
        row.createCell(0).setCellValue(standardName);
    }

    public String getFile(String uuid) {
        log.debug("Получение готового файла...");
        try(ByteArrayOutputStream fileOut = new ByteArrayOutputStream()) {
            workbook.write(fileOut);
            // Здесь мы записываем файл на сервер
            Files.createDirectories(Paths.get(System.getProperty("user.home") + File.separator + "translate" + File.separator + uuid.substring(0, 2) + File.separator + uuid.substring(2, 4)));
            String filePath = System.getProperty("user.home") + File.separator + "translate" + File.separator + uuid.substring(0, 2) + File.separator + uuid.substring(2, 4) +  File.separator + uuid + ".xlsx";

            FileOutputStream outputStream = new FileOutputStream(filePath);
            System.out.println("UUID: " + uuid);
            workbook.write(outputStream);
            return filePath;
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                workbook.close();
            } catch (IOException ignored) {
            }
        }
    }
}
