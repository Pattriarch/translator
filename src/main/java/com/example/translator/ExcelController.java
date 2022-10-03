package com.example.translator;

import lombok.extern.slf4j.Slf4j;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.http.HttpSession;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

@Slf4j
@RestController
@CrossOrigin(origins = {"http://localhost:1234", "https://reqstd.ru"}, maxAge = 3600, allowCredentials = "true")
@RequestMapping("/api/translate")
public class ExcelController {
//    @Resource(name = "sessionScopedBean")
//    @Autowired
//    public ExcelReader excelReader;

    Map<String, ExcelReader> readers = new HashMap<>();

    @PostMapping(path = "/{userId}/upload")
    public ResponseEntity<Object> handlePost(@RequestParam(name = "file") MultipartFile file,
                                             @RequestParam(name = "columnIndex") String columnIndexStr,
                                             @PathVariable String userId) {
        readers.put(userId, new ExcelReader());
//        System.out.println("HANDLE: " + readers.get(userId));
        JSONArray array = new JSONArray();
        JSONObject objectName = new JSONObject();
        String fileName = file.getOriginalFilename();
        objectName.put("name", fileName);
        array.put(objectName);

        int columnIndex = Integer.parseInt(columnIndexStr) - 1;

        if (fileName != null && fileName.endsWith(".xlsx")) {
            log.debug("Загружается файл под названием {} для парсинга", fileName);

            try (InputStream excelIs = file.getInputStream()) {

                UUID uuid = UUID.randomUUID();
                String uuidAsString = uuid.toString();
                byte[] translatedFile = readers.get(userId).getTranslatedFile(excelIs, uuidAsString, columnIndex);

                log.debug("Файл " + fileName + " актуализирован. Отправка ответа.");

                JSONObject objectUUID = new JSONObject();
                objectUUID.put("uuid", uuidAsString);
                array.put(objectUUID);

                JSONObject object = new JSONObject();
                String encodedString = java.util.Base64.getEncoder().encodeToString(translatedFile);
                object.put("encodedString", encodedString);
                array.put(object);

                JSONObject objectPath = new JSONObject();
                objectPath.put("path", uuidAsString.substring(0, 2) + File.separator + uuidAsString.substring(2, 4) + File.separator + uuidAsString + ".xlsx");
                array.put(objectPath);

                JSONObject objectSize = new JSONObject();
                objectSize.put("size", (translatedFile.length / 1024));
                array.put(objectSize);

                return createResponseEntity(array.toList());
            } catch (IOException e) {
                log.error("Произошла ошибка во время парсинга файла с названием {}", fileName);
                return null;
            }
        }
        return null;
    }

    @GetMapping(path = "/{userId}/progress")
    public Integer[] getProgress(@PathVariable String userId, HttpSession httpSession) {
        return new Integer[] {
                this.readers.get(userId).getNumberOfRows(),
                this.readers.get(userId).getNumberOfReadyRows()
        };
    }

    @GetMapping(path = "/{userId}/resetProgress")
    public void resetProgress(@PathVariable String userId) {
        this.readers.get(userId).setNumberOfRows(0);
        this.readers.get(userId).setNumberOfReadyRows(0);
    }

    private ResponseEntity<Object> createResponseEntity(
            Object report
    ) {
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_JSON)
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + "SomeName")
                .body(report);
    }
}