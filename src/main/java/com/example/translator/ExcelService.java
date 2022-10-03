//package com.example.translator;
//
//import lombok.ToString;
//import org.springframework.beans.factory.annotation.Autowired;
//import org.springframework.stereotype.Service;
//import org.springframework.web.context.annotation.ApplicationScope;
//import org.springframework.web.context.annotation.SessionScope;
//
//import java.io.IOException;
//import java.io.InputStream;
//import java.nio.file.Files;
//import java.nio.file.Path;
//import java.util.Map;
//
//@Service
//@SessionScope
//@ToString
//public class ExcelService {
//    public int numberOfRows = 0;
//    public int numberOfReadyRows = 0;
//
//    @Autowired
//    private ExcelWriter excelWriter;
//    @Autowired
//    private ExcelReader excelReader;
//
//    public byte[] getTranslatedFile(InputStream inputStream, String uuid, int columnIndex) throws IOException {
//        excelWriter.init();
//        return Files.readAllBytes(Path.of(excelReader.readTest(this, inputStream, uuid, columnIndex)));
//    }
//}
