package com.cherniak.xlsxfile.xlsfilecreate;

import java.io.IOException;
import java.net.URI;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.LinkedList;
import java.util.List;
import java.util.Objects;
import lombok.RequiredArgsConstructor;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.security.util.InMemoryResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.xml.sax.SAXException;

@RestController
@RequiredArgsConstructor
@RequestMapping
public class FileController {

  private final ExcelService excelService;
  private final JExcelService jExcelService;
  private final static int ROWS = 1000;

  @GetMapping(path = "/download", produces = MediaType.MULTIPART_FORM_DATA_VALUE)
  public ResponseEntity<Resource> getAllByDate() throws IOException {
    long startTime = System.currentTimeMillis();
    Resource file = getXlsFile();
    ResponseEntity<Resource> response = ResponseEntity.ok()
        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename*=UTF-8''" +
            URI.create(file.getDescription()).toASCIIString())
        .body(file);
    long endTime = System.currentTimeMillis();
    System.out.println("Total execution time: " + (endTime - startTime) + "ms");
    return response;
  }

  @GetMapping(path = "/download2", produces = MediaType.MULTIPART_FORM_DATA_VALUE)
  public ResponseEntity<Resource> getAllByDate2()
      throws IOException, OpenXML4JException, SAXException, InterruptedException {
    long startTime = System.currentTimeMillis();
    Resource file = getXlsFile2();
    System.out.println("file.getDescription() " + 500 + file.getFilename());
    ResponseEntity<Resource> response = ResponseEntity.ok()
        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename*=UTF-8''" +
            URI.create(500 + file.getFilename()).toASCIIString())
        .body(file);
    long endTime = System.currentTimeMillis();
    System.out.println("Total execution time: " + (endTime - startTime) + "ms");
    return response;
//    return ResponseEntity.status(HttpStatus.NOT_FOUND)
//        .body(new ByteArrayResource("Файл не найден".getBytes(StandardCharsets.UTF_8)));
  }
//
//  @GetMapping(path = "/download3", produces = MediaType.MULTIPART_FORM_DATA_VALUE)
//  public ResponseEntity<Resource> getAllByDate3() throws IOException, WriteException {
//    long startTime = System.currentTimeMillis();
//    Resource file = getXlsFile3();
//    ResponseEntity<Resource> response =  ResponseEntity.ok()
//        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename*=UTF-8''" +
//            URI.create(file.getDescription()).toASCIIString())
//        .body(file);
//    long endTime = System.currentTimeMillis();
//    System.out.println("Total execution time: " + (endTime-startTime) + "ms");
//    return response;
//  }


  private Resource getXlsFile() throws IOException {
    List<Accident> allAccidents = new LinkedList<>();
    for (int i = 0; i < ROWS; i++) {
      Accident accident = new Accident();
      accident.setAccidentDate(LocalDate.now());
      accident.setBodyNum("setBodyNum" + i);
      accident.setDocNum("setDocNum" + i);
      accident.setCreateDate(LocalDate.now());
      accident.setDriver("не_предоставляется");
      accident.setBodyNum("setBodyNum" + i);
      accident.setExtId("setExtId" + i);
      accident.setIss("setIss" + i);
      accident.setDocType("setDocType" + i);
      accident.setPolicyNum("setPolicyNum" + i);
      accident.setRegNum("setRegNum" + i);
      accident.setSkName("setSkName" + i);
      accident.setVin("setVin" + i);

      allAccidents.add(accident);
    }
    return new InMemoryResource(
        excelService.generateFile(allAccidents, LocalDate.now()).toByteArray(),
        "myfile_OUT.xlsx"
    );
  }


  private Resource getXlsFile2() throws IOException, OpenXML4JException, SAXException {
    List<Accident> allAccidents = new LinkedList<>();
    for (int i = 0; i < ROWS; i++) {
      Accident accident = new Accident();
      accident.setAccidentDate(LocalDate.now());
      accident.setBodyNum("setBodyNum" + i);
      accident.setDocNum("setDocNum" + i);
      accident.setCreateDate(LocalDate.now());
      accident.setDriver("не_предоставляется");
      accident.setBodyNum("setBodyNum" + i);
      accident.setExtId("setExtId" + i);
      accident.setIss("setIss" + i);
      accident.setDocType("setDocType" + i);
      accident.setPolicyNum("setPolicyNum" + i);
      accident.setRegNum("setRegNum" + i);
      accident.setSkName("setSkName" + i);
      accident.setVin("setVin" + i);

      allAccidents.add(accident);
    }
//    return new InMemoryResource(
//        excelService.generateOldFile(allAccidents, LocalDate.now()).toByteArray(),
//        "myfile_OUT2.xlsx"
//    );
     return new FileSystemResource(excelService.generateOldFile(allAccidents, LocalDate.now()));
  }
//
//  private Resource getXlsFile3() throws IOException, WriteException {
//    List<Accident> allAccidents = new LinkedList<>();
//    for (int i = 0; i < ROWS; i++) {
//      Accident accident = new Accident();
//      accident.setAccidentDate(LocalDate.now());
//      accident.setBodyNum("setBodyNum" + i);
//      accident.setDocNum("setDocNum" + i);
//      accident.setCreateDate(LocalDate.now());
//      accident.setDriver("не_предоставляется");
//      accident.setBodyNum("setBodyNum" + i);
//      accident.setExtId("setExtId" + i);
//      accident.setIss("setIss" + i);
//      accident.setDocType("setDocType" + i);
//      accident.setPolicyNum("setPolicyNum" + i);
//      accident.setRegNum("setRegNum" + i);
//      accident.setSkName("setSkName" + i);
//      accident.setVin("setVin"+i);
//
//      allAccidents.add(accident);
//    }
//    return new InMemoryResource(
//        jExcelService.generateFile(allAccidents, LocalDate.now()).toByteArray(),
//        "myfile_OUT3.xlsx"
//    );
//  }

}
