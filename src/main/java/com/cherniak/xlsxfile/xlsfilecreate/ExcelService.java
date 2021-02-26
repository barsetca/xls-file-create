package com.cherniak.xlsxfile.xlsfilecreate;

import com.incesoft.tools.excel.xlsx.Sheet.SheetRowReader;
import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook;
import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import javax.xml.parsers.ParserConfigurationException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.ltgfmt.TestCase;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;
import org.springframework.stereotype.Service;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

@Service
@Slf4j
public class ExcelService {

  private final List<String> columnNames = List.of(
      "№",
      "Серия и номер полиса",
      "Прямой страховщик",
      "Дата ДТП",
      "ФИО",
      "Тип документа",
      "Серия и номер документа",
      "VIN",
      "Гос. номер",
      "Номер кузова",
      "ИСС",
      "Время обработки запроса",
      "Описание ошибки");

  private final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd.MM.yyyy");

  private final int rowAccessWindowSize = 200;
  private final boolean compressTmpFiles = false;

  //  /**
//   * parse file
//   *
//   * @param inputStream is
//   * @return parsed data as dictionary
//   * @throws IOException
//   */

  //inmemory parser
  public Map<String, String> parseFile(InputStream inputStream) throws IOException {
    Sheet sheet = new XSSFWorkbook(inputStream).getSheetAt(0);
    Map<String, String> result = new HashMap<>();

    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      Row r = sheet.getRow(i);
      if (r.getCell(2) == null || r.getCell(2) == null) {
        break;
      }
      result.put(r.getCell(2).getStringCellValue(), r.getCell(2).getStringCellValue());
    }

    int size = result.keySet().size();
    log.info("Found {} record(s) in file", size);
    if (size == 0) {
      throw new IllegalArgumentException(size + " record(s) in file");
    }
    return result;

  }

  /**
   * generate file
   *
   * @param accidentList accident list
   * @param createDate   file creation date
   * @return file as byte array
   * @throws IOException
   */
  //inmemory generate xlsx (for small files)
  //http://poi.apache.org/components/spreadsheet/how-to.html#user_api
  public ByteArrayOutputStream generateInmemoryFile(List<Accident> accidentList, Instant createDate)
      throws IOException {

    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet sheet = workbook.createSheet("ЗАРЕГИСТРИРОВАННЫЕ СЛУЧАИ");
    sheet.setDefaultColumnWidth(30);

    XSSFTable table = sheet.createTable();
    CTTable cttable = table.getCTTable();

    cttable.setDisplayName("Table");
    cttable.setId(1);
    cttable.setName("Registred Accidents");
    cttable.setRef("A2:M65536");
    cttable.setTotalsRowShown(false);

    CTTableStyleInfo styleInfo = cttable.addNewTableStyleInfo();
    styleInfo.setName("TableStyleMedium2");
    styleInfo.setShowColumnStripes(false);
    styleInfo.setShowRowStripes(true);

    CTTableColumns columns = cttable.addNewTableColumns();
    int columnsCount = columnNames.size();
    columns.setCount(columnsCount);

    sheet.createRow(0)
        .createCell(0)
        .setCellValue(Optional.of(createDate)
            .map(created -> new Date(created.toEpochMilli()))
            .map(simpleDateFormat::format)
            .map(s -> "Дата формирования реестра: " + s)
            .orElse(""));

    for (int i = 1; i <= columnsCount; i++) {
      CTTableColumn column = columns.addNewTableColumn();
      column.setId(i);
      column.setName("Col" + i);
    }

    XSSFRow headerRow = sheet.createRow(1);
    for (int i = 0; i < columnsCount; i++) {
      headerRow.createCell(i).setCellValue(columnNames.get(i));
    }

    SXSSFWorkbook workbookS = new SXSSFWorkbook(workbook, 200, false);
    SXSSFSheet sheetS = workbookS.getSheetAt(0);

    for (int i = 0; i < accidentList.size(); i++) {
      SXSSFRow tableRow = sheetS.createRow(i + sheet.getLastRowNum() + 1);
      int cell = 0;
      Accident a = accidentList.get(i);
      //Report r = a.getReport();

      tableRow.createCell(cell++).setCellValue(i + 1);
      tableRow.createCell(cell++).setCellValue(a.getPolicyNum());
      tableRow.createCell(cell++).setCellValue(a.getSkName());
      tableRow.createCell(cell++).setCellValue(Optional.ofNullable(a.getAccidentDate())
          .map(created -> new Date())
          .map(simpleDateFormat::format).orElse(null));
      tableRow.createCell(cell++).setCellValue(a.getDriver());
      tableRow.createCell(cell++).setCellValue(a.getDocType());
      tableRow.createCell(cell++).setCellValue(a.getDocNum());
      tableRow.createCell(cell++).setCellValue(a.getVin());
      tableRow.createCell(cell++).setCellValue(a.getRegNum());
      tableRow.createCell(cell++).setCellValue(a.getBodyNum());
      tableRow.createCell(cell++).setCellValue(a.getIss());
      tableRow.createCell(cell++);
      tableRow.createCell(cell).setCellValue("r.getErrorDescription()");
    }
    ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
    workbookS.write(byteOut);
    workbookS.dispose();
    workbookS.close();
    return byteOut;
  }

  //stream generate xlsx with pure (only) SXSSFWorkbook (set style for each cell)
  //http://poi.apache.org/components/spreadsheet/how-to.html#sxssf
  public ByteArrayOutputStream generateFile(List<Accident> accidentList, LocalDate createDate)
      throws IOException {
    SXSSFWorkbook workbook = new SXSSFWorkbook(200);
    Map<String, CellStyle> styles = createStyles(workbook);
    SXSSFSheet sheet = workbook.createSheet("ЗАРЕГИСТРИРОВАННЫЕ СЛУЧАИ");
    sheet.setDefaultColumnWidth(30);

    sheet.createRow(0).createCell(0)
        .setCellValue(Optional.of(createDate)
            .map(created -> new Date())
            .map(simpleDateFormat::format)
            .map(s -> "Дата формирования реестра: " + s).orElse(""));

    SXSSFRow headerRow = sheet.createRow(1);
    for (int i = 0; i < columnNames.size(); i++) {
      Cell cell = headerRow.createCell(i);
      cell.setCellStyle(styles.get("header"));
      cell.setCellValue(columnNames.get(i));
    }

    for (int i = 0; i < accidentList.size(); i++) {
      SXSSFRow tableRow = sheet.createRow(i + 2);
      int cell = 0;
      Accident a = accidentList.get(i);
//      Report r = a.getReport();
      tableRow.createCell(cell++).setCellValue(i + 1);
      tableRow.createCell(cell++).setCellValue(a.getPolicyNum());
      tableRow.createCell(cell++).setCellValue(a.getSkName());
      tableRow.createCell(cell++).setCellValue(Optional.ofNullable(a.getAccidentDate())
          .map(created -> new Date())
          .map(simpleDateFormat::format).orElse(null));
      tableRow.createCell(cell++).setCellValue(a.getDriver());
      tableRow.createCell(cell++).setCellValue(a.getDocType());
      tableRow.createCell(cell++).setCellValue(a.getDocNum());
      tableRow.createCell(cell++).setCellValue(a.getVin());
      tableRow.createCell(cell++).setCellValue(a.getRegNum());
      tableRow.createCell(cell++).setCellValue(a.getBodyNum());
      tableRow.createCell(cell++).setCellValue(a.getIss());
      tableRow.createCell(cell++);
      tableRow.createCell(cell).setCellValue("r.getErrorDescription()");
      if (i % 2 == 0) {
        for (int j = 0; j < tableRow.getLastCellNum(); j++) {
          tableRow.getCell(j).setCellStyle(styles.get("odd"));
        }
      }
    }

    File file = new File("C:\\Users\\barse\\IdeaProjects\\xls-file-create\\pomes.xlsx");
    file.getParentFile().mkdirs();

    FileOutputStream outFile = new FileOutputStream(file);
    workbook.write(outFile);
    System.out.println("Created file: " + file.getAbsolutePath());
    outFile.close();

    ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
//    workbook.write(byteOut);
    workbook.dispose();
    workbook.close();
    return byteOut;
  }

  //stream generate xlsx using XSSFWorkbook how template for table size in SXSSFWorkbook
  public InputStream generateOldFile(List<Accident> accidentList, LocalDate createDate, String id)
      throws IOException, OpenXML4JException, SAXException {
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet sheet = workbook.createSheet("ЗАРЕГИСТРИРОВАННЫЕ СЛУЧАИ");
    sheet.setDefaultColumnWidth(30);

    XSSFTable table = sheet.createTable();
    CTTable cttable = table.getCTTable();

    cttable.setDisplayName("Table");
    cttable.setId(1);
    cttable.setName("Registred Accidents");
    cttable.setRef("A2:M65536");
    cttable.setTotalsRowShown(false);

    CTTableStyleInfo styleInfo = cttable.addNewTableStyleInfo();
    styleInfo.setName("TableStyleMedium2");
    styleInfo.setShowColumnStripes(false);
    styleInfo.setShowRowStripes(true);

    CTTableColumns columns = cttable.addNewTableColumns();
    columns.setCount(columnNames.size());

    sheet.createRow(0).createCell(0)
        .setCellValue(Optional.of(createDate)
            .map(created -> new Date())
            .map(simpleDateFormat::format)
            .map(s -> "Дата формирования реестра: " + s).orElse(""));

    for (int i = 1; i <= columnNames.size(); i++) {
      CTTableColumn column = columns.addNewTableColumn();
      column.setId(i);
      column.setName("Col" + i);
    }

    XSSFRow headerRow = sheet.createRow(1);
    for (int i = 0; i < columnNames.size(); i++) {
      headerRow.createCell(i).setCellValue(columnNames.get(i));
    }
    SXSSFWorkbook workbookS = new SXSSFWorkbook(workbook, rowAccessWindowSize, compressTmpFiles);
    SXSSFSheet sheetS = workbookS.getSheetAt(0);

    for (int i = 0; i < accidentList.size(); i++) {
      SXSSFRow tableRow = sheetS.createRow(i + sheet.getLastRowNum() + 1);
      int cell = 0;
      Accident a = accidentList.get(i);

      tableRow.createCell(cell++).setCellValue(i + 1);
      tableRow.createCell(cell++).setCellValue(a.getPolicyNum());
      tableRow.createCell(cell++).setCellValue(a.getSkName());
      tableRow.createCell(cell++).setCellValue(Optional.ofNullable(a.getAccidentDate())
          .map(created -> new Date())
          .map(simpleDateFormat::format).orElse(null));
      tableRow.createCell(cell++).setCellValue(a.getDriver());
      tableRow.createCell(cell++).setCellValue(a.getDocType());
      tableRow.createCell(cell++).setCellValue(a.getDocNum());
      tableRow.createCell(cell++).setCellValue(a.getVin());
      tableRow.createCell(cell++).setCellValue(a.getRegNum());
      tableRow.createCell(cell++).setCellValue(a.getBodyNum());
      tableRow.createCell(cell++).setCellValue(a.getIss());

      tableRow.createCell(cell++);

      tableRow.createCell(cell).setCellValue("r.getErrorDescription()");
    }

    Path filePath = Files.createTempFile(Paths.get(""), "temp-download-", ".xlsx");
//File file = filePath.toFile();

    //Path filePath = Paths.get(id + "_OUT.xlsx");
//    File file = File.createTempFile("test", ".xlsx");
//    file.deleteOnExit();
    //Files.deleteIfExists(filePath);
    //new File("myfile_OUT.xlsx");
    //file.getParentFile().mkdirs();

    try (var outFile = Files.newOutputStream(filePath)) {
      //FileOutputStream outFile = new FileOutputStream(file);
      workbookS.write(outFile);
      System.out.println("Created file: " + filePath.toAbsolutePath().toString());
      //outFile.close();
    }
    //ByteArrayOutputStream byteOut = new ByteArrayOutputStream();

    //FileSystemResource fsr = new FileSystemResource("C:\\Users\\barse\\IdeaProjects\\xls-file-create\\myfile_OUT2.xlsx");
//    workbookS.write(byteOut);
    //System.out.println("ByteArrayOutputStream записан " + byteOut.toString(Charset.defaultCharset()));
    //workbookS.close();
    workbookS.close();
    if (!workbookS.dispose()) {
      System.out.println("Временные файлы workbookS НЕ удалены!");
    } else {
      System.out.println("Временные файлы workbookS удалены!");
    }

//    System.out.println("Временные файлы workbookS удалены? " + workbookS.dispose());
    //
    System.out.println("Main before return Файл существует? " + Files.exists(filePath));
    if (Files.exists(filePath)) {
      System.out.println(filePath.getFileName().toString() + " length " + Files.size(filePath));
    }

    //checking stream parsers
/*
 FileInputStream inputStream = new FileInputStream(file);
    long timeStart = System.currentTimeMillis();
 */

    //BufferedInputStream inputStream = new BufferedInputStream(new XlsxInputStreamImpl(file));
    //InputStream is = new XlsxInputStreamImpl(file);
    //inputStream.close();
    //is.close();

    //Map<String, String> map = parseFileSjxlsx(inputStream);
/*
   Map<String, String> map = parseExcelWithoutMapping(inputStream);
    long timeEnd = System.currentTimeMillis();
    System.out.println("creating map for " + (timeEnd - timeStart) + " ms");
    System.out.println("map size " + map.size());
    inputStream.close();
 */
    //map.forEach((k,v) -> System.out.println(k + " - " + v));

//
//    InputStream inputStream = new InputStreamDelegator(new BufferedInputStream(new XlsxInputStreamImpl(file))
//        , file
//    );

    //    return new InputStream() {
//      final BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
//
//      @Override
//      public int read() throws IOException {
//        return bis.read();
//      }
//
//      @Override
//      public void close() throws IOException {
//        bis.close();
//       boolean isDelete = false;
//           //Files.deleteIfExists(file.toPath());
//        System.out.println(
//            "Files.deleteIfExists " + isDelete + "\n" + file
//                .getAbsolutePath());
//      }
//    };
    return new InputStreamDelegator(new BufferedInputStream(Files.newInputStream(filePath)),
        () -> deleteFileSilently(filePath), filePath.toFile());

  }

  //    return new InputStreamDelegator(new BufferedInputStream(new XlsxInputStreamImpl(file)), () -> deleteTempFile(file));
//    //return filePath;
//  }
  private void deleteFileSilently(Path filePath) {
    try {
//      System.out.println("Файл удален? " + Files.deleteIfExists(Paths.get("pomes.xlsx")));

      System.out.println("deleteTempFile Файл существует? " + Files.exists(filePath));
      if (Files.exists(filePath)) {
        System.out
            .println(filePath.toAbsolutePath().toString() + " length " + Files.size(filePath));
      }
      //boolean delete = Files.deleteIfExists(file.toPath());
      Files.delete(filePath);
      System.out.println("Временный файл XLSX удален!" + "\n" + filePath.toAbsolutePath().toString());
      System.out.println("deleteTempFile Файл существует? " + Files.exists(filePath));

    } catch (IOException e) {
      System.out.println("Попали сюда");
      System.out.println("Временный файл XLSX НЕ удален!");
      System.out.println("deleteTempFile Файл существует? " + Files.exists(filePath));
      e.printStackTrace();
    }
  }

  //my deprecated inmemory generators
//  public void parseFileNew(InputStream inputStream)
//      throws IOException, OpenXML4JException, SAXException {
////
//    OPCPackage container = OPCPackage.open(inputStream);
//    ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(container);
//
//    XSSFReader xssfReader = new XSSFReader(container);
//    StylesTable styles = xssfReader.getStylesTable();
//    XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
//    while (iter.hasNext()) {
//      InputStream stream = iter.next();
//      processSheet(styles, strings, stream);
//      stream.close();
//    }

//    XSSFWorkbook wb_template = new XSSFWorkbook(inputStream);
//    inputStream.close();
//    Map<String, String> result = new HashMap<>();
//    SXSSFWorkbook wb = new SXSSFWorkbook(wb_template, 100, true);
//    //Sheet sheet = wb.getXSSFWorkbook().getSheetAt(0);
//    SXSSFSheet sh = wb.getSheetAt(0);
//
//    File file = new File("C:\\Users\\barse\\IdeaProjects\\xls-file-create\\pomesTemp.xlsx");
//    file.getParentFile().mkdirs();
//
//    FileOutputStream outFile = new FileOutputStream(file);
//    wb.write(outFile);
//    System.out.println("Created file: " + file.getAbsolutePath());
//    outFile.close();
//    wb.dispose();
//    wb.close();
//    XSSFWorkbook workbook = new XSSFWorkbook(file);

//    Iterator<Sheet> iterator = wb.sheetIterator();
//    while (iterator.hasNext()){
//      Sheet sheet = iterator.next();
//          for (int i = 1; i <= sheet.getLastRowNum(); i++) {
//      Row r = sheet.getRow(i);
//      if (r == null){
//        System.out.println("row = null");
//      }
//
//      if (r.getCell(2) == null || r.getCell(2) == null) {
//        break;
//      }
//      result.put(r.getCell(2).getStringCellValue(), r.getCell(2).getStringCellValue());
//    }
//    }

//    Sheet sheet = wb.getXSSFWorkbook().getSheetAt(0);
//    Iterator<Row> rowIterator = sheet.rowIterator();
//    while (rowIterator.hasNext()) {
//      Row row = rowIterator.next();
//      if (row == null) {
//        System.out.println("row = null");
//      }
//
//      if (row.getCell(2) == null || row.getCell(2) == null) {
//        continue;
//      }
//      result.put(row.getCell(2).getStringCellValue(), row.getCell(2).getStringCellValue());
//    }

//    Sheet sheet = workbook.getSheetAt(0);
//
//    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
//      Row r = sheet.getRow(i);
//      if (r == null) {
//        System.out.println("row = null");
//      }
//
//      if (r.getCell(2) == null || r.getCell(2) == null) {
//        break;
//      }
//      result.put(r.getCell(2).getStringCellValue(), r.getCell(2).getStringCellValue());
//    }
//
//    int size = result.keySet().size();
//    log.info("Found {} record(s) in file", size);
//    if (size == 0) {
//      throw new IllegalArgumentException(size + " record(s) in file");
//    }
//    wb.dispose();

  // return null;
//  }


  //stream xlsx parser Implementation with MappingFromXml.class
  //https://stackanswers.net/questions/error-while-reading-large-excel-files-xlsx-via-apache-poi
  public Map<String, String> parseExcel(InputStream inputStream) throws IOException {
    Map<String, String> map = new HashMap<>();
    try {
      OPCPackage xlsxPackage = OPCPackage.open(inputStream);
      ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(xlsxPackage);
      XSSFReader xssfReader = new XSSFReader(xlsxPackage);
      StylesTable styles = xssfReader.getStylesTable();
      XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
      int index = 0;
      while (iter.hasNext()) {
        try (InputStream stream = iter.next()) {
          processSheet(styles, strings, new MappingFromXml(map), stream);
        }
        ++index;
      }
      System.out.println(index);
    } catch (OpenXML4JException | SAXException e) {
      e.printStackTrace();
    }

    return map;
  }

  //Implementation with MappingFromXml.class
  //https://stackanswers.net/questions/error-while-reading-large-excel-files-xlsx-via-apache-poi
  private static void processSheet(
      StylesTable styles,
      ReadOnlySharedStringsTable strings,
      MappingFromXml mappingFromXml,
      InputStream sheetInputStream) throws IOException, SAXException {
    DataFormatter formatter = new DataFormatter();
    InputSource sheetSource = new InputSource(sheetInputStream);
    try {
      XMLReader sheetParser = SAXHelper.newXMLReader();
      ContentHandler handler = new XSSFSheetXMLHandler(
          styles, null, strings, mappingFromXml, formatter, false);
      sheetParser.setContentHandler(handler);
      sheetParser.parse(sheetSource);
    } catch (ParserConfigurationException e) {
      throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
    }
  }

  //stream xlsx parser Implementation with anonymous SheetContentsHandler
  public Map<String, String> parseExcelWithoutMapping(InputStream inputStream) throws IOException {
    Map<String, String> map = new HashMap<>();
    try {
      OPCPackage xlsxPackage = OPCPackage.open(inputStream);
      ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(xlsxPackage);
      XSSFReader xssfReader = new XSSFReader(xlsxPackage);
      StylesTable styles = xssfReader.getStylesTable();
      XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
      //int index = 0;
      while (iter.hasNext()) {
        try (InputStream stream = iter.next()) {
//          String[] key = new String[1];
//          int[] lineNumber = new int[1];
          processSheetWithoutMapping(styles, strings, new SheetContentsHandler() {
            int lineNumber;
            String key;

            @Override
            public void startRow(int i) {
              lineNumber = i;
            }

            @Override
            public void endRow(int i) {
            }

            @Override
            public void cell(String cellReference, String formattedValue, XSSFComment comment) {
              int columnIndex = (new CellReference(cellReference)).getCol();
              if (lineNumber > 0) {
                switch (columnIndex) {
                  case 0: {
                    if (formattedValue != null && !formattedValue.isEmpty()) {
                      key = formattedValue;
                    }
                  }
                  break;
                  case 1: {
                    if (formattedValue != null && !formattedValue.isEmpty()) {
                      map.put(formattedValue, key);
                    }
                  }
                  break;
                  default:
                }
              }
            }

            @Override
            public void headerFooter(String s, boolean b, String s1) {

            }
          }, stream);
        }
        //++index;
      }
    } catch (OpenXML4JException | SAXException e) {
      e.printStackTrace();
    }
    return map;
  }

  //with anonymous SheetContentsHandler
  private static void processSheetWithoutMapping(
      StylesTable styles,
      ReadOnlySharedStringsTable strings,
      SheetContentsHandler sheetContentsHandler,
      InputStream sheetInputStream) throws IOException, SAXException {
    DataFormatter formatter = new DataFormatter();
    InputSource sheetSource = new InputSource(sheetInputStream);
    try {
      XMLReader sheetParser = SAXHelper.newXMLReader();
      ContentHandler handler = new XSSFSheetXMLHandler(
          styles, null, strings, sheetContentsHandler, formatter, false);
      sheetParser.setContentHandler(handler);
      sheetParser.parse(sheetSource);
    } catch (ParserConfigurationException e) {
      throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
    }
  }

  //stream parser Implementation parser xlsx with my fixing sjxlsx lib
  //https://github.com/davidpelfree/sjxlsx
  public Map<String, String> parseFileSjxlsx(InputStream inputStream)
      throws IOException {
    Map<String, String> map = new HashMap<>();
    Path tempPath = Files.createTempFile(Paths.get(""), "sjxlsx-", ".xlsx");
    Files.copy(inputStream, tempPath, StandardCopyOption.REPLACE_EXISTING);
    inputStream.close();
    SimpleXLSXWorkbook simpleWorkbook = new SimpleXLSXWorkbook(tempPath.toFile());

    com.incesoft.tools.excel.xlsx.Sheet sheetToRead = simpleWorkbook.getSheet(0);
    SheetRowReader reader = sheetToRead.newReader();
    com.incesoft.tools.excel.xlsx.Cell[] row;
    int rows = sheetToRead.getRowCount();
    for (int i = 0; i < rows; i++) {
      row = reader.readRow();
      if (i != 0) {
        map.put( row[1].getValue(), row[0].getValue());
      }
    }
    simpleWorkbook.close();
    Files.delete(tempPath);
    return map;
  }


  //Create map of cells style for pure SXSSFWorkbook
  private Map<String, CellStyle> createStyles(Workbook wb) {
    Map<String, CellStyle> styles = new HashMap<>();
    CellStyle style;

    Font titleFont = wb.createFont();
    titleFont.setFontHeightInPoints((short) 11);
    titleFont.setFontName("Calibri");
    titleFont.setBold(true);
    titleFont.setColor(IndexedColors.WHITE.index);
    style = wb.createCellStyle();
    style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    style.setFont(titleFont);
    style.setWrapText(false);
    style.setBorderBottom(BorderStyle.THIN);
    style.setBottomBorderColor(IndexedColors.BLUE1.getIndex());
    styles.put("header", style);

    Font rowFont = wb.createFont();
    rowFont.setColor(IndexedColors.BLACK.index);
    style = wb.createCellStyle();
    style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    style.setFont(rowFont);
    style.setWrapText(false);
    style.setBorderBottom(BorderStyle.THIN);
    style.setBorderTop(BorderStyle.THIN);
    style.setBottomBorderColor(IndexedColors.BLUE1.getIndex());
    style.setTopBorderColor(IndexedColors.BLUE1.getIndex());
    styles.put("odd", style);
    return styles;
  }

  //For check upload stream generate xlsx using XSSFWorkbook how template for table size in SXSSFWorkbook
  public void generateOldFileUpload()
      throws IOException {

    //checking stream parsers
    File file = new File("C:\\Users\\barse\\IdeaProjects\\xls-file-create\\parse.xlsx");
 FileInputStream inputStream = new FileInputStream(file);
 long timeStart = System.currentTimeMillis();


    //BufferedInputStream inputStream = new BufferedInputStream(new XlsxInputStreamImpl(file));
    //InputStream is = new XlsxInputStreamImpl(file);
    //inputStream.close();
    //is.close();

    //Map<String, String> map = parseFileSjxlsx(inputStream);

  // Map<String, String> map = parseExcel(inputStream);
   Map<String, String> map = parseExcelWithoutMapping(inputStream);
    long timeEnd = System.currentTimeMillis();
    System.out.println("creating map for " + (timeEnd - timeStart) + " ms");
    System.out.println("map size " + map.size());
    map.forEach((k,v) -> System.out.println(k + " - " + v));
    inputStream.close();

  }
}
