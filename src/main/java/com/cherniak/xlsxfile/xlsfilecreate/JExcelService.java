package com.cherniak.xlsxfile.xlsfilecreate;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.springframework.stereotype.Service;

@Service
public class JExcelService {

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

  public ByteArrayOutputStream generateFile(List<Accident> accidentList, LocalDate createDate)
      throws IOException, WriteException {

    ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
    //WritableWorkbook workbook = Workbook.createWorkbook(byteOut);
    File file = new File("C:\\Users\\barse\\IdeaProjects\\xls-file-create\\pomesJ.xls");
    file.getParentFile().mkdirs();

    WritableWorkbook workbook = Workbook.createWorkbook(file);
    WritableSheet sheet = workbook.createSheet("ЗАРЕГИСТРИРОВАННЫЕ СЛУЧАИ", 0);

    WritableCellFormat normalFormat = new WritableCellFormat();
    normalFormat.setWrap(true);

    Label firstLabel0 = new Label(0, 0, "Дата формирования реестра: ", normalFormat);
    sheet.addCell(firstLabel0);
    Label firstLabel1 = new Label(1, 0, Optional.of(createDate)
        .map(created -> new Date())
        .map(simpleDateFormat::format)
        .orElse(""), normalFormat);
    sheet.addCell(firstLabel1);

    WritableCellFormat headerFormat = new WritableCellFormat();
    WritableFont font
        = new WritableFont(WritableFont.ARIAL, 11, WritableFont.BOLD);
    headerFormat.setFont(font);
    headerFormat.setBackground(Colour.BLUE_GREY);
    headerFormat.setWrap(true);

    for (int i = 0; i < columnNames.size(); i++) {
      Label label = new Label(i, 1, columnNames.get(i), headerFormat);
      sheet.setColumnView(0, 60);
      sheet.addCell(label);
    }

    for (int i = 0; i < accidentList.size(); i++) {
      int cell = 0;
      Accident a = accidentList.get(i);
      List<Label> list = new ArrayList<>();
      list.add(new Label(cell++, i + 2, (i + 1) + "", normalFormat));
      list.add(new Label(cell++, i + 2, a.getPolicyNum() + "", normalFormat));
      list.add(new Label(cell++, i + 2, a.getSkName() + "", normalFormat));
      list.add(new Label(cell++, i + 2, Optional.of(createDate)
          .map(created -> new Date())
          .map(simpleDateFormat::format)
          .orElse(""), normalFormat));
      list.add(new Label(cell++, i + 2, a.getDriver(), normalFormat));
      list.add(new Label(cell++, i + 2, a.getDocType() + "", normalFormat));
      list.add(new Label(cell++, i + 2, a.getDocNum() + "", normalFormat));
      list.add(new Label(cell++, i + 2, a.getVin() + "", normalFormat));
      list.add(new Label(cell++, i + 2, a.getRegNum() + "", normalFormat));
      list.add(new Label(cell++, i + 2, a.getBodyNum() + "", normalFormat));
      list.add(new Label(cell++, i + 2, a.getIss() + "", normalFormat));
      list.add(new Label(cell++, i + 2, LocalDate.now().getDayOfMonth() + "", normalFormat));
      list.add(new Label(cell, i + 2, "r.getErrorDescription()", normalFormat));

      for (int j = 0; j < list.size(); j++) {
        sheet.addCell(list.get(j));
      }
    }
    workbook.write();
    workbook.close();

    return byteOut;

  }

}
