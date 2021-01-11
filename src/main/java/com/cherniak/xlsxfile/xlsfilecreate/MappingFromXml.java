package com.cherniak.xlsxfile.xlsfilecreate;


import java.util.Map;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

public class MappingFromXml implements SheetContentsHandler {

  private final Map<String, String> map;
  private int lineNumber = 0;
  private String key;

  /**
   * Destination for data
   */
  public MappingFromXml(Map<String, String> map) {
    this.map = map;
  }

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
            map.put(key, formattedValue);
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
}
