package com.cherniak.xlsxfile.xlsfilecreate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public class XlsxInputStreamImpl extends FileInputStream {

  private final Path pathTempFile;

  public XlsxInputStreamImpl(File file) throws FileNotFoundException {
    super(file);
    pathTempFile = file.toPath();
  }

  @Override
  public void close() throws IOException {
    super.close();
    Files.deleteIfExists(pathTempFile);
  }
}
