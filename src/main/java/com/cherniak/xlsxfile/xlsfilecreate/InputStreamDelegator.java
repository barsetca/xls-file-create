package com.cherniak.xlsxfile.xlsfilecreate;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import lombok.experimental.Delegate;

public class InputStreamDelegator extends InputStream {

  @Delegate(excludes = Closeable.class)
  private final InputStream delegator;

  private final Runnable onClose;
 private File file;

  public InputStreamDelegator(InputStream delegator
   , Runnable onClose
     , File file
  ) {
    this.delegator = delegator;
    this.onClose = onClose;
   this.file = file;
  }

  @Override
  public void close() throws IOException {
    System.out.println("InputStreamDelegator close");
    System.out.println("InputStreamDelegator close Файл существует? " + file.exists());
    if (file.exists()) {
      System.out.println(file.getName() + " length " + file.length());
    }
    delegator.close();
    System.out.println("InputStreamDelegator close Файл существует? " + file.exists());
//    new Thread(onClose).start();
    onClose.run();
  }
}
