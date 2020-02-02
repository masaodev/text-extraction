package net.masaodev.text.extraction;

import java.io.File;
import java.io.IOException;
import java.util.Collection;
import org.apache.commons.io.FileUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Application {

  private static final Logger logger = LoggerFactory.getLogger(Application.class);

  public static void main(String[] args) throws IOException {
    logger.info("対象ディレクトリルート:{}", args[0]);
    logger.info("出力先:{}", args[1]);

    Collection<File> searchExcelFiles = ExcelUtil.searchExcelFiles(args[0]);

    int i = 0;
    for (File file : searchExcelFiles) {
      i++;
      String str = ExcelUtil.extractStringFromExcelBook(file);
      FileUtils.writeStringToFile(new File(args[1], i + ".txt"), str, "utf-8");
    }
  }
}
