package net.masaodev.text.extraction;

import java.io.File;
import java.io.IOException;
import java.util.Collection;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Application {

  private static final Logger logger = LoggerFactory.getLogger(Application.class);

  public static void main(String[] args) throws IOException {
    String targetDir = args[0];
    String outputDir = args[1];
    logger.info("対象ディレクトリルート:{}", targetDir);
    logger.info("出力先:{}", outputDir);

    Collection<File> searchExcelFiles = ExcelUtil.searchExcelFiles(targetDir);

    for (File targetFile : searchExcelFiles) {
      String str = ExcelUtil.extractStringFromExcelBook(targetFile);

      String destFilePath = targetFile.getAbsolutePath().replace(targetDir, outputDir);
      String parentPath = FilenameUtils.getFullPath(destFilePath);
      File parent = new File(parentPath);
      parent.mkdirs();

      FileUtils.writeStringToFile(new File(parent, targetFile.getName() + ".txt"), str, "utf-8");
    }
  }
}
