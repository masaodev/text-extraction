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
    String outputDir = null;
    if (args.length >= 2) {
      outputDir = args[1];
    }

    logger.info("対象ディレクトリルート:{}", targetDir);
    logger.info("出力先:{}", outputDir);

    File file = new File(targetDir);
    if (file.isFile()) {
      // 単一ファイル時

      String str = ExcelUtil.extractStringFromExcelBook(file);
      FileUtils.writeStringToFile(new File(outputDir, file.getName() + ".txt"), str, "utf-8");

    } else {
      // ディレクトリ一括時

      Collection<File> searchExcelFiles = ExcelUtil.searchExcelFiles(targetDir);
      Runtime r = Runtime.getRuntime();
      for (File targetFile : searchExcelFiles) {
        String str = ExcelUtil.extractStringFromExcelBook(targetFile);
        String destFilePath = targetFile.getAbsolutePath().replace(targetDir, outputDir);
        String parentPath = FilenameUtils.getFullPath(destFilePath);
        File parent = new File(parentPath);
        parent.mkdirs();

        FileUtils.writeStringToFile(new File(parent, targetFile.getName() + ".txt"), str, "utf-8");
        r.gc();
      }
    }
  }
}
