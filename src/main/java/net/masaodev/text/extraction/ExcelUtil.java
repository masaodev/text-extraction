package net.masaodev.text.extraction;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;
import java.util.stream.Collectors;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.FileFilterUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFShapeGroup;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExcelUtil {

  private static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

  public static String extractStringFromExcelBook(File targetExcelBook) {

    // Excelファイルの読込み
    InputStream inputStream;
    Workbook workbook;

    try {
      inputStream = new FileInputStream(targetExcelBook);
      workbook = WorkbookFactory.create(inputStream);
      inputStream.close();
    } catch (EncryptedDocumentException | IOException e) {
      logger.error("Excelファイルが壊れている？:{}", targetExcelBook.getName());
      return "";
    }

    // 数式の参照更新設定
    workbook.setForceFormulaRecalculation(true);
    StringBuilder s = new StringBuilder();
    s.append(targetExcelBook.getAbsolutePath() + "\r\n");

    // シート枚数分ループ処理
    workbook.forEach(
        sheet -> {
          s.append("【Sheet】:" + sheet.getSheetName() + "\r\n");
          // 最終行までループ処理
          sheet.forEach(
              row -> {
                // 行内の最後のセルまでループ処理
                row.forEach(
                    cell -> {
                      String cellValue = getCellValue(cell);
                      if (cellValue != "") {
                        s.append(
                            convertCellPos(cell.getRowIndex(), cell.getColumnIndex())
                                + ":"
                                + getCellValue(cell)
                                + "\r\n");
                      }
                    });
              });
          Drawing<?> createDrawingPatriarch = sheet.createDrawingPatriarch();
          s.append("【Shape】\r\n");
          for (Shape shape : createDrawingPatriarch) {
            String shapeString = handleShape(shape);
            if (StringUtils.isNotBlank(shapeString)) {
              s.append("Shape:" + shapeString + "\r\n");
            }
          }
        });
    try {
      workbook.close();
    } catch (IOException e) {
      // TODO 自動生成された catch ブロック
      e.printStackTrace();
    }

    return s.toString();
  }

  private static String getCellValue(Cell cell) {
    // セルのタイプを取得
    CellType cellType = cell.getCellType();
    // セルの値が文字列の場合
    if (cellType == CellType.STRING) {
      return cell.getStringCellValue();
    }
    // セルの値が数値の場合
    else if (cellType == CellType.NUMERIC) {
      return "" + cell.getNumericCellValue();
    }
    return "";
  }

  /**
   * セルの位置情報を返す。
   *
   * <pre>
   * 引数の行番号とカラム番号から、セルの位置情報を特定し返却する。
   * 例えば左上のセルは"A1"となる。
   * </pre>
   *
   * @param aRowNum (０から始まる)行番号
   * @param aColNum (０から始まる)カラム番号
   * @return セルを位置を表す文字列
   */
  static String convertCellPos(int aRowNum, int aColNum) {
    // カラムを表すアルファベットの配列を生成
    final char[] charArray = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".toCharArray();
    final int charSize = charArray.length;
    // オフセットを取得
    int offset = aColNum / charSize;

    String cellPos = "";
    if (offset == 0) {
      cellPos = String.valueOf(charArray[aColNum]);
    } else if (offset < charSize) {
      cellPos =
          String.valueOf(charArray[offset - 1])
              + String.valueOf(charArray[aColNum - charSize * offset]);
    } else {
      throw new IllegalArgumentException("範囲外のセルが指定されています。");
    }
    return String.format("%s%d", cellPos, aRowNum + 1);
  }

  /** 指定ディレクトリ内のエクセルファイルを探す. */
  public static Collection<File> searchExcelFiles(final String aPath) {

    Collection<File> listFiles =
        FileUtils.listFiles(
            new File(aPath),
            FileFilterUtils.or(
                FileFilterUtils.suffixFileFilter("xlsx"),
                FileFilterUtils.suffixFileFilter("xls"),
                FileFilterUtils.suffixFileFilter("xlsm")),
            FileFilterUtils.trueFileFilter());

    // ~(チルダ)で始まるファイル名は対象外
    listFiles =
        listFiles.stream()
            .filter(data -> !data.getName().startsWith("~"))
            .collect(Collectors.toList());

    return listFiles;
  }

  // オートシェイプを処理するメソッド
  private static String handleShape(Object d) {
    String s = "";
    try {
      // shapeの処理(XLSX形式)
      if (d instanceof XSSFSimpleShape) {
        s = ((XSSFSimpleShape) d).getText();
      }
      // shapeの処理(XLS形式)
      if (d instanceof HSSFSimpleShape) {
        s = ((HSSFSimpleShape) d).getString().getString();
      }
      // グループ化されたshapeの処理(XLSX形式)
      if (d instanceof XSSFShapeGroup) {
        ((XSSFShapeGroup) d).forEach(gs -> handleShape(gs));
      }
      // グループ化されたshapeの処理(XLS形式)
      if (d instanceof HSSFShapeGroup) {
        ((HSSFShapeGroup) d).forEach(gs -> handleShape(gs));
      }
    } catch (Exception e) {
      logger.error("error", e);
    }
    return s;
  }
}
