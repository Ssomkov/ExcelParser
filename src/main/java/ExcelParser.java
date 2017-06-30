import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;


class ExcelParser {

    private static HSSFWorkbook wb = null;
    private static HashMap<String, Integer> map = new HashMap<>();
    private static HashMap<String, Integer> exclusions = new HashMap<>();
    private static String fileName;

    static {
        exclusions.put("удаление", 0);
        exclusions.put("создание", 0);
        exclusions.put("полное", 0);
        exclusions.put("было фото", 0);
    }

    static void parse(String name) {
        fileName = name;
        String coffee;
        openFile();
        Sheet sheet = wb.getSheetAt(0);
        for (Row row : sheet) {
            for (Cell cell : row) {
                CellType cellType = cell.getCellTypeEnum();
                if (cellType == CellType.STRING) {
                    coffee = cell.getStringCellValue();
                    if ((!map.containsKey(coffee)) || (exclusions.containsKey(coffee))) {
                        map.put(coffee, 0);
                    } else {
                        changeCellColor(cell);
                    }
                }
            }
        }
        saveDataToFile();
    }

    private static void changeCellColor(Cell cell) {
        CellStyle style = wb.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        cell.setCellStyle(style);
    }

    private static void saveDataToFile() {
        try (FileOutputStream fos = new FileOutputStream("out_" + fileName)) {
            wb.write(fos);
            wb.close();
        } catch (Exception e) {
            System.out.println("Can't write to file");
        }
    }

    private static void openFile() {
        try (InputStream in = new FileInputStream(fileName)) {
            wb = new HSSFWorkbook(in);
        } catch (IOException e) {
            System.out.println("Can't open file");
        }
    }
}
