import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ReadFileXlsx {

    private static Path arquivo = Paths.get("./src/main/resources", "planilha-apache-poi.xlsx");

    public static void main(String[] args) {

        try {
            System.out.println("Inicianlizando Leitura do arquivo..");
            try (InputStream inputStream = new FileInputStream(arquivo.toFile())) {
                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

                XSSFSheet sheet = workbook.getSheetAt(0);
                var valorProduto = getNewCell(sheet, 3, 3);
                valorProduto.setCellValue(750.5);

                var formulaEvaluator = new XSSFFormulaEvaluator(workbook);
                formulaEvaluator.evaluateAll();

                inputStream.close();

                try (OutputStream file = new FileOutputStream(arquivo.toFile())) {

                    workbook.write(file);
                    System.out.println("Edição realizada com sucesso.");
                }
                workbook.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Cell getNewCell(XSSFSheet sheet, int indexRow, int indexCell) {
        Row row = sheet.getRow(indexRow);
        Cell cell = row.getCell(indexCell);
        return cell;
    }
}
