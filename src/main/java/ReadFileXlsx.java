import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.logging.Logger;

import static java.util.Objects.isNull;

public class ReadFileXlsx {

    private static Path arquivo = Paths.get("./src/main/resources", "planilha-apache-poi.xlsx");
    static Logger log = Logger.getLogger(ReadFileXlsx.class.getName());

    public static void main(String[] args) {

        try {
            log.info("Inicianlizando Leitura do arquivo..");
            try (InputStream inputStream = new FileInputStream(arquivo.toFile())) {
                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

                XSSFSheet sheet = workbook.getSheetAt(0);
                var valorProduto = getNewCell(sheet, 3, 3);
                valorProduto.setCellValue(520);


                var tituloValorD5 = getNewCell(sheet,4,3);
                tituloValorD5.setCellValue("VALOR");
                var valorCellD6 = getNewCell(sheet,5,3);
                valorCellD6.setCellValue(Double.valueOf(500.30));

                var quantidadeD4 = getNewCell(sheet,4,4);
                quantidadeD4.setCellValue("QUANTIDADE");
                var quantidadeCellE6 = getNewCell(sheet,5,4);
                quantidadeCellE6.setCellValue(2.0);

                var total = getNewCell(sheet,5,5);
                total.setCellFormula("SUM(D6:E6)");

                var multiplicar = getNewCell(sheet,5,6);
                multiplicar.setCellFormula("D6*E6");

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
        if(isNull(cell)){
            cell = row.createCell(indexCell);
        }
        return cell;
    }
}
