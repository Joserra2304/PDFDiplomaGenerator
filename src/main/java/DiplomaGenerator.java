import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import static org.apache.logging.log4j.core.util.Loader.getClassLoader;

public class DiplomaGenerator {

    public static void main(String[] args) {
        String excelPath = "src/main/resources/assets/ejemplo.xlsx";
        String templatePath = "src/main/resources/assets/fotodiploma.jpg";
        String outputPath = "src/main/resources/assets/generados";

        List<String> names = readNamesFromExcel(excelPath);

        for (String name : names) {
            generateDiploma(name, templatePath, outputPath + File.separator + name + ".pdf");
        }
    }
    public static List<String> readNamesFromExcel(String path) {
        List<String> names = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(path))) {
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                Cell cell = row.getCell(0);
                if (cell != null) {
                    names.add(cell.getStringCellValue());
                }
            }

            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return names;
    }

    public static void generateDiploma(String name, String templatePath, String outputPath) {
        try {
            Image img = Image.getInstance(templatePath);
            Document document = new Document(img);
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPath));
            document.open();

            // Configura la imagen para que ocupe to do el espacio del documento.
            img.setAbsolutePosition(0, 0);
            img.scaleToFit(document.getPageSize().getWidth(), document.getPageSize().getHeight());

            // Añade la imagen como fondo.
            document.add(img);

            // Define el nombre y la fuente.
            String fontPath = "src/main/resources/fonts/roboto/Roboto-Bold.ttf";
            BaseFont baseFont = BaseFont.createFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);


            Font robotoFont = new Font(baseFont, 24); // 24 es el tamaño de la fuente.

            Phrase phrase = new Phrase(name, robotoFont);

            // Obtiene las coordenadas del centro de la página.
            float x = document.getPageSize().getWidth() / 2;
            float y = (document.getPageSize().getHeight() / 2) - 30;  // Ajusta este valor según lo necesites.

            // Muestra el texto centrado en el centro de la página.
            ColumnText.showTextAligned(writer.getDirectContent(), Element.ALIGN_CENTER, phrase, x, y, 0);

            document.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }



}
