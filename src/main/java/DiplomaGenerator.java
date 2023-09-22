import com.itextpdf.text.*;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class DiplomaGenerator {

    public static void main(String[] args) {
        String excelPath = "C:\\Users\\joser\\OneDrive\\Documents\\MisCodigos\\DiplomaPDF\\ejemplo.xlsx";
        String templatePath = "C:\\Users\\joser\\OneDrive\\Documents\\MisCodigos\\DiplomaPDF\\fotodiploma.jpg";
        String outputPath = "C:\\Users\\joser\\OneDrive\\Documents\\MisCodigos\\DiplomaPDF\\generados";

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
            Font font = new Font(Font.FontFamily.HELVETICA, 24, Font.BOLD);
            Phrase phrase = new Phrase(name, font);

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
