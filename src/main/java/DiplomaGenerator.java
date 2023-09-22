import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class DiplomaGenerator {

    public static void main(String[] args) {
        String excelPath = "src/main/resources/assets/ejemplo.xlsx";
        String templatePath = "src/main/resources/assets/diploma_definitivo.pdf";
        String outputPath = "src/main/resources/assets/generados/DiplomasGraciasTotales_PDF.pdf";

        List<String> nombres = leerNombresDesdeExcel(excelPath);

        Document documento = null;
        PdfWriter escritor = null;

        try {
            if (templatePath.endsWith(".jpg") || templatePath.endsWith(".jpeg")) {
                Image imagen = Image.getInstance(templatePath);
                documento = new Document(imagen);
                escritor = PdfWriter.getInstance(documento, new FileOutputStream(outputPath));
                documento.open();

                for (String nombre : nombres) {
                    generarDiplomaConImg(documento, escritor, nombre, templatePath);
                    documento.newPage();
                }

            } else if (templatePath.endsWith(".pdf")) {
                // Antes de abrir el documento, establece su tamaño
                PdfReader lectorPrevio = new PdfReader(templatePath);
                Rectangle plantillaTamaño = lectorPrevio.getPageSize(1);
                documento = new Document(plantillaTamaño);
                escritor = PdfWriter.getInstance(documento, new FileOutputStream(outputPath));
                documento.open();

                for (String nombre : nombres) {
                    generarDiplomaConPdf(documento, escritor, nombre, templatePath);
                    documento.newPage();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (documento != null) {
                documento.close();
            }
        }
    }
    public static List<String> leerNombresDesdeExcel(String path) {
        List<String> nombres = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(path))) {
            XSSFWorkbook libro = new XSSFWorkbook(fis);
            XSSFSheet hoja = libro.getSheetAt(0);

            for (Row fila : hoja) {
                Cell celda = fila.getCell(0);
                if (celda != null) {
                    nombres.add(celda.getStringCellValue());
                }
            }

            libro.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return nombres;
    }
    public static void generarDiplomaConPdf(Document documento, PdfWriter escritor, String nombre, String rutaPlantillaPDF) {
        try {
            PdfReader lector = new PdfReader(rutaPlantillaPDF);
            Rectangle plantillaTamaño = lector.getPageSize(1);
            documento.setPageSize(plantillaTamaño);

            PdfImportedPage pagina = escritor.getImportedPage(lector, 1);

            // Agrega la página del PDF de la plantilla al documento
            Image instance = Image.getInstance(pagina);
            instance.setAbsolutePosition(0, 0);
            documento.add(instance);

            // Todo lo demás permanece similar para añadir el nombre al documento
            float tamanoFuente = 30;  // Tamaño predeterminado
            String rutaFuente = "src/main/resources/fonts/roboto/Roboto-Bold.ttf";
            BaseFont fuenteBase = BaseFont.createFont(rutaFuente, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            BaseColor colorTexto = new BaseColor(160, 209, 180);
            if (nombre.length() > 30) {
                tamanoFuente = 24;
            }
            if (nombre.length() > 34) {
                tamanoFuente = 22;
            }
            Font fuenteRoboto = new Font(fuenteBase, tamanoFuente, Font.NORMAL, colorTexto);
            Phrase frase = new Phrase(nombre, fuenteRoboto);
            float xInicio = 46f;
            float y = (documento.getPageSize().getHeight() / 2) + 61f;
            ColumnText.showTextAligned(escritor.getDirectContent(), Element.ALIGN_LEFT, frase, xInicio, y, 0);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void generarDiplomaConImg(Document documento, PdfWriter escritor, String nombre, String rutaPlantilla) {
        try {
            float tamanoFuente = 50;  // Tamaño predeterminado

            Image imagen = Image.getInstance(rutaPlantilla);

            // Configura la imagen para que ocupe to do el espacio del documento.
            imagen.setAbsolutePosition(0, 0);
            imagen.scaleToFit(documento.getPageSize().getWidth(), documento.getPageSize().getHeight());

            // Añade la imagen como fondo.
            documento.add(imagen);

            // Define el nombre y la fuente.
            String rutaFuente = "src/main/resources/fonts/roboto/Roboto-Bold.ttf";
            BaseFont fuenteBase = BaseFont.createFont(rutaFuente, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

            BaseColor colorTexto = new BaseColor(160, 209, 180);

            if (nombre.length() > 30) {
                tamanoFuente = 40; // Reduce aún más si es aún más largo
            }
            if (nombre.length() > 40) {
                tamanoFuente = 35; // Reduce aún más si es aún más largo
            }

            Font fuenteRoboto = new Font(fuenteBase, tamanoFuente, Font.NORMAL, colorTexto);

            Phrase frase = new Phrase(nombre, fuenteRoboto);

            // Define un punto de inicio fijo para el nombre.
            float xInicio = 95f;  // Ajusta este valor a tu necesidad.
            float y = (documento.getPageSize().getHeight() / 2) + 130f;

            // Muestra el texto comenzando desde el punto especificado en el eje X.
            ColumnText.showTextAligned(escritor.getDirectContent(), Element.ALIGN_LEFT, frase, xInicio, y, 0);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
