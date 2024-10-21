package io.github.josecarlosbran;

import com.josebran.LogsJB.LogsJB;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Objects;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CompletionException;
import java.util.concurrent.ExecutionException;

public class writeWordUtils {

    public static XWPFDocument initDoc(String ruta, InputStream plantilla) {
        XWPFDocument doc = null;
        try {
            File directory = new File(ruta).getParentFile();
            if (!directory.exists()) {
                directory.mkdirs();
            }
            writeWordUtils.deleteDoc(ruta);
            if (Objects.isNull(plantilla)) {
                doc = writeWordUtils.openDoc(ruta);
            } else {
                doc = new XWPFDocument(plantilla);
            }
            //Guardamos el archivo
            writeWordUtils.saveDoc(doc, ruta);
            doc = writeWordUtils.openDoc(ruta);
        } catch (Exception e) {
            LogsJB.error("Excepcion capturada al escribir al Inicializar del Documento" + ExceptionUtils.getStackTrace(e));
        } finally {
            return doc;
        }
    }

    public static XWPFDocument openDoc(String ruta) {
        XWPFDocument doc = null;
        try {
            FileInputStream fis = new FileInputStream(ruta);
            OPCPackage opcPackage = OPCPackage.open(fis);
            doc = new XWPFDocument(opcPackage);
        } catch (Exception e) {
            LogsJB.error("Excepcion capturada al abrir el Documento" + ExceptionUtils.getStackTrace(e));
        } finally {
            return doc;
        }
    }

    public static void saveDoc(XWPFDocument doc, String ruta) {
        try {
            writeWordUtils.deleteDoc(ruta);
            FileOutputStream out = new FileOutputStream(ruta);
            doc.write(out);
            out.close();
        } catch (Exception e) {
            LogsJB.error("Excepcion capturada al guardar el Documento" + ExceptionUtils.getStackTrace(e));
        }
    }

    public static void deleteDoc(String ruta) {
        try {
            File file = new File(ruta);
            if (file.exists()) {
                FileUtils.forceDelete(file);
            }
        } catch (Exception e) {
            LogsJB.error("Excepcion capturada al eliminar el Documento" + ExceptionUtils.getStackTrace(e));
        }
    }

    public static XWPFDocument writeTitleDoc(XWPFDocument doc
            , BreakType breakType, String title, String color, int size, boolean isBold) throws ExecutionException, InterruptedException {
        return writeWordUtils.writeTitleDocCompleteableFuture(doc, breakType, title, color, size, isBold).get();
    }

    public static CompletableFuture<XWPFDocument> writeTitleDocCompleteableFuture(XWPFDocument doc
            , BreakType breakType, String title, String color, int size, boolean isBold) {
        return CompletableFuture.supplyAsync(() -> {
            try {
                /****Agregamos el nuevo contenido****/
                //Agregamos el Titulo del Documento
                XWPFParagraph para = doc.createParagraph();
                para.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun run = para.createRun();
                run.addBreak(breakType);
                writeWordUtils.writeTextRun(run, title, color, size, isBold);
            } catch (Exception e) {
                LogsJB.error("Excepcion capturada al escribir el Titulo del Documento" + ExceptionUtils.getStackTrace(e));
                throw new CompletionException(e);
            } finally {
                return doc;
            }
        });
    }

    public static XWPFRun writeTextRun(XWPFRun run, String text, String color, int size, boolean isBold) {
        try {
            run.setColor(color);
            run.setFontSize(size);
            run.setBold(isBold);
            run.setText(text);
        } catch (Exception e) {
            LogsJB.error("Error inesperado al agregar el texto al documento de evidencia: " + text + " " + ExceptionUtils.getStackTrace(e));
        } finally {
            return run;
        }
    }

    public static XWPFDocument writeImageDoc(XWPFDocument doc, String imageRoute, String title,
                                             BreakType breakType, int spacinBefore, String color, int size, boolean isBold,
                                             int widthImage, int heightImage) throws ExecutionException, InterruptedException {
        return writeImageDocCompleteableFuture(doc, imageRoute, title, breakType, spacinBefore, color, size, isBold, widthImage, heightImage).get();
    }

    public static CompletableFuture<XWPFDocument> writeImageDocCompleteableFuture(XWPFDocument doc, String imageRoute, String title,
                                                                                  BreakType breakType, int spacinBefore, String color, int size, boolean isBold,
                                                                                  int widthImage, int heightImage) {
        return CompletableFuture.supplyAsync(() -> {
            try {
                //Verificamos si es la primera vez que se escribira el documento
                //Obtenemos las dimensiones de la imagen
                InputStream pic = new FileInputStream(imageRoute);
                //El ancho maximo de la hoja puede quedar es a 19.5 cm
                XWPFParagraph para = doc.createParagraph();
                para.setAlignment(ParagraphAlignment.LEFT);
                XWPFRun run = para.createRun();
                if (!Objects.isNull(spacinBefore)) {
                    para.setSpacingBefore(spacinBefore);
                }
                if (!Objects.isNull(breakType)) {
                    //Agregamos un unico salto de linea
                    run.addBreak(breakType);
                }
                //Agregamos el Titulo de la Imagen
                run.addTab();
                writeWordUtils.writeTextRun(run, title, color, size, isBold);
                //Agregamos la Imagen
                XWPFParagraph par = doc.createParagraph();
                par.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun runimage = par.createRun();
                //Asignamos la altura de la imagen para la siguiente validación
                runimage.addPicture(pic, Document.PICTURE_TYPE_PNG, new File(imageRoute).getName(),
                        widthImage,
                        heightImage);
                pic.close();
            } catch (Exception e) {
                LogsJB.error("Error inesperado al agregar la captura al documento de evidencia: " + imageRoute + " " + ExceptionUtils.getStackTrace(e));
                throw new CompletionException(e);
            } finally {
                return doc;
            }
        });
    }

    public static XWPFDocument writeTextDoc(XWPFDocument doc, String text, String color, int size, boolean isBold) throws ExecutionException, InterruptedException {
        return writeTextDocCompleteableFuture(doc, text, color, size, isBold).get();
    }

    public static CompletableFuture<XWPFDocument> writeTextDocCompleteableFuture(XWPFDocument doc, String text, String color, int size, boolean isBold) {
        return CompletableFuture.supplyAsync(() -> {
            try {
                //Agregamos el nuevo contenido
                XWPFParagraph para = doc.createParagraph();
                para.setAlignment(ParagraphAlignment.LEFT);
                XWPFRun run = para.createRun();
                writeWordUtils.writeTextRun(run, text, color, size, isBold);
            } catch (Exception e) {
                LogsJB.error("Error inesperado al agregar el texto al documento de evidencia: " + text + " " + ExceptionUtils.getStackTrace(e));
                throw new CompletionException(e);
            } finally {
                return doc;
            }
        });
    }
}
