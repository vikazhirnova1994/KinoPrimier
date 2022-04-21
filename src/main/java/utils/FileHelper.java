package utils;

import model.NewFilm;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class FileHelper {

    public static File createFile(String name) {
        try {
            String path = new File("").getAbsolutePath();
            File newFile = new File(path + name);
            if (newFile.createNewFile()) {
                System.out.print("Файл создан, ");
                System.out.println( "можно найти по ссылке: " + newFile.getAbsolutePath());
            } else {
                System.out.print("Файл уже существует, ");
                System.out.println( "можно найти по ссылке: " + newFile.getAbsolutePath());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new File(new File("").getAbsolutePath() + name);
    }

    public static boolean writeExcel(File file, List<NewFilm> newFilms){
        int rownum = 0;
        int cellnum = 0;
        boolean isSuccessWrite = false;
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("NewFilms");
        for (NewFilm newFilm : newFilms) {
            Row row = sheet.createRow(rownum++);
            Cell cell = row.createCell(cellnum);
            cell.setCellValue(newFilm.getName() + ": " + newFilm.getLink());
        }
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            book.write(fileOutputStream);
            isSuccessWrite = true;
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
            isSuccessWrite = false;
        }
        return isSuccessWrite;
    }
}
