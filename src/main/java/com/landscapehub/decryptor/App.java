package com.landscapehub.decryptor;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.FileOutputStream;
import java.security.GeneralSecurityException;
import java.util.Iterator;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        String excelFilePath = "avail.xlsx";

        try {
            NPOIFSFileSystem fileSystem = new NPOIFSFileSystem(new File(excelFilePath));
            EncryptionInfo info = new EncryptionInfo(fileSystem);
            Decryptor decryptor = Decryptor.getInstance(info);
             
            if (!decryptor.verifyPassword("plants")) {
                System.out.println("Unable to process: document is encrypted.");
                return;
            }
             
            InputStream dataStream = decryptor.getDataStream(fileSystem);
             
            Workbook workbook = new XSSFWorkbook(dataStream);
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = firstSheet.iterator();

            try (FileOutputStream outputStream = new FileOutputStream("JavaBooks.xlsx")) {
                workbook.write(outputStream);
            }
             
            while (iterator.hasNext()) {
                Row nextRow = iterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();
                 
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    System.out.print(cell.getStringCellValue() + "\t");
                }
                System.out.println();
            }
             
            workbook.close();
            dataStream.close();    
            fileSystem.close();
        } catch (GeneralSecurityException | IOException ex) {
            ex.printStackTrace();
        }
        System.out.println( "Hello World!" );
    }
}
