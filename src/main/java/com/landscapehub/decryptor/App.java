package com.landscapehub.decryptor;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.FileOutputStream;
import java.security.GeneralSecurityException;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App 
{
    public static void main( String[] args )
    {
        String excelFilePath = "avail.xlsx";

        // String inputFile = args[0];
        // String outputFile = args[1];
        // String password = args[2];

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

            try (FileOutputStream outputStream = new FileOutputStream("JavaBooks.xlsx")) {
                workbook.write(outputStream);
            }
             
            workbook.close();
            dataStream.close();    
            fileSystem.close();
        } catch (GeneralSecurityException | IOException ex) {
            ex.printStackTrace();
        }
    }
}
