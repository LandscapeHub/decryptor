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

// Reference:
// https://www.codejava.net/coding/java-example-to-read-password-protected-excel-files-using-apache-poi
//
public class App 
{
    public static void main( String[] args )
    {
        String inputFile = args[0];
        String outputFile = args[1];
        String password = args[2];

        try {
            NPOIFSFileSystem fileSystem = new NPOIFSFileSystem(new File(inputFile));
            EncryptionInfo info = new EncryptionInfo(fileSystem);
            Decryptor decryptor = Decryptor.getInstance(info);
             
            if (!decryptor.verifyPassword(password)) {
                System.out.println("Unable to process: document is encrypted.");
                return;
            }
             
            InputStream dataStream = decryptor.getDataStream(fileSystem);
             
            Workbook workbook = new XSSFWorkbook(dataStream);

            try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
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
