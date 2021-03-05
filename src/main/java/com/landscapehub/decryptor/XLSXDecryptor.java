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
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;

class XLSXDecryptor implements GenericDecryptor {

    public void run(NPOIFSFileSystem fileSystem, String outputFile, String password) throws IOException {
        EncryptionInfo info = new EncryptionInfo(fileSystem);
        Decryptor decryptor = Decryptor.getInstance(info);

        try {
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

        } catch (GeneralSecurityException ex) {
            ex.printStackTrace();
        }
    }
}