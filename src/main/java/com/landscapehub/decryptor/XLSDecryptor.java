package com.landscapehub.decryptor;
import java.io.IOException;
import java.io.FileOutputStream;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;

class XLSDecryptor implements GenericDecryptor {

    public void run(NPOIFSFileSystem fileSystem, String outputFile, String password) throws IOException {
        Biff8EncryptionKey.setCurrentUserPassword(password);
            
        Workbook workbook = new HSSFWorkbook(fileSystem);

        Biff8EncryptionKey.setCurrentUserPassword(null);

        try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
            workbook.write(outputStream);
        }

        workbook.close();
    }
}
