package com.landscapehub.decryptor;
import java.io.File;
import java.io.IOException;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import com.landscapehub.decryptor.GenericDecryptor;
import com.landscapehub.decryptor.XLSDecryptor;
import com.landscapehub.decryptor.XLSXDecryptor;

// Reference:
// https://www.codejava.net/coding/java-example-to-read-password-protected-excel-files-using-apache-poi
//
public class App 
{
    public static void main( String[] args ) {
        String inputFile = args[0];
        String outputFile = args[1];
        String password = args[2];

        try {
            NPOIFSFileSystem fileSystem = new NPOIFSFileSystem(new File(inputFile));
            
            GenericDecryptor decryptor;

            if (inputFile.endsWith("xls")) {
                decryptor = new XLSDecryptor();
            } else {
                decryptor = new XLSXDecryptor();
            }

            decryptor.run(fileSystem, outputFile, password);
            fileSystem.close();

        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}
