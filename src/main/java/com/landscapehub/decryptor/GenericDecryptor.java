package com.landscapehub.decryptor;
import java.io.IOException;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;

interface GenericDecryptor {
    public void run(NPOIFSFileSystem fileSystem, String outputFile, String password) throws IOException;
}
