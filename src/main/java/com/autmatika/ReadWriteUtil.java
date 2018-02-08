package com.autmatika;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ReadWriteUtil {

    public Path writeFilesInTargetFolder(String nestedFolderName, String fileName, byte[] textToWrite) throws IOException {

        Path targetFolder = Paths.get(getClass().getProtectionDomain().getCodeSource().getLocation().getPath()).getParent();
        Path jsonsFolder = Paths.get(targetFolder.toUri().getPath(), nestedFolderName);
        Files.createDirectories(jsonsFolder);

        Path sheetJsonPath = Paths.get(jsonsFolder.toUri().getPath(), fileName);

        return Files.write(sheetJsonPath, textToWrite);
    }

    public String readFileFromTargerFolder(String nestedFolderName, String fileName) throws IOException {

        Path targetFolder = Paths.get(getClass().getProtectionDomain().getCodeSource().getLocation().getPath()).getParent();
        Path jsonsFolder = Paths.get(targetFolder.toUri().getPath(), nestedFolderName);
        Path jsonFilePath = Paths.get(jsonsFolder.toUri().getPath(), fileName);

        return new String(Files.readAllBytes(jsonFilePath), StandardCharsets.UTF_8);
    }

    public String readFileFromTargerFolder(Path path) throws IOException {

        return new String(Files.readAllBytes(path), StandardCharsets.UTF_8);
    }
}
