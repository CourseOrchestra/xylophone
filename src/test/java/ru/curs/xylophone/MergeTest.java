package ru.curs.xylophone;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.*;
import java.net.URISyntaxException;
import java.nio.file.Paths;

public class MergeTest {
    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void mergeUpTest() throws URISyntaxException, IOException, XML2SpreadSheetError {
        File descriptor = Paths.get(TestOverall.class.getResource("testdescriptor.xml").toURI()).toFile();
        InputStream dataStream = TestReader.class
                .getResourceAsStream("test_data_with_spec_symbols.xml");
        File template = Paths.get(TestOverall.class.getResource("mergeup_template.xls").toURI()).toFile();

        File createdTempOutputFile = temporaryFolder.newFile("temp.xls");

        try (OutputStream outputStream = new FileOutputStream(createdTempOutputFile)) {
            XML2Spreadsheet.process(dataStream, descriptor, template,
                    false, false, outputStream);
        }
    }

    @Test
    public void mergeLeftTest() throws URISyntaxException, IOException, XML2SpreadSheetError {
        File descriptor = Paths.get(TestOverall.class.getResource("testdescriptor.xml").toURI()).toFile();
        InputStream dataStream = TestReader.class
                .getResourceAsStream("test_data_with_spec_symbols.xml");
        File template = Paths.get(TestOverall.class.getResource("mergeleft_template.xls").toURI()).toFile();

        File createdTempOutputFile = temporaryFolder.newFile("temp.xls");

        try (OutputStream outputStream = new FileOutputStream(createdTempOutputFile)) {
            XML2Spreadsheet.process(dataStream, descriptor, template,
                    false, false, outputStream);
        }
    }

    @Test
    public void mergeTest() throws IOException, URISyntaxException, XML2SpreadSheetError {
        File descriptor = Paths.get(TestOverall.class.getResource("testdescriptor.xml").toURI()).toFile();
        InputStream dataStream = TestReader.class
                .getResourceAsStream("test_data_with_spec_symbols.xml");
        File template = Paths.get(TestOverall.class.getResource("merge_template.xls").toURI()).toFile();

        File createdTempOutputFile = temporaryFolder.newFile("temp.xls");

        try (OutputStream outputStream = new FileOutputStream(createdTempOutputFile)) {
            XML2Spreadsheet.process(dataStream, descriptor, template,
                    false, false, outputStream);
        }
    }

    @Test
    public void mergeComplexTest() throws IOException, URISyntaxException, XML2SpreadSheetError {
        File descriptor = Paths.get(TestOverall.class.getResource("merge_descriptor.xml").toURI()).toFile();
        InputStream dataStream = TestReader.class
                .getResourceAsStream("merge_data.xml");
        File template = Paths.get(TestOverall.class.getResource("merge_complex_template.xls").toURI()).toFile();

        File createdTempOutputFile = temporaryFolder.newFile("temp.xls");

        try (OutputStream outputStream = new FileOutputStream(createdTempOutputFile)) {
            XML2Spreadsheet.process(dataStream, descriptor, template,
                    false, false, outputStream);
        }
    }
}
