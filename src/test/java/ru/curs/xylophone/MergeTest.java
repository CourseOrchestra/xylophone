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
    public void mergeUpTest() throws URISyntaxException, IOException, XylophoneError {
        File descriptorF = Paths.get(TestOverall.class.getResource("testdescriptor.json").toURI()).toFile();
        FileInputStream descriptor = new FileInputStream(descriptorF);
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
    public void mergeLeftTest() throws URISyntaxException, IOException, XylophoneError {
        File descriptorF = Paths.get(TestOverall.class.getResource("testdescriptor.json").toURI()).toFile();
        FileInputStream descriptor = new FileInputStream(descriptorF);
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
    public void mergeTest() throws URISyntaxException, IOException, XylophoneError {
        File descriptorF = Paths.get(TestOverall.class.getResource("testdescriptor.json").toURI()).toFile();
        FileInputStream descriptor = new FileInputStream(descriptorF);
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
    public void mergeComplexTest() throws URISyntaxException, IOException, XylophoneError {
        File descriptorF = Paths.get(TestOverall.class.getResource("merge_descriptor.json").toURI()).toFile();
        FileInputStream descriptor = new FileInputStream(descriptorF);
        InputStream dataStream = TestReader.class
                .getResourceAsStream("merge_data.xml");
        File template = Paths.get(TestOverall.class.getResource("merge_complex_template.xls").toURI()).toFile();

        File createdTempOutputFile = temporaryFolder.newFile("temp.xls");

        try (OutputStream outputStream = new FileOutputStream(createdTempOutputFile)) {
            XML2Spreadsheet.process(dataStream, descriptor, template,
                    false, false, outputStream);
        }
    }

    @Test
    public void mergeComplexTestWithLeftUp() throws URISyntaxException, IOException, XylophoneError {
        File descriptorF = Paths.get(TestOverall.class.getResource("merge_descriptor.json").toURI()).toFile();
        FileInputStream descriptor = new FileInputStream(descriptorF);
        InputStream dataStream = TestReader.class
                .getResourceAsStream("merge_data.xml");
        File template = Paths.get(TestOverall.class.getResource(
                "merge_leftup_complex_template.xls").toURI()).toFile();

        File createdTempOutputFile = temporaryFolder.newFile("temp.xls");

        try (OutputStream outputStream = new FileOutputStream(createdTempOutputFile)) {
            XML2Spreadsheet.process(dataStream, descriptor, template,
                    false, false, outputStream);
        }
    }

    @Test
    public void mergeComplexTestWithUpLeft() throws URISyntaxException, IOException, XylophoneError {

        File descriptorF = Paths.get(TestOverall.class.getResource("merge_descriptor_up_left.json").toURI()).toFile();
        FileInputStream descriptor = new FileInputStream(descriptorF);
        InputStream dataStream = TestReader.class
                .getResourceAsStream("merge_data_up_left.xml");
        File template = Paths.get(TestOverall.class.getResource(
                "merge_upleft_complex_template.xls").toURI()).toFile();

        File createdTempOutputFile = temporaryFolder.newFile("temp.xls");

        try (OutputStream outputStream = new FileOutputStream(createdTempOutputFile)) {
            XML2Spreadsheet.process(dataStream, descriptor, template,
                    false, false, outputStream);
        }
    }
}
