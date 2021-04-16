package ru.curs.xylophone;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.approvaltests.Approvals;
import org.approvaltests.approvers.FileApprover;
import org.approvaltests.core.Options;
import org.approvaltests.writers.ApprovalBinaryFileWriter;
import org.lambda.functions.Function2;

import java.io.*;

import static bad.robot.excel.matchers.Matchers.sameWorkbook;

public class FullApprovalsTester {
    /**
     * Функция, сравнивающая файлы с содержимым-таблицами.
     *
     *  Внутри вызывает matchesSafely нескольких классов в bad.robot.excel.matchers. Это
     * 		 WorkbookMatcher
     * 		   SheetsMatcher.hasSameSheetsAs
     * 		     SheetNumberMatcher.hasSameNumberOfSheetsAs
     * 		     SheetNameMatcher.containsSameNamedSheetsAs
     * 		   SheetsMatcher
     * 		     RowNumberMatcher.hasSameNumberOfRowAs
     * 		     RowsMatcher.hasSameRowsAs
     * 		       RowInSheetMatcher.hasSameRow
     * 		         RowMissingMatcher.rowIsPresent
     * 		         CellNumberMatcher.hasSameNumberOfCellsAs
     * 		         CellsMatcher.hasSameCellsAs
     * 		           CellInRowMatcher.hasSameCell
     * 		             bad.robot.excel.matchers.CellType.adaptPoi -> bad.robot.excel.cell.Cell
     * 		             Cell: StringCell->StyledCell
     * 		             StyledCell.equals = org.apache.commons.lang3.builder.EqualsBuilder.reflectionEquals
     *
     */
    public Function2<File, File, Boolean> compareSpreadsheetFiles = (actualFile, expectedFile) -> {
        try {
            Workbook actual = new XSSFWorkbook(actualFile);
            Workbook expected = new XSSFWorkbook(expectedFile);
            return sameWorkbook(expected).matches(actual);
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
            return false;
        }
    };

    public void approvalTest(String descriptorPath, String dataPath, String templatePath,
                             OutputType outputType, boolean useSax)
            throws XML2SpreadSheetError {
        InputStream descrStream = TestReader.class.getResourceAsStream(descriptorPath);
        InputStream dataStream = TestReader.class.getResourceAsStream(dataPath);
        InputStream templateStream = TestReader.class.getResourceAsStream(templatePath);

        // write results to binary buffer
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        XML2Spreadsheet.process(dataStream, descrStream, templateStream, outputType, useSax, bos);
        byte[] writtenData = bos.toByteArray();

        // verify it
        Options options = new Options();
        Approvals.verify(
                new FileApprover(
                        new ApprovalBinaryFileWriter(new ByteArrayInputStream(writtenData),
                                outputType.getExtension()),
                        options.forFile().getNamer(),
                        compareSpreadsheetFiles
                ),
                options
        );
    }
}
