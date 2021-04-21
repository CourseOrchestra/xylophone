package ru.curs.xylophone;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.approvaltests.Approvals;
import org.approvaltests.approvers.FileApprover;
import org.approvaltests.core.Options;
import org.approvaltests.writers.ApprovalBinaryFileWriter;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import org.lambda.functions.Function2;

import java.io.*;

import static bad.robot.excel.matchers.Matchers.sameWorkbook;


public class TestOverall {
    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

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

	@Test
	public void test1() throws XylophoneError {
		InputStream descrStream = TestReader.class
				.getResourceAsStream("testdescriptor3.json");
		InputStream dataStream = TestReader.class
				.getResourceAsStream("testdata.xml");
		InputStream templateStream = TestReader.class
				.getResourceAsStream("template.xlsx");

		// write results to binary buffer
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		XML2Spreadsheet.process(dataStream, descrStream, templateStream,
				OutputType.XLSX, false, false, bos);
		byte[] writtenData = bos.toByteArray();

		// verify it
		Options options = new Options();
		Approvals.verify(
				new FileApprover(
					new ApprovalBinaryFileWriter(new ByteArrayInputStream(writtenData),
							"xlsx"),
					options.forFile().getNamer(),
					compareSpreadsheetFiles
				),
				options
		);
	}

	@Test
	public void test2() throws XylophoneError {
		InputStream descrStream = TestReader.class
				.getResourceAsStream("testsaxdescriptor3.json");
		InputStream dataStream = TestReader.class
				.getResourceAsStream("testdata.xml");
		InputStream templateStream = TestReader.class
				.getResourceAsStream("template.xlsx");

		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		XML2Spreadsheet.process(dataStream, descrStream, templateStream,
				OutputType.XLSX, true, false, bos);
		byte[] writtenData = bos.toByteArray();

		// verify it
		Options options = new Options();
		Approvals.verify(
				new FileApprover(
						new ApprovalBinaryFileWriter(new ByteArrayInputStream(writtenData),
								"xlsx"),
						options.forFile().getNamer(),
						compareSpreadsheetFiles
				),
				options
		);
	}

}
