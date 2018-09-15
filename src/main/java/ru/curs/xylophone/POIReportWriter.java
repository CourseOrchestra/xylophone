/*
   (с) 2016 ООО "КУРС-ИТ"

   Этот файл — часть КУРС:Xylophone.

   КУРС:Xylophone — свободная программа: вы можете перераспространять ее и/или изменять
   ее на условиях Стандартной общественной лицензии ограниченного применения GNU (LGPL)
   в том виде, в каком она была опубликована Фондом свободного программного обеспечения; либо
   версии 3 лицензии, либо (по вашему выбору) любой более поздней версии.

   Эта программа распространяется в надежде, что она будет полезной,
   но БЕЗО ВСЯКИХ ГАРАНТИЙ; даже без неявной гарантии ТОВАРНОГО ВИДА
   или ПРИГОДНОСТИ ДЛЯ ОПРЕДЕЛЕННЫХ ЦЕЛЕЙ. Подробнее см. в Стандартной
   общественной лицензии GNU.

   Вы должны были получить копию Стандартной общественной лицензии  ограниченного
   применения GNU (LGPL) вместе с этой программой. Если это не так,
   см. http://www.gnu.org/licenses/.


   Copyright 2016, COURSE-IT Ltd.

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Lesser General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Lesser General Public License for more details.
   You should have received a copy of the GNU Lesser General Public License
   along with this program.  If not, see http://www.gnu.org/licenses/.

*/
package ru.curs.xylophone;

import java.io.IOException;
import java.io.InputStream;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Класс, ответственный за формирование результирующего вывода в документ при
 * помощи библиотеки POI.
 */
abstract class POIReportWriter extends ReportWriter {

    /**
     * Регексп для числа (в общем случае, с плавающей точкой --- целые числа
     * также должны попадать под этот регексп.
     */
    private static final Pattern NUMBER = Pattern
            .compile("[+-]?\\d+(\\.\\d+)?([eE][+-]?\\d+)?");
    /**
     * Регексп для даты в ISO-формате.
     */
    private static final Pattern DATE = Pattern
            .compile("(\\d\\d\\d\\d)-(0[1-9]|1[0-2])-(0[1-9]|[12]\\d|3[01])");
    private final Workbook template;
    private final Workbook result;
    private Sheet activeTemplateSheet;
    private Sheet activeResultSheet;
    private boolean needEval = false;
    private final Map<CellStyle, CellStyle> stylesMap = new HashMap<>();

    POIReportWriter(InputStream template, InputStream templateCopy)
            throws XML2SpreadSheetError {
        try {
            this.template = WorkbookFactory.create(template);
            // Создаём новую книгу
            result = createResultWb(templateCopy);
        } catch (InvalidFormatException | IOException e) {
            throw new XML2SpreadSheetError(e.getMessage());
        }
        final Map<Short, Font> fontMap = new HashMap<>();

        // Копируем шрифты
        // Внимание: в цикле --- <=, а не < из-за ошибки то ли в названии,
        // то ли в реализации метода getNumberOfFonts ;-)
        for (short i = 0; i <= this.template.getNumberOfFonts(); i++) {
            Font fSource;
            try {
                // Но в некоторых файлах тут случается вот такое..
                fSource = this.template.getFontAt(i);
            } catch (IndexOutOfBoundsException e) {
                break;
            }
            Font fResult = (i == 0) ? result.getFontAt((short) 0) : result
                    .createFont();
            // Для XLSX, похоже, не работает...
            if (this instanceof XLSReportWriter)
                fResult.setCharSet(fSource.getCharSet());
            fResult.setColor(fSource.getColor());
            fResult.setFontHeight(fSource.getFontHeight());
            fResult.setFontName(fSource.getFontName());
            fResult.setItalic(fSource.getItalic());
            fResult.setStrikeout(fSource.getStrikeout());
            fResult.setTypeOffset(fSource.getTypeOffset());
            fResult.setUnderline(fSource.getUnderline());
            fResult.setBold(fSource.getBold());
            fontMap.put(fSource.getIndex(), fResult);
        }

        DataFormat df = result.createDataFormat();

        // Копируем стили ячеек (cloneStyleFrom не работает для нас)
        for (short i = 0; i < this.template.getNumCellStyles(); i++) {

            CellStyle csSource = this.template.getCellStyleAt(i);
            CellStyle csResult = result.createCellStyle();

            csResult.setAlignment(csSource.getAlignmentEnum());
            csResult.setBorderBottom(csSource.getBorderBottomEnum());
            csResult.setBorderLeft(csSource.getBorderLeftEnum());
            csResult.setBorderRight(csSource.getBorderRightEnum());
            csResult.setBorderTop(csSource.getBorderTopEnum());
            csResult.setBottomBorderColor(csSource.getBottomBorderColor());
            csResult.setDataFormat(df.getFormat(csSource.getDataFormatString()));
            csResult.setFillBackgroundColor(csSource.getFillBackgroundColor());
            csResult.setFillForegroundColor(csSource.getFillForegroundColor());
            csResult.setFillPattern(csSource.getFillPatternEnum());
            Font f = fontMap.get(csSource.getFontIndex());
            if (f != null)
                csResult.setFont(f);

            csResult.setHidden(csSource.getHidden());
            csResult.setIndention(csSource.getIndention());
            csResult.setLeftBorderColor(csSource.getLeftBorderColor());
            csResult.setLocked(csSource.getLocked());
            csResult.setRightBorderColor(csSource.getRightBorderColor());
            csResult.setRotation(csSource.getRotation());
            csResult.setTopBorderColor(csSource.getTopBorderColor());
            csResult.setVerticalAlignment(csSource.getVerticalAlignmentEnum());
            csResult.setWrapText(csSource.getWrapText());

            stylesMap.put(csSource, csResult);
        }

    }

    abstract Workbook createResultWb(InputStream is)
            throws InvalidFormatException, IOException;

    private void updateActiveTemplateSheet(String sourceSheet)
            throws XML2SpreadSheetError {
        if (sourceSheet != null)
            activeTemplateSheet = template.getSheet(sourceSheet);
        if (activeTemplateSheet == null)
            activeTemplateSheet = template.getSheetAt(0);
        if (activeTemplateSheet == null)
            throw new XML2SpreadSheetError(String.format(
                    "Sheet '%s' does not exist.", sourceSheet));
    }

    @Override
    void newSheet(String sheetName, String sourceSheet,
            int startRepeatingColumn, int endRepeatingColumn,
            int startRepeatingRow, int endRepeatingRow)
            throws XML2SpreadSheetError {

        updateActiveTemplateSheet(sourceSheet);
        activeResultSheet = result.getSheet(sheetName);
        if (activeResultSheet != null)
            return;
        activeResultSheet = result.createSheet(sheetName);

        // Ищем число столбцов в исходнике
        int maxCol = 1;
        for (int i = activeTemplateSheet.getFirstRowNum(); i <= activeTemplateSheet
                .getLastRowNum(); i++) {
            Row r = activeTemplateSheet.getRow(i);
            if (r == null)
                continue;
            int c = r.getLastCellNum();
            if (c > maxCol)
                maxCol = c;
        }
        // Копируем ширины колонок (знак <, а не <= здесь не случайно, т. к.
        // getLastCellNum возвращает ширину строки ПЛЮС ЕДИНИЦА)
        for (int i = 0; i < maxCol; i++) {
            activeResultSheet.setColumnWidth(i,
                    activeTemplateSheet.getColumnWidth(i));
            // Скрытые столбцы
            activeResultSheet.setColumnHidden(i,
                    activeTemplateSheet.isColumnHidden(i));
            // Столбцы с разрывом страницы
            if (activeTemplateSheet.isColumnBroken(i))
                activeResultSheet.setColumnBreak(i);
        }
        // Переносим дефолтную высоту
        activeResultSheet.setDefaultRowHeight(activeTemplateSheet
                .getDefaultRowHeight());
        // Копируем все настройки печати
        PrintSetup sourcePS = activeTemplateSheet.getPrintSetup();
        PrintSetup resultPS = activeResultSheet.getPrintSetup();
        resultPS.setCopies(sourcePS.getCopies());
        resultPS.setDraft(sourcePS.getDraft());
        resultPS.setFitHeight(sourcePS.getFitHeight());
        resultPS.setFitWidth(sourcePS.getFitWidth());
        resultPS.setFooterMargin(sourcePS.getFooterMargin());
        resultPS.setHeaderMargin(sourcePS.getHeaderMargin());
        resultPS.setHResolution(sourcePS.getHResolution());
        resultPS.setLandscape(sourcePS.getLandscape());
        resultPS.setLeftToRight(sourcePS.getLeftToRight());
        resultPS.setNoColor(sourcePS.getNoColor());
        resultPS.setNoOrientation(sourcePS.getNoOrientation());
        resultPS.setNotes(sourcePS.getNotes());
        resultPS.setPageStart(sourcePS.getPageStart());
        resultPS.setPaperSize(sourcePS.getPaperSize());
        resultPS.setScale(sourcePS.getScale());
        resultPS.setUsePage(sourcePS.getUsePage());
        resultPS.setValidSettings(sourcePS.getValidSettings());
        resultPS.setVResolution(sourcePS.getVResolution());
        resultPS.setHResolution(sourcePS.getHResolution());

        activeResultSheet.setFitToPage(activeTemplateSheet.getFitToPage());
        for (short i = 0; i < 4; i++)
            activeResultSheet.setMargin(i, activeTemplateSheet.getMargin(i));
        activeResultSheet.setDisplayZeros(activeTemplateSheet.isDisplayZeros());

        // Копируем колонтитулы
        Header resultH = activeResultSheet.getHeader();
        Header sourceH = activeTemplateSheet.getHeader();
        resultH.setCenter(sourceH.getCenter());
        resultH.setRight(sourceH.getRight());
        resultH.setLeft(sourceH.getLeft());

        Footer resultF = activeResultSheet.getFooter();
        Footer sourceF = activeTemplateSheet.getFooter();
        resultF.setCenter(sourceF.getCenter());
        resultF.setLeft(sourceF.getLeft());
        resultF.setRight(sourceF.getRight());

        // Копируем сквозные ячейки
        if (startRepeatingRow >= 0)
            activeResultSheet.setRepeatingRows(new CellRangeAddress(
                    startRepeatingRow, endRepeatingRow, -1, -1));
        if (startRepeatingColumn >= 0)
            activeResultSheet.setRepeatingColumns(new CellRangeAddress(-1, -1,
                    startRepeatingColumn, endRepeatingColumn));
    }

    @Override
    void putSection(XMLContext context, CellAddress growthPoint,
            String sourceSheet, RangeAddress range) throws XML2SpreadSheetError {
        updateActiveTemplateSheet(sourceSheet);
        if (activeResultSheet == null)
            sheet("Sheet1", sourceSheet, -1, -1, -1, -1);

        for (int i = range.top(); i <= range.bottom(); i++) {
            Row sourceRow = activeTemplateSheet.getRow(i - 1);
            if (sourceRow == null)
                continue;
            Row resultRow = activeResultSheet.getRow(growthPoint.getRow() + i
                    - range.top() - 1);
            if (resultRow == null)
                resultRow = activeResultSheet.createRow(growthPoint.getRow()
                        + i - range.top() - 1);

            // Высоты строк (если отличаются от дефолтной высоты)
            if (sourceRow.getHeight() != activeTemplateSheet
                    .getDefaultRowHeight())
                resultRow.setHeight(sourceRow.getHeight());
            // Скрытые строки
            resultRow.setZeroHeight(sourceRow.getZeroHeight());

            for (int j = range.left(); j < Math.min(range.right(),
                    sourceRow.getLastCellNum()) + 1; j++) {
                Cell sourceCell = sourceRow.getCell(j - 1);
                if (sourceCell == null)
                    continue;
                Cell resultCell = resultRow.createCell(growthPoint.getCol() + j
                        - range.left() - 1);

                // Копируем стиль...
                CellStyle csResult = stylesMap.get(sourceCell.getCellStyle());
                if (csResult != null)
                    resultCell.setCellStyle(csResult);

                // Копируем значение...
                String val;
                String buf;
                switch (sourceCell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    resultCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    resultCell.setCellValue(sourceCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    // ДЛЯ СТРОКОВЫХ ЯЧЕЕК ВЫЧИСЛЯЕМ ПОДСТАНОВКИ!!
                    val = sourceCell.getStringCellValue();
                    buf = context.calc(val);
                    // Если ячейка содержит строковое представление числа и при
                    // этом содержит плейсхолдер --- меняем его на число.
                    writeTextOrNumber(resultCell, buf,
                            context.containsPlaceholder(val));
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    // Обрабатываем формулу
                    val = sourceCell.getCellFormula();
                    val = FormulaModifier
                            .modifyFormula(
                                    val,
                                    resultCell.getColumnIndex()
                                            - sourceCell.getColumnIndex(),
                                    resultCell.getRowIndex()
                                            - sourceCell.getRowIndex());
                    resultCell.setCellFormula(val);
                    needEval = true;
                    break;
                // Остальные типы ячеек пока игнорируем
                default:
                    break;
                }
            }
        }

        // Разбираемся с merged-ячейками
        arrangeMergedCells(growthPoint, range);
    }

    private void writeTextOrNumber(Cell resultCell, String buf, boolean decide) {
        if (decide
                && !"@".equals(resultCell.getCellStyle().getDataFormatString())) {
            Matcher numberMatcher = NUMBER.matcher(buf.trim());
            Matcher dateMatcher = DATE.matcher(buf.trim());
            // может, число?
            if (numberMatcher.matches())
                resultCell.setCellValue(Double.parseDouble(buf));
            // может, дата?
            else if (dateMatcher.matches()) {
                Calendar c = Calendar.getInstance();
                c.clear();
                c.set(Integer.parseInt(dateMatcher.group(1)),
                        Integer.parseInt(dateMatcher.group(2)) - 1,
                        Integer.parseInt(dateMatcher.group(3)));
                resultCell.setCellValue(c.getTime());
            } else {
                resultCell.setCellValue(buf);
            }
        } else {
            resultCell.setCellValue(buf);
        }
    }

    private void arrangeMergedCells(CellAddress growthPoint, RangeAddress range)
            throws XML2SpreadSheetError {
        int mr = activeTemplateSheet.getNumMergedRegions();
        for (int i = 0; i < mr; i++) {
            // Диапазон смёрдженных ячеек на листе шаблона
            RangeAddress ra = new RangeAddress(activeTemplateSheet
                    .getMergedRegion(i).formatAsString());

            if (!(ra.top() >= range.top() && ra.bottom() <= range.bottom()
                    && ra.left() >= range.left() && ra.right() <= range.right()))
                continue;

            int ydiff = -range.top() + growthPoint.getRow() - 1;
            int firstRow = ra.top() + ydiff;
            int lastRow = ra.bottom() + ydiff;

            int xdiff = -range.left() + growthPoint.getCol() - 1;
            int firstCol = ra.left() + xdiff;
            int lastCol = ra.right() + xdiff;
            CellRangeAddress res = new CellRangeAddress(firstRow, lastRow,
                    firstCol, lastCol);

            activeResultSheet.addMergedRegion(res);
        }
    }

    abstract void evaluate();

    @Override
    void mergeUp(CellAddress a1, CellAddress a2) {
        CellRangeAddress res = new CellRangeAddress(a1.getRow() - 1,
                a2.getRow() - 1, a1.getCol() - 1, a2.getCol() - 1);
        activeResultSheet.addMergedRegion(res);
    }

    @Override
    void addNamedRegion(String name, CellAddress a1, CellAddress a2) {
        Name region = activeResultSheet.getWorkbook().getName(name);
        if (region == null)
            region = activeResultSheet.getWorkbook().createName();
        region.setNameName(name);
        // don't forget to replace single quote with double quotes!
        String formula = String.format("'%s'!%s:%s", activeResultSheet
                .getSheetName().replaceAll("'", "''"), a1.getAddress(), a2
                .getAddress());
        region.setRefersToFormula(formula);
    }

    @Override
    public void flush() throws XML2SpreadSheetError {
        if (needEval)
            evaluate();
        try {
            result.write(getOutput());
        } catch (IOException e) {
            throw new XML2SpreadSheetError(e.getMessage());
        }
    }

    @Override
    void putRowBreak(int rowNumber) {
        if (activeResultSheet != null && rowNumber >= 0)
            activeResultSheet.setRowBreak(rowNumber);
    }

    @Override
    void putColBreak(int colNumber) {
        if (activeResultSheet != null && colNumber >= 0)
            activeResultSheet.setColumnBreak(colNumber);
    }

    Workbook getResult() {
        return result;
    }
}
