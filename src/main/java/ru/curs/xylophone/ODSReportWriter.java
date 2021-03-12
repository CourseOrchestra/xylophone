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

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.InputStream;

//import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;

/**
 * Реализация ReportWriter для вывода в формат OpenOffice (ODS).
 */
final class ODSReportWriter extends ReportWriter {

    private final int rows = 3;
    private final int columns = 3;
    private final com.github.miachm.sods.Sheet sheet = new com.github.miachm.sods.Sheet("A", rows, columns);
    private SpreadSheet spread = new SpreadSheet();

    private InputStream template;
    private InputStream templateCopy;


    ODSReportWriter(InputStream template, InputStream templateCopy) {
        // TODO Auto-generated constructor stub
        this.template = template;
        this.templateCopy = templateCopy;
        try {
            sheet.getDataRange().setValues(1,2,3,4,5,6,7,8,9);

            // Set the underline style in the (3,3) cell
            sheet.getRange(2,2).setFontUnderline(true);

            // Set a bold font to the first 2x2 grid
            sheet.getRange(0,0,2,2).setFontBold(true);

            spread.appendSheet(sheet);
            spread.save(new File("Out.ods"));
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    @Override
    void newSheet(String sheetName, String sourceSheet,
            int startRepeatingColumn, int endRepeatingColumn,
            int startRepeatingRow, int endRepeatingRow) {
        // TODO Auto-generated method stub


    }

    @Override
    void putSection(XMLContext context, CellAddress growthPoint2,
            String sourceSheet, RangeAddress range) {
        // TODO Auto-generated method stub

    }

    @Override
    public void flush() throws XML2SpreadSheetError {
        // TODO Auto-generated method stub
        try {
            spread.save(getOutput());
        } catch (IOException e) {
            throw new XML2SpreadSheetError(e.getMessage());
        }
    }

    @Override
    void mergeUp(CellAddress a1, CellAddress a2) {
        // TODO Auto-generated method stub

    }

    @Override
    void addNamedRegion(String name, CellAddress a1, CellAddress a2) {
        // TODO Auto-generated method stub

    }

    @Override
    void putRowBreak(int rowNumber) {
        // TODO Auto-generated method stub

    }

    @Override
    void putColBreak(int colNumber) {
        // TODO Auto-generated method stub

    }

    @Override
    public Sheet getSheet() {
        throw new UnsupportedOperationException();
//        Sheet activeResultSheet;
//        Workbook result = createResultWb(templateCopy);
//        activeResultSheet = result.getSheet(sheetName);


        // cast from ODS sheet to POI sheet ???
//        return activeResultSheet;
    }
}
