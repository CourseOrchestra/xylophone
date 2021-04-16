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

import org.apache.poi.ss.util.CellRangeAddress;

import java.io.InputStream;
import java.util.stream.Stream;

/**
 * Реализация ReportWriter для вывода в формат OpenOffice (ODS).
 */
final class ODSReportWriter extends ReportWriter {

    ODSReportWriter(InputStream template, InputStream templateCopy) throws XylophoneError {
        throw new XylophoneError("In work ...");
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
    public void flush() {
        // TODO Auto-generated method stub

    }

    @Override
    void mergeUp(CellAddress a1, CellAddress a2) {
        // TODO Auto-generated method stub

    }

    // пока не понял что эта функция делает
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
    void applyMergedRegions(Stream<CellRangeAddress> mergedRegions){
        // TODO Auto-generated method stub

    }

}
