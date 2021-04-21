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
import java.util.stream.Stream;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Реализация ReportWriter для вывода в формат MSOffice (XLSX).
 */
final class XLSXReportWriter extends POIReportWriter {

    private XSSFWorkbook wb;

    XLSXReportWriter(InputStream template, InputStream templateCopy)
            throws XylophoneError {
        super(template, templateCopy);
    }

    void setResultWbMetadata() {
        POIXMLProperties props = wb.getProperties();

        POIXMLProperties.CoreProperties coreProp = props.getCoreProperties();
        coreProp.setCreator("");
        coreProp.setDescription("");
        coreProp.setKeywords("");
        coreProp.setTitle("");

        POIXMLProperties.ExtendedProperties extProp = props.getExtendedProperties();
        extProp.getUnderlyingProperties().setCompany("");

        POIXMLProperties.CustomProperties custProp = props.getCustomProperties();
        custProp.addProperty("Author", "");
        custProp.addProperty("Program name", "");
    }

    @Override
    Workbook createResultWb(InputStream templateCopy)
            throws InvalidFormatException, IOException {
        if (templateCopy == null) {
            wb = new XSSFWorkbook();
        } else {
            wb = (XSSFWorkbook) WorkbookFactory.create(templateCopy);
        }
        setResultWbMetadata();
        return wb;
    }

    @Override
    void evaluate() {
        XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
    }

    @Override
    void applyMergedRegions(Stream<CellRangeAddress> mergedRegions){
        mergedRegions.forEach(getSheet()::addMergedRegion);
    }
}
