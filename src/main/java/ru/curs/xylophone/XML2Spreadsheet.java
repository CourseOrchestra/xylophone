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

import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

/**
 * Основной класс построителя отчётов из XML данных в формате электронных
 * таблиц.
 */
public final class XML2Spreadsheet {

    private XML2Spreadsheet() {
    }

    /**
     * Запускает построение отчётов на исходных данных. Перегруженная версия
     * метода, работающая на потоках.
     *
     * @param xmlData          Исходные данные.
     * @param descriptorStream Дескриптор, описывающий порядок итерации по исходным данным.
     * @param template         Шаблон отчёта.
     * @param outputType       Тип шаблона отчёта (OpenOffice, XLS, XLSX).
     * @param useSAX           Режим процессинга (DOM или SAX).
     * @param copyTemplate     Копировать ли шаблон полностью перед началом обработки.
     * @param output           Поток, в который записывается результирующий отчёт.
     * @throws XylophoneError в случае возникновения ошибок
     */
    public static void process(
            InputStream xmlData,
            InputStream descriptorStream,
            InputStream template,
            OutputType outputType,
            boolean useSAX,
            boolean copyTemplate,
            OutputStream output)
            throws XylophoneError {
        ReportWriter writer = ReportWriter.createWriter(template, outputType,
                copyTemplate, output);
        XMLDataReader reader = XMLDataReader.createReader(
                xmlData,
                descriptorStream,
                useSAX,
                writer);
        reader.process();
    }

    /**
     * Запускает построение отчёта формата Excel (.XLS) на исходных данных и
     * возвращает объектное представление отчёта.
     *
     * @param xmlData       Исходные данные.
     * @param xmlDescriptor Дескриптор, описывающий порядок итерации по исходным данным.
     * @param template      Шаблон отчёта.
     * @param useSAX        Режим процессинга (DOM или SAX).
     * @param copyTemplate  Копировать ли шаблон полностью перед началом обработки.
     * @throws XylophoneError в случае возникновения ошибок
     * @throws IOException          Если файлы шаблона или дескриптора не найдены или произошла
     *                              иная ошибка ввода-вывода.
     */
    public static Workbook toPOIWorkbook(InputStream xmlData,
                                         FileInputStream xmlDescriptor, File template, boolean useSAX,
                                         boolean copyTemplate) throws XylophoneError, IOException {

        OutputType outputType = getOutputType(template);
        if (!(outputType == OutputType.XLS || outputType == OutputType.XLSX))
            throw new XylophoneError(
                    "toPOIWorkbook method works only for POI output types (XLS, XLSX).");

        try (InputStream descr = xmlDescriptor;
             InputStream templ = new FileInputStream(template)) {
            ReportWriter writer = ReportWriter.createWriter(templ, outputType,
                    copyTemplate, new OutputStream() {
                        @Override
                        public void write(int b) {
                            // Do nothing.
                        }
                    });
            XMLDataReader reader = XMLDataReader.createReader(
                    xmlData,
                    descr,
                    useSAX,
                    writer);
            reader.process();
            return ((POIReportWriter) writer).getResult();
        }

    }

    /**
     * Запускает построение отчётов на исходных данных. Перегруженная версия
     * метода, работающая на файлах (для удобства использования из
     * Python-скриптов).
     *
     * @param xmlData      Исходные данные.
     * @param descriptor   Дескриптор, описывающий порядок итерации по исходным данным.
     * @param template     Шаблон отчёта. Тип шаблона отчёта определяется по расширению.
     * @param useSAX       Режим процессинга (false, если DOM, или true, если SAX).
     * @param copyTemplate Копировать ли шаблон.
     * @param output       Поток, в который записывается результирующий отчёт.
     * @throws FileNotFoundException в случае, если указанные файлы не существуют
     * @throws XylophoneError  в случае иных ошибок
     */
    public static void process(InputStream xmlData, InputStream descriptor,
                               File template, boolean useSAX, boolean copyTemplate,
                               OutputStream output) throws FileNotFoundException, XylophoneError {
        OutputType outputType = getOutputType(template);
        try (
                InputStream descr = descriptor;
                InputStream templ = new FileInputStream(template)
        ) {
            process(xmlData, descr, templ, outputType, useSAX, copyTemplate,
                    output);
        } catch (FileNotFoundException e) {
            throw e;
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static OutputType getOutputType(File template)
            throws XylophoneError {
        String buf = template.toString();
        int dotInd = buf.lastIndexOf('.');
        buf = (dotInd > 0 && dotInd < buf.length()) ? buf.substring(dotInd + 1)
                : null;
        OutputType outputType;
        if ("ods".equalsIgnoreCase(buf)) {
            outputType = OutputType.ODS;
        } else if ("xls".equalsIgnoreCase(buf)) {
            outputType = OutputType.XLS;
        } else if ("xlsx".equalsIgnoreCase(buf)) {
            outputType = OutputType.XLSX;
        } else {
            throw new XylophoneError(
                    "Cannot define output format, template has non-standard extention.");
        }
        return outputType;
    }
}
