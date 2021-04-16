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

import ru.curs.xylophone.descriptor.DescriptorElement;
import ru.curs.xylophone.descriptor.DescriptorOutput;

import java.io.InputStream;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Класс, ответственный за чтение из XML-файла и перенаправление команд на вывод
 * в объект ReportWriter.
 */
abstract class XMLDataReader {

    private static final Pattern XQUERY = Pattern
            .compile("([^\\[]+)\\[@([^=]+)=('([^']+)'|\"([^\"]+)\")]");

    private final ReportWriter writer;
    private final DescriptorElement descriptor;

    XMLDataReader(ReportWriter writer, DescriptorElement descriptor) {
        this.writer = writer;
        this.descriptor = descriptor;
    }

    /**
     * Создаёт объект-читальщик исходных данных на основании предоставленных
     * сведений.
     *
     * @param xmlData          Поток с исходными данными.
     * @param descriptorStream Дескриптор отчёта.
     * @param useSAX           Режим обработки (DOM или SAX).
     * @param writer           Объект, осуществляющий вывод.
     * @throws XylophoneError В случае ошибки обработки дескриптора отчёта.
     */
    static XMLDataReader createReader(
            InputStream xmlData,
            InputStream descriptorStream,
            boolean useSAX,
            ReportWriter writer)
            throws XylophoneError {
        if (xmlData == null)
            throw new XylophoneError("Data stream is null.");
        if (descriptorStream == null)
            throw new XylophoneError("Descriptor stream is null.");

        // Сначала парсится дескриптор и строится его объектное представление.
        DescriptorElement root;
        try {
            root = DescriptorElement.jsonDeserialize(descriptorStream);
        } catch (Exception e) {
            throw new XylophoneError(
                    "Error while processing json descriptor: " + e.getMessage());
        }
        // Затем инстанцируется конкретная реализация (DOM или SAX) ридера
        if (useSAX)
            return new SAXDataReader(xmlData, root, writer);
        else
            return new DOMDataReader(xmlData, root, writer);
    }

    /**
     * Осуществляет генерацию отчёта.
     *
     * @throws XylophoneError В случае возникновения ошибок ввода-вывода или при
     *                              интерпретации данных, шаблона или дескриптора.
     */
    abstract void process() throws XylophoneError;

    /**
     * Общий для DOM и SAX реализации метод обработки вывода.
     *
     * @param c Контекст.
     * @param o Дескриптор секции.
     * @throws XylophoneError В случае возникновения ошибок ввода-вывода или при
     *                              интерпретации шаблона.
     */
    final void processOutput(XMLContext c, DescriptorOutput o)
            throws XylophoneError {
        if (o.getWorksheet() != null) {
            String wsName = c.calc(o.getWorksheet());
            getWriter().sheet(wsName, o.getSourceSheet(),
                    o.getStartRepeatingColumn(), o.getEndRepeatingColumn(),
                    o.getStartRepeatingRow(), o.getEndRepeatingRow());
        }
        if (o.getRange() != null)
            getWriter().section(c, o.getSourceSheet(), o.getRange(), o.getPageBreak());

    }

    static boolean compareIndices(int expected, int actual) {
        return (expected < 0) || (actual == expected);
    }

    static boolean compareNames(String expected, String actual,
                                Map<String, String> attributes) {
        if (expected == null)
            return actual == null;

        if (expected.startsWith("*") || expected.equals(actual))
            return true;

        Matcher m = XQUERY.matcher(expected);
        if (m.matches()) {
            String tagName = m.group(1);
            if (!tagName.equals(actual))
                return false;
            String attribute = m.group(2);
            String value = m.group(4) == null ? m.group(5) : m.group(4);
            return value.equals(attributes.get(attribute));
        } else
            return false;
    }

    final ReportWriter getWriter() {
        return writer;
    }

    final DescriptorElement getDescriptor() {
        return descriptor;
    }

}
