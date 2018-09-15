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

import java.io.InputStream;
import java.util.Deque;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.transform.TransformerFactory;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamSource;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * Класс, ответственный за чтение из XML-файла и перенаправление команд на вывод
 * в объект ReportWriter.
 */
abstract class XMLDataReader {

    private static final Pattern RANGE = Pattern
            .compile("(-?[0-9]+):(-?[0-9]+)");
    private static final Pattern XQUERY = Pattern
            .compile("([^\\[]+)\\[@([^=]+)=('([^']+)'|\"([^\"]+)\")]");

    private final ReportWriter writer;
    private final DescriptorElement descriptor;

    XMLDataReader(ReportWriter writer, DescriptorElement descriptor) {
        this.writer = writer;
        this.descriptor = descriptor;
    }

    private enum ParserState {
        ELEMENT, ITERATION, OUTPUT
    }

    private static final class DescriptorParser extends DefaultHandler {

        private final Deque<DescriptorElement> elementsStack = new LinkedList<>();
        private DescriptorElement root;
        private ParserState parserState = ParserState.ITERATION;

        @Override
        public void startElement(String uri, String localName, String name,
                final Attributes atts) throws SAXException {

            abstract class AttrReader<T> {
                T getValue(String qName) throws XML2SpreadSheetError {
                    String buf = atts.getValue(qName);
                    if (buf == null || "".equals(buf))
                        return getIfEmpty();
                    else
                        return getIfNotEmpty(buf);
                }

                abstract T getIfNotEmpty(String value)
                        throws XML2SpreadSheetError;

                abstract T getIfEmpty();
            }

            final class StringAttrReader extends AttrReader<String> {
                @Override
                String getIfNotEmpty(String value) {
                    return value;
                }

                @Override
                String getIfEmpty() {
                    return null;
                }
            }

            try {
                switch (parserState) {
                case ELEMENT:
                    if ("iteration".equals(localName)) {
                        int index = (new AttrReader<Integer>() {
                            @Override
                            Integer getIfNotEmpty(String value) {
                                return Integer.parseInt(value);
                            }

                            @Override
                            Integer getIfEmpty() {
                                return -1;
                            }
                        }).getValue("index");

                        int merge = (new AttrReader<Integer>() {
                            @Override
                            Integer getIfNotEmpty(String value) {
                                return Integer.parseInt(value);
                            }

                            @Override
                            Integer getIfEmpty() {
                                return 0;
                            }
                        }).getValue("merge");

                        boolean horizontal = (new AttrReader<Boolean>() {
                            @Override
                            Boolean getIfNotEmpty(String value) {
                                return "horizontal".equalsIgnoreCase(value);
                            }

                            @Override
                            Boolean getIfEmpty() {
                                return false;
                            }
                        }).getValue("mode");
                        String regionName = new StringAttrReader()
                                .getValue("regionName");
                        DescriptorIteration currIteration = new DescriptorIteration(
                                index, horizontal, merge, regionName);
                        elementsStack.peek().getSubelements()
                                .add(currIteration);
                        parserState = ParserState.ITERATION;
                    } else if ("output".equals(localName)) {
                        RangeAddress range = (new AttrReader<RangeAddress>() {
                            @Override
                            RangeAddress getIfNotEmpty(String value)
                                    throws XML2SpreadSheetError {
                                return new RangeAddress(value);
                            }

                            @Override
                            RangeAddress getIfEmpty() {
                                return null;
                            }
                        }).getValue("range");
                        StringAttrReader sar = new StringAttrReader();

                        boolean pagebreak = (new AttrReader<Boolean>() {
                            @Override
                            Boolean getIfNotEmpty(String value) {
                                return "true".equalsIgnoreCase(value);
                            }

                            @Override
                            Boolean getIfEmpty() {
                                return false;
                            }
                        }).getValue("pagebreak");

                        DescriptorOutput output = new DescriptorOutput(
                                sar.getValue("worksheet"), range,
                                sar.getValue("sourcesheet"),
                                sar.getValue("repeatingcols"),
                                sar.getValue("repeatingrows"), pagebreak);
                        elementsStack.peek().getSubelements().add(output);

                        parserState = ParserState.OUTPUT;
                    }
                    break;
                case ITERATION:
                    if ("element".equals(localName)) {
                        String elementName = (new StringAttrReader())
                                .getValue("name");
                        DescriptorElement currElement = new DescriptorElement(
                                elementName);

                        if (root == null)
                            root = currElement;
                        else {
                            // Добываем контекст текущей итерации...
                            List<DescriptorSubelement> subelements = elementsStack
                                    .peek().getSubelements();
                            DescriptorIteration iter = (DescriptorIteration) subelements
                                    .get(subelements.size() - 1);
                            iter.getElements().add(currElement);
                        }
                        elementsStack.push(currElement);
                        parserState = ParserState.ELEMENT;
                    }
                    break;
                }
            } catch (XML2SpreadSheetError e) {
                throw new SAXException(e.getMessage());
            }
        }

        @Override
        public void endElement(String uri, String localName, String name) {
            switch (parserState) {
            case ELEMENT:
                elementsStack.pop();
                parserState = ParserState.ITERATION;
                break;
            case ITERATION:
                parserState = ParserState.ELEMENT;
                break;
            case OUTPUT:
                parserState = ParserState.ELEMENT;
                break;
            }
        }

    }

    /**
     * Создаёт объект-читальщик исходных данных на основании предоставленных
     * сведений.
     *
     * @param xmlData
     *            Поток с исходными данными.
     * @param xmlDescriptor
     *            Дескриптор отчёта.
     * @param useSAX
     *            Режим обработки (DOM или SAX).
     * @param writer
     *            Объект, осуществляющий вывод.
     * @throws XML2SpreadSheetError
     *             В случае ошибки обработки дескриптора отчёта.
     */
    static XMLDataReader createReader(InputStream xmlData,
            InputStream xmlDescriptor, boolean useSAX, ReportWriter writer)
            throws XML2SpreadSheetError {
        if (xmlData == null)
            throw new XML2SpreadSheetError("XML Data is null.");
        if (xmlDescriptor == null)
            throw new XML2SpreadSheetError("XML descriptor is null.");

        // Сначала парсится дескриптор и строится его объектное представление.
        DescriptorParser parser = new DescriptorParser();
        try {
            TransformerFactory
                    .newInstance()
                    .newTransformer()
                    .transform(new StreamSource(xmlDescriptor),
                            new SAXResult(parser));
        } catch (Exception e) {
            throw new XML2SpreadSheetError(
                    "Error while processing XML descriptor: " + e.getMessage());
        }
        // Затем инстанцируется конкретная реализация (DOM или SAX) ридера
        if (useSAX)
            return new SAXDataReader(xmlData, parser.root, writer);
        else
            return new DOMDataReader(xmlData, parser.root, writer);
    }

    /**
     * Осуществляет генерацию отчёта.
     *
     * @throws XML2SpreadSheetError
     *             В случае возникновения ошибок ввода-вывода или при
     *             интерпретации данных, шаблона или дескриптора.
     */
    abstract void process() throws XML2SpreadSheetError;

    /**
     * Общий для DOM и SAX реализации метод обработки вывода.
     *
     * @param c
     *            Контекст.
     * @param o
     *            Дескриптор секции.
     * @throws XML2SpreadSheetError
     *             В случае возникновения ошибок ввода-вывода или при
     *             интерпретации шаблона.
     */
    final void processOutput(XMLContext c, DescriptorOutput o)
            throws XML2SpreadSheetError {
        if (o.getWorksheet() != null) {
            String wsName = c.calc(o.getWorksheet());
            getWriter().sheet(wsName, o.getSourceSheet(),
                    o.getStartRepeatingColumn(), o.getEndRepeatingColumn(),
                    o.getStartRepeatingRow(), o.getEndRepeatingRow());
        }
        if (o.getRange() != null)
            getWriter().section(c, o.getSourceSheet(), o.getRange(), o.getPageBreak());
    }

    static final boolean compareIndices(int expected, int actual) {
        return (expected < 0) || (actual == expected);
    }

    static final boolean compareNames(String expected, String actual,
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

    static final class DescriptorElement {
        private final String elementName;
        private final List<DescriptorSubelement> subelements = new LinkedList<>();

        DescriptorElement(String elementName) {
            this.elementName = elementName;
        }

        String getElementName() {
            return elementName;
        }

        List<DescriptorSubelement> getSubelements() {
            return subelements;
        }
    }

    abstract static class DescriptorSubelement {
    }

    static final class DescriptorIteration extends DescriptorSubelement {
        private final int index;
        private final int merge;
        private final boolean horizontal;
        private final String regionName;
        private final List<DescriptorElement> elements = new LinkedList<>();

        DescriptorIteration(int index, boolean horizontal, int merge,
                String regionName) {
            this.index = index;
            this.horizontal = horizontal;
            this.merge = merge;
            this.regionName = regionName;
        }

        int getIndex() {
            return index;
        }

        boolean isHorizontal() {
            return horizontal;
        }

        List<DescriptorElement> getElements() {
            return elements;
        }

        public int getMerge() {
            return merge;
        }

        public String getRegionName() {
            return regionName;
        }
    }

    static final class DescriptorOutput extends DescriptorSubelement {
        private final String worksheet;
        private final RangeAddress range;
        private final String sourceSheet;
        private final int startRepeatingColumn;
        private final int endRepeatingColumn;
        private final int startRepeatingRow;
        private final int endRepeatingRow;
        private final boolean pageBreak;

        DescriptorOutput(String worksheet, RangeAddress range,
                String sourceSheet, String repeatingCols, String repeatingRows,
                boolean pageBreak) throws XML2SpreadSheetError {
            this.worksheet = worksheet;
            this.range = range;
            this.sourceSheet = sourceSheet;
            this.pageBreak = pageBreak;
            Matcher m1 = RANGE.matcher(repeatingCols == null ? "-1:-1"
                    : repeatingCols);
            Matcher m2 = RANGE.matcher(repeatingRows == null ? "-1:-1"
                    : repeatingRows);
            if (m1.matches() && m2.matches()) {
                this.startRepeatingColumn = Integer.parseInt(m1.group(1));
                this.endRepeatingColumn = Integer.parseInt(m1.group(2));
                this.startRepeatingRow = Integer.parseInt(m2.group(1));
                this.endRepeatingRow = Integer.parseInt(m2.group(2));
            } else {
                throw new XML2SpreadSheetError(String.format(
                        "Invalid col/row range %s %s", repeatingCols,
                        repeatingRows));
            }

        }

        String getWorksheet() {
            return worksheet;
        }

        String getSourceSheet() {
            return sourceSheet;
        }

        RangeAddress getRange() {
            return range;
        }

        public int getStartRepeatingColumn() {
            return startRepeatingColumn;
        }

        public int getEndRepeatingColumn() {
            return endRepeatingColumn;
        }

        public int getStartRepeatingRow() {
            return startRepeatingRow;
        }

        public int getEndRepeatingRow() {
            return endRepeatingRow;
        }

        public boolean getPageBreak() {
            return pageBreak;
        }
    }

}
