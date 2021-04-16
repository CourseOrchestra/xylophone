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

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;
import ru.curs.xylophone.XMLContext.SAXContext;
import ru.curs.xylophone.descriptor.DescriptorElement;
import ru.curs.xylophone.descriptor.DescriptorIteration;
import ru.curs.xylophone.descriptor.DescriptorOutput;
import ru.curs.xylophone.descriptor.DescriptorOutputBase;

import javax.xml.transform.Source;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamSource;
import java.io.InputStream;
import java.util.Deque;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;

/**
 * Класс, ответственный за чтение из XML-файла и перенаправление команд на вывод
 * в объект ReportWriter методом SAX.
 */
final class SAXDataReader extends XMLDataReader {

    private final Source xmlData;

    SAXDataReader(
            InputStream xmlData,
            DescriptorElement xmlDescriptor,
            ReportWriter writer) {
        super(writer, xmlDescriptor);
        this.xmlData = new StreamSource(xmlData);

    }

    /**
     * Адаптирует дескриптор элемента к SAX-парсингу.
     */
    private static final class SAXElementDescriptor {
        private int elementIndex = -1;
        private int position = 0;
        private final int desiredIndex;
        private final XMLContext context;
        private final boolean iterate;
        private final boolean horizontal;
        private final List<DescriptorOutput> preOutputs = new LinkedList<>();
        private final List<DescriptorElement> expectedElements = new LinkedList<>();
        private final List<DescriptorOutput> postOutputs = new LinkedList<>();
        private final List<DescriptorOutput> headerOutputs = new LinkedList<>();
        private final List<DescriptorOutput> footerOutputs = new LinkedList<>();
        private final int merge;
        private final String regionName;

        SAXElementDescriptor() {
            context = null;
            iterate = false;
            horizontal = false;
            desiredIndex = -1;
            merge = 0;
            regionName = null;
        }

        SAXElementDescriptor(DescriptorElement e, XMLContext context)
                throws SAXException {
            this.context = context;
            boolean iterate = false;
            boolean horizontal = false;
            int desiredIndex = -1;
            int merge = 0;
            String regionName = null;
            for (DescriptorOutputBase se : e.getSubElements())
                if (!iterate) {
                    // До тэга iteration
                    if (se instanceof DescriptorOutput)
                        preOutputs.add((DescriptorOutput) se);
                    else if (se instanceof DescriptorIteration) {
                        for (DescriptorElement de : ((DescriptorIteration) se)
                                .getElements()) {
                            if ("(before)".equals(de.getName())) {
                                for (DescriptorOutputBase se2 : de
                                        .getSubElements())
                                    if (se2 instanceof DescriptorOutput)
                                        headerOutputs
                                                .add((DescriptorOutput) se2);
                            } else if ("(after)".equals(de.getName())) {
                                for (DescriptorOutputBase se2 : de
                                        .getSubElements())
                                    if (se2 instanceof DescriptorOutput)
                                        footerOutputs
                                                .add((DescriptorOutput) se2);
                            } else
                                expectedElements.add(de);
                        }
                        desiredIndex = ((DescriptorIteration) se).getIndex();
                        iterate = true;
                        horizontal = ((DescriptorIteration) se).isHorizontal();
                        merge = ((DescriptorIteration) se).getMerge();
                        regionName = ((DescriptorIteration) se).getRegionName();
                    }
                } else {
                    // После тэга iteration
                    if (se instanceof DescriptorOutput)
                        postOutputs.add((DescriptorOutput) se);
                    else if (se instanceof DescriptorIteration)
                        throw new SAXException(
                                "For SAX mode only one iteration element is allowed for each element descriptor.");
                }
            this.iterate = iterate;
            this.horizontal = horizontal;
            this.desiredIndex = desiredIndex;
            this.merge = merge;
            this.regionName = regionName;
        }
    }

    @Override
    void process() throws XylophoneError {

        final class Parser extends DefaultHandler {
            private final Deque<SAXElementDescriptor> elementsStack = new LinkedList<>();

            private void bypass() {
                elementsStack.push(new SAXElementDescriptor());
            }

            @Override
            public void startElement(String uri, String localName, String name,
                                     Attributes atts) throws SAXException {
                SAXElementDescriptor curDescr = elementsStack.peek();
                curDescr.elementIndex++;
                if (compareIndices(curDescr.desiredIndex, curDescr.elementIndex)) {
                    boolean found = false;
                    HashMap<String, String> attsmap = new HashMap<>();
                    for (int i = 0; i < atts.getLength(); i++)
                        attsmap.put(atts.getLocalName(i), atts.getValue(i));

                    searchElements:
                    for (DescriptorElement e : curDescr.expectedElements) {
                        if (compareNames(e.getName(), localName, attsmap)) {

                            XMLContext context = new SAXContext(atts,
                                    curDescr.position + 1);
                            SAXElementDescriptor sed = new SAXElementDescriptor(
                                    e, context);
                            elementsStack.push(sed);

                            // По пред-выводам выполняем вывод.
                            for (DescriptorOutput o : sed.preOutputs)
                                try {
                                    processOutput(sed.context, o);
                                } catch (XylophoneError e1) {
                                    throw new SAXException(e1.getMessage());
                                }
                            // Начинаем обрамление итерации
                            try {
                                if (sed.iterate) {
                                    getWriter().startSequence(sed.horizontal);
                                    for (DescriptorOutput deo : sed.headerOutputs)
                                        processOutput(sed.context, deo);
                                }
                            } catch (XylophoneError e1) {
                                throw new SAXException(e1.getMessage());
                            }
                            found = true;
                            break searchElements;
                        }
                    }
                    if (found)
                        curDescr.position++;
                    else
                        bypass();
                } else
                    bypass();
            }

            @Override
            public void endElement(String uri, String localName, String name)
                    throws SAXException {
                SAXElementDescriptor sed = elementsStack.pop();
                try {
                    // Завершаем обрамление итерации
                    if (sed.iterate) {
                        for (DescriptorOutput deo : sed.footerOutputs)
                            processOutput(sed.context, deo);
                        getWriter().endSequence(sed.merge, sed.regionName);
                    }
                    // По пост-выводам выполняем вывод
                    for (DescriptorOutput o : sed.postOutputs)
                        processOutput(sed.context, o);
                } catch (XylophoneError e1) {
                    throw new SAXException(e1.getMessage());
                }
            }
        }

        Parser parser = new Parser();
        SAXElementDescriptor sed = new SAXElementDescriptor();
        sed.expectedElements.add(getDescriptor());
        parser.elementsStack.push(sed);

        try {
            TransformerFactory.newInstance().newTransformer()
                    .transform(xmlData, new SAXResult(parser));
        } catch (Exception e) {
            throw new XylophoneError("Error while processing XML data: "
                    + e.getMessage());

        }
        getWriter().flush();
    }
}
