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

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import ru.curs.xylophone.descriptor.DescriptorElement;
import ru.curs.xylophone.descriptor.DescriptorIteration;
import ru.curs.xylophone.descriptor.DescriptorOutput;
import ru.curs.xylophone.descriptor.DescriptorOutputBase;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.InputStream;
import java.util.HashMap;

/**
 * Класс, ответственный за чтение из XML-файла и перенаправление команд на вывод
 * в объект ReportWriter методом DOM.
 */
final class DOMDataReader extends XMLDataReader {

    private final Document xmlData;

    DOMDataReader(InputStream xmlData, DescriptorElement xmlDescriptor,
            ReportWriter writer) throws XylophoneError {
        super(writer, xmlDescriptor);
        try {
            DocumentBuilder db = DocumentBuilderFactory.newInstance()
                    .newDocumentBuilder();
            this.xmlData = db.parse(xmlData);
        } catch (Exception e) {
            throw new XylophoneError("Error while parsing input data: "
                    + e.getMessage());
        }

    }

    // В режиме итерации нашёлся подходящий элемент
    private void processElement(String elementPath, DescriptorElement de,
            Element xe, int position) throws XylophoneError {
        XMLContext context = null;
        for (DescriptorOutputBase se : de.getSubElements()) {
            if (se instanceof DescriptorIteration) {
                processIteration(elementPath, xe, (DescriptorIteration) se,
                        position);
            } else if (se instanceof DescriptorOutput) {
                // Контекст имеет смысл создавать лишь если есть хоть один
                // output
                if (context == null)
                    context = new XMLContext.DOMContext(xe, elementPath,
                            position);
                processOutput(context, (DescriptorOutput) se);
            }
        }

    }

    // По субэлементам текущего элемента надо провести итерацию
    private void processIteration(String elementPath, Element parent,
            DescriptorIteration i, int position) throws XylophoneError {

        final HashMap<String, Integer> elementIndices = new HashMap<>();

        getWriter().startSequence(i.isHorizontal());

        for (DescriptorElement de : i.getElements())
            if ("(before)".equals(de.getName()))
                processElement(elementPath, de, parent, position);

        Node n = parent.getFirstChild();
        int elementIndex = -1;

        int pos = 0;
        iteration:
        while (n != null) {
            // Нас интересуют только элементы
            if (n.getNodeType() == Node.ELEMENT_NODE) {
                // Поддерживаем таблицу с нумерацией нод для вычисления пути
                Integer ind = elementIndices.get(n.getNodeName());
                if (ind == null)
                    ind = 0;
                elementIndices.put(n.getNodeName(), ind + 1);

                elementIndex++;
                boolean found = false;
                if (compareIndices(i.getIndex(), elementIndex)) {
                    HashMap<String, String> atts = new HashMap<>();

                    for (int j = 0; j < n.getAttributes().getLength(); j++) {
                        Node att = n.getAttributes().item(j);
                        atts.put(att.getNodeName(), att.getNodeValue());
                    }

                    for (DescriptorElement e : i.getElements())
                        if (compareNames(e.getName(), n.getNodeName(),
                                atts)) {
                            found = true;
                            processElement(String.format("%s/%s[%s]",
                                    elementPath, n.getNodeName(),
                                    elementIndices.get(n.getNodeName())
                                            .toString()), e, (Element) n,
                                    pos + 1);
                        }
                    // Если явно задан индекс, то на этом заканчиваем итерацию
                    if (i.getIndex() >= 0)
                        break iteration;
                }
                if (found)
                    pos++;
            }
            n = n.getNextSibling();
        }

        for (DescriptorElement de : i.getElements())
            if ("(after)".equals(de.getName()))
                processElement(elementPath, de, parent, position);

        getWriter().endSequence(i.getMerge(), i.getRegionName());
    }


    @Override
    void process() throws XylophoneError {
        // Обработка в DOM-режиме --- рекурсивная, управляемая дескриптором.
        if (getDescriptor().getName().equals(
                xmlData.getDocumentElement().getNodeName())) {
            processElement("/" + getDescriptor().getName() + "[1]",
                    getDescriptor(), xmlData.getDocumentElement(), 1);
        }
        getWriter().applyMergedRegions(MergeRegionContainer.getContainer().getMergedRegions());
        MergeRegionContainer.getContainer().clear();

        getWriter().flush();
    }
}
