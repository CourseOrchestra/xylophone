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

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.Node;
import org.xml.sax.Attributes;

/**
 * Указывает на контекст XML файла, в котором могут быть вычислены
 * XPath-выражения.
 */
abstract class XMLContext {

    private static final Pattern P = Pattern.compile("~\\{([^}]+)}");
    private static final String CURRENT = "current";
    private static final String POSITION = "position";
    private static final Pattern FUNCTION = Pattern.compile("((" + CURRENT
            + ")|(" + POSITION + "))\\(\\)");

    boolean containsPlaceholder(String formatString) {
        return P.matcher(formatString).find();
    }

    /**
     * Вычисляет значение строки, содержащей, возможно, xpath-выражение.
     *
     * @param formatString
     *            строка, содержащая, возможно, xpath-выражения.
     */
    final String calc(String formatString) {
        Matcher m = P.matcher(formatString);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            String val = getXPathValue(m.group(1));
            m.appendReplacement(sb, val == null ? "" : val);
        }
        m.appendTail(sb);
        return sb.toString();
    }

    abstract String getXPathValue(String xpath);

    /**
     * Реализация XMLContext для DOM.
     */
    static final class DOMContext extends XMLContext {
        private final Node n;
        private final String path;
        private final int position;
        private XPath evaluator;

        DOMContext(Node n, String path, int position) {
            if (n == null)
                throw new NullPointerException();
            this.n = n;
            this.path = path;
            this.position = position;
        }

        @Override
        String getXPathValue(String xpath) {
            Matcher m = FUNCTION.matcher(xpath);
            StringBuffer sb = new StringBuffer();
            while (m.find()) {
                if (CURRENT.equals(m.group(1)))
                    m.appendReplacement(sb, path);
                else if (POSITION.equals(m.group(1)))
                    m.appendReplacement(sb, Integer.toString(position));
            }
            m.appendTail(sb);
            if (evaluator == null)
                evaluator = XPathFactory.newInstance().newXPath();
            try {
                XPathExpression expr = evaluator.compile(sb.toString());
                return (String) expr.evaluate(n, XPathConstants.STRING);
            } catch (XPathExpressionException e) {
                return "{" + e.getMessage() + "}";
            }
        }
    }

    /**
     * Реализация XMLContext для SAX.
     */
    static final class SAXContext extends XMLContext {

        private final Attributes attr;
        private final int position;

        SAXContext(Attributes attr, int position) {
            this.attr = attr;
            this.position = position;
        }

        @Override
        String getXPathValue(String xpath) {
            if (xpath.startsWith("@")) {
                return attr.getValue(xpath.substring(1));
            } else if (xpath.startsWith(POSITION)) {
                return Integer.toString(position);
            } else {
                return "{Only references to attributes or position() function in SAX mode are allowed}";
            }
        }
    }

}
