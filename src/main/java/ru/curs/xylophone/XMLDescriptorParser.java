/*
 * XMLDescriptor parser is left intentionally for casual conversions of old xml descriptors to
 * new JSON descriptors files by means parsing XML and JSON serializing descriptor classes
 */

package ru.curs.xylophone;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;
import ru.curs.xylophone.descriptor.DescriptorElement;
import ru.curs.xylophone.descriptor.DescriptorIteration;
import ru.curs.xylophone.descriptor.DescriptorOutput;
import ru.curs.xylophone.descriptor.DescriptorOutputBase;

import javax.xml.transform.TransformerFactory;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamSource;
import java.io.InputStream;
import java.util.Deque;
import java.util.LinkedList;
import java.util.List;

enum ParserState {
    ELEMENT, ITERATION, OUTPUT
}

class XMLDescriptorParser extends DefaultHandler {

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
                        elementsStack.peek().getSubElements()
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
                        elementsStack.peek().getSubElements().add(output);

                        parserState = ParserState.OUTPUT;
                    } else {
                        throw new XML2SpreadSheetError(String.format("Tag <element> is not allowed inside <element>. "
                                + "Error inside element with name %s.", elementsStack.peek().getName()));
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
                            List<DescriptorOutputBase> subelements = elementsStack
                                    .peek().getSubElements();
                            DescriptorIteration iter = (DescriptorIteration) subelements
                                    .get(subelements.size() - 1);
                            iter.getElements().add(currElement);
                        }
                        elementsStack.push(currElement);
                        parserState = ParserState.ELEMENT;
                    } else {
                        throw new XML2SpreadSheetError(
                                String.format("Tag <%s> is not allowed inside <iteration>. "
                                                + "Error inside element with name %s.",
                                        localName, elementsStack.peek().getName()));
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

    public static DescriptorElement readXMLDescriptor(InputStream descriptorStream) throws XML2SpreadSheetError {
        XMLDescriptorParser parser = new XMLDescriptorParser();
        try {
            TransformerFactory
                    .newInstance()
                    .newTransformer()
                    .transform(new StreamSource(descriptorStream),
                            new SAXResult(parser));
        } catch (Exception e) {
            throw new XML2SpreadSheetError(
                    "Error while processing XML descriptor: " + e.getMessage());
        }
        return parser.root;
    }
}
