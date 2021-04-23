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

import org.apache.commons.io.FilenameUtils;
import ru.curs.xylophone.descriptor.DescriptorElement;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Класс для запуска из командной строки.
 */
public class Main {
    private static final String DATA = "-data";
    private static final String TEMPLATE = "-template";
    private static final String DESCR = "-descr";
    private static final String OUT = "-out";
    private static final String SAX = "-sax";
    private static final String COPYTEMPLATE = "-copytemplate";
    private static final String CONVERT_DESCRIPTOR = "-dconv"; // converts XML descriptor into JSON one

    private enum State {
        READTOKEN, READDATA, READTEMPLATE, READDESCR, READOUT, WRITEDESCR,
    }

    /**
     * Главный метод класса.
     *
     * @param args аргументы
     * @throws XylophoneError        в случае, если произошла ошибка конвертации
     * @throws FileNotFoundException в случае, если файл не найден
     */
    public static void main(String[] args) throws FileNotFoundException,
            XylophoneError {

        FileInputStream iff = null;
        String descrFileName = null;     // for descriptor converter
        String outDescrFileName = null;     // for descriptor converter
        File template = null;
        FileOutputStream output = null;
        boolean useSAX = false;
        boolean copyTemplate = false;

        boolean convDescriptor = false;

        if (args.length == 0) {
            showHelp();
            return;
        }

        State state = State.READTOKEN;

        for (String s : args)
            switch (state) {
                case READTOKEN:
                    if (DATA.equalsIgnoreCase(s))
                        state = State.READDATA;
                    else if (TEMPLATE.equalsIgnoreCase(s))
                        state = State.READTEMPLATE;
                    else if (DESCR.equalsIgnoreCase(s))
                        state = State.READDESCR;
                    else if (OUT.equalsIgnoreCase(s))
                        state = State.READOUT;
                    else if (SAX.equalsIgnoreCase(s))
                        useSAX = true;
                    else if (COPYTEMPLATE.equalsIgnoreCase(s))
                        copyTemplate = true;
                    else if (CONVERT_DESCRIPTOR.equalsIgnoreCase(s))
                        state = State.WRITEDESCR;
                    else {
                        showHelp();
                        return;
                    }
                    break;
                case READDATA:
                    iff = new FileInputStream(s);
                    state = State.READTOKEN;
                    break;
                case READTEMPLATE:
                    template = new File(s);
                    state = State.READTOKEN;
                    break;
                case READDESCR:
                    descrFileName = s;
                    state = State.READTOKEN;
                    break;
                case READOUT:
                    output = new FileOutputStream(s);
                    state = State.READTOKEN;
                    break;
                case WRITEDESCR:
                    convDescriptor = true;
                    outDescrFileName = s;
                    state = State.READTOKEN;
                    break;
                default:
                    break;
            }

        if (convDescriptor && descrFileName != null && outDescrFileName != null) {
            convertDescriptor(descrFileName, outDescrFileName);
            return;
        } else if (iff != null && descrFileName != null && template != null && output != null)
            XML2Spreadsheet.process(iff, new FileInputStream(descrFileName), template, useSAX, copyTemplate, output);
        else {
            showHelp();
            return;
        }

        System.out.println("Spreadsheet created successfully.");
    }


    private static void convertDescriptor(String inName, String outName) throws XylophoneError {
        String inExt = FilenameUtils.getExtension(inName).toLowerCase();
        String outExt = FilenameUtils.getExtension(outName).toLowerCase();
        DescriptorElement root;
        try {
            FileInputStream inFileStream = new FileInputStream(inName);
            if (inExt.equals("xml")) {
                root = XMLDescriptorParser.readXMLDescriptor(inFileStream);
            } else
                throw new XylophoneError("Invalid file name extension for " +
                        "input descriptor file: \"" + inName + "\"");

            FileOutputStream outFileStream = new FileOutputStream(outName);
            if (outExt.equals("json")) {
                root.jsonSerialize(outFileStream);
            } else
                throw new XylophoneError("Invalid file name extension for " +
                        "output descriptor file: \"" + inName + "\"");

            FileOutputStream outDescriptor = new FileOutputStream(outName);
            root.jsonSerialize(outDescriptor);
        } catch (Exception e) {
            throw new XylophoneError("Error while converting XML to JSON descriptor: " + e.getMessage());
        }
        System.out.println(inExt.toUpperCase() + " descriptor was successfully converted to " +
                outExt.toUpperCase() + " format.");
    }

    private static void showHelp() {
        System.out.println();
        System.out.println("  Xylophone should be called with the following parameters (any order):");
        System.out.println("    " + DATA + " XML data file");
        System.out.println("    " + TEMPLATE + " XLS/XLSX template file");
        System.out.println("    " + DESCR + " descriptor file");
        System.out.println("    [" + SAX
                + "] use SAX engine (instead of DOM) to parse data file");
        System.out.println("    [" + COPYTEMPLATE
                + "] copy the template file to output before processing");
        System.out.println("    [" + CONVERT_DESCRIPTOR + " in_descriptor_file] converts descriptor from ");
        System.out.println("      XML type into JSON type");
    }
}
