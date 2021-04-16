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

/**
 * Указывает на диапазон ячеек.
 */
public final class RangeAddress {

    private static final Pattern RANGE_ADDRESS = Pattern
            .compile("([A-Z]+[0-9]+)(:([A-Z]+[0-9]+))?");
    private final CellAddress topLeft;
    private final CellAddress bottomRight;

    public RangeAddress(String address) throws XylophoneError {
        Matcher m = RANGE_ADDRESS.matcher(address);
        if (!m.matches())
            throw new XylophoneError("Incorrect range: " + address);
        CellAddress c1 = new CellAddress(m.group(1));
        CellAddress c2 = new CellAddress(m.group(3) == null ? m.group(1)
                : m.group(3));
        topLeft = new CellAddress(Math.min(c1.getCol(), c2.getCol()), Math.min(
                c1.getRow(), c2.getRow()));
        bottomRight = new CellAddress(Math.max(c1.getCol(), c2.getCol()),
                Math.max(c1.getRow(), c2.getRow()));
    }

    public CellAddress topLeft() {
        return topLeft;
    }

    public CellAddress bottomRight() {
        return bottomRight;
    }

    public int left() {

        return topLeft.getCol();
    }

    public int right() {
        return bottomRight.getCol();
    }

    public int top() {
        return topLeft.getRow();
    }

    public int bottom() {
        return bottomRight.getRow();
    }

    public void setRight(int max) {
        bottomRight.setCol(max);
    }

    public void setBottom(int max) {
        bottomRight.setRow(max);

    }

    public String getAddress() {
        return topLeft.getAddress() + ":" + bottomRight.getAddress();
    }
}
