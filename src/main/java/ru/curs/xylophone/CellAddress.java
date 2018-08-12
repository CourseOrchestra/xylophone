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
 * Преобразует адрес ячейки в формате A1 в пару "строка, столбец" и обратно.
 *
 */
public final class CellAddress {

    private static final int RADIX = 'Z' - 'A' + 1;
    private static final Pattern CELL_ADDRESS = Pattern
            .compile("([A-Z]+)([0-9]+)");

    private int row;
    private int col;

    public CellAddress(String address) {
        setAddress(address);
    }

    public CellAddress(int col, int row) {
        this.col = col;
        this.row = row;
    }

    /**
     * Возвращает номер строки.
     */
    public int getRow() {
        return row;
    }

    /**
     * Возвращает номер колонки.
     */
    public int getCol() {
        return col;
    }

    /**
     * Устанавливает номер строки.
     *
     * @param row
     *            номер строки
     */
    public void setRow(int row) {
        this.row = row;
    }

    /**
     * Устанавливает номер колонки.
     *
     * @param col
     *            номер колонки
     *
     */
    public void setCol(int col) {
        this.col = col;
    }

    /**
     * Возвращает адрес.
     */
    public String getAddress() {
        int c = col;
        String sc = "";
        do {
            int digit = c % RADIX;
            char d;
            if (digit == 0) {
                c -= RADIX;
                d = 'Z';
            } else {
                d = (char) (digit + 'A' - 1);
            }

            sc = d + sc;

            c /= RADIX;

        } while (c > 0);
        return sc + String.valueOf(row);
    }

    /**
     * Устанавливает адрес.
     *
     * @param address
     *            адрес
     */
    public void setAddress(String address) {
        Matcher m = CELL_ADDRESS.matcher(address);
        m.matches();
        row = Integer.parseInt(m.group(2));

        col = 0;
        String a = m.group(1);
        for (int i = 0; i < a.length(); i++) {
            col = col * RADIX;
            char c = a.charAt(i);
            int d = c - 'A' + 1;
            col += d;
        }
    }

    @Override
    public int hashCode() {
        return row * col + row;
    }

    @Override
    public boolean equals(Object obj) {
        if (obj instanceof CellAddress) {
            return equals((CellAddress) obj);
        } else {
            return super.equals(obj);
        }
    }

    /**
     * Метод equals специально для сравнения адресов ячеек.
     *
     * @param a
     *            другой адрес ячейки
     */
    public boolean equals(CellAddress a) {
        return row == a.row && col == a.col;
    }
}
