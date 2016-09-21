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
 * Содержит метод обработки формул (сдвижки адресов ячеек).
 */
public final class FormulaModifier {

	private static final Pattern CELL_ADDRESS = Pattern
			.compile("[A-Z]+[0-9]+[(]?");

	private FormulaModifier() {
	}

	/**
	 * Модифицирует формулу, сдвигая адреса ячеек.
	 * 
	 * @param val
	 *            Текстовое значение формулы.
	 * @param dx
	 *            Сдвижка по колонкам.
	 * @param dy
	 *            Сдвижка по строкам.
	 */
	public static String modifyFormula(String val, int dx, int dy) {
		Matcher m = CELL_ADDRESS.matcher(val);
		StringBuffer sb = new StringBuffer();
		while (m.find()) {
			String buf = m.group(0);
			if (buf.endsWith("(")) {
				// Случай формулы (LOG10, ATAN2 и т. п.) обрабатываем отдельно
				m.appendReplacement(sb, buf);
			} else {
				CellAddress cellAddr = new CellAddress(buf);
				cellAddr.setCol(cellAddr.getCol() + dx);
				cellAddr.setRow(cellAddr.getRow() + dy);
				m.appendReplacement(sb, cellAddr.getAddress());
			}
		}
		m.appendTail(sb);
		return sb.toString();
	}
}
