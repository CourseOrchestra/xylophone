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

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Класс для работы с бинарными данными.
 */
public final class XML2SpreadseetBLOB {
	private DataPage data;
	private boolean isModified;
	private int size;

	/**
	 * Пустой (NULL) BLOB.
	 */
	public XML2SpreadseetBLOB() {
	}

	/**
	 * BLOB на основе данных потока.
	 *
	 * @param source
	 *            Поток, из которого данные прочитываются в BLOB.
	 * @throws IOException
	 *             При ошибке чтения.
	 */
	XML2SpreadseetBLOB(final InputStream source) throws IOException {
		InputStream counter = new InputStream() {
			@Override
			public int read() throws IOException {
				int result = source.read();
				if (result >= 0)
					size++;
				return result;
			}
		};
		int buf = counter.read();
		data = buf < 0 ? new DataPage(0) : new DataPage(buf, counter);
	}

	/**
	 * Клон-BLOB, указывающий на ту же самую страницу данных.
	 */
	public XML2SpreadseetBLOB clone() {
		XML2SpreadseetBLOB result = new XML2SpreadseetBLOB();
		result.data = data;
		result.size = size;
		return result;
	}

	/**
	 * Были ли данные BLOB-а изменены.
	 */
	public boolean isModified() {
		return isModified;
	}

	/**
	 * Возвращает поток для чтения данных.
	 */
	public InputStream getInStream() {
		return data == null ? null : data.getInStream();
	}

	/**
	 * Возвращает поток для записи данных, сбросив при этом текущие данные
	 * BLOB-а.
	 */
	public OutputStream getOutStream() {
		isModified = true;
		data = new DataPage();
		size = 0;
		return new OutputStream() {
			private DataPage tail = data;

			@Override
			public void write(int b) {
				tail = tail.write(b);
				size++;
			}
		};
	}

	/**
	 * Принимает ли данное поле в таблице значение NULL.
	 */
	public boolean isNull() {
		return data == null;
	}

	/**
	 * Сбрасывает BLOB в значение NULL.
	 */
	public void setNull() {
		isModified = isModified || (data != null);
		size = 0;
		data = null;
	}

	/**
	 * Возвращает размер данных.
	 */
	public int size() {
		return size;
	}

	/**
	 * Данные BLOB-поля.
	 */
	private static final class DataPage {
		private static final int DEFAULT_PAGE_SIZE = 0xFFFF;
		private static final int BYTE_MASK = 0xFF;

		private final byte[] data;
		private DataPage nextPage;
		private int pos;

		DataPage() {
			this(DEFAULT_PAGE_SIZE);
		}

		private DataPage(int size) {
			data = new byte[size];
		}

		private DataPage(int firstByte, InputStream source) throws IOException {
			this();
			int buf = firstByte;
			while (pos < data.length && buf >= 0) {
				data[pos++] = (byte) buf;
				buf = source.read();
			}
			nextPage = buf < 0 ? null : new DataPage(buf, source);
		}

		DataPage write(int b) {
			if (pos < data.length) {
				data[pos++] = (byte) (b & BYTE_MASK);
				return this;
			} else {
				DataPage result = new DataPage();
				nextPage = result;
				return result.write(b);
			}
		}

		InputStream getInStream() {
			return new InputStream() {
				private int i = 0;
				private DataPage currentPage = DataPage.this;

				@Override
				public int read() {
					if (i < currentPage.pos)
						return (int) currentPage.data[i++] & BYTE_MASK;
					else if (currentPage.nextPage != null) {
						i = 0;
						currentPage = currentPage.nextPage;
						return read();
					} else {
						return -1;
					}
				}
			};
		}
	}
}