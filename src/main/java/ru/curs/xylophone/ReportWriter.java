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

import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Deque;
import java.util.LinkedList;
import java.util.stream.Stream;

/**
 * Класс, ответственный за формирование результирующего вывода в табличный
 * документ.
 */
abstract class ReportWriter {

    /**
     * "Точка роста" -- ячейка, в которую предполагается копировать следующий
     * элемент.
     */
    private final CellAddress growthPoint = new CellAddress("A1");

    private OutputStream output;

    /**
     * Блок раскладки, состоящий из вложенных блоков, растущих внутри него
     * горизонтально или вертикально.
     *
     */
    private static class LayoutBlock {

        /**
         * true, если раскладка элементов в блоке идёт слева направо.
         */
        private boolean horizontal;
        /**
         * Заполненные границы блока, null, если блок пустой.
         */
        private RangeAddress borders;

    }

    private final Deque<LayoutBlock> blocks = new LinkedList<>();

    /**
     * Создаёт форматировщик на основе шаблона и типа форматировщика.
     *
     * @param template
     *            Шаблон
     * @param type
     *            Тип форматировщика
     * @param output
     *            Поток для вывода результата
     * @throws XylophoneError
     *             В случае, если шаблон имеет неверный формат
     */
    static ReportWriter createWriter(InputStream template, OutputType type,
            boolean copyTemplate, OutputStream output)
            throws XylophoneError {
        InputStream localTemplate = template;
        InputStream templateCopy = null;
        try {
            if (copyTemplate) {
                XML2SpreadseetBLOB b = new XML2SpreadseetBLOB(template);
                localTemplate = b.getInStream();
                templateCopy = b.getInStream();
                template.close();
            }
        } catch (IOException e) {
            throw new XylophoneError(e.getMessage());
        }
        ReportWriter result;
        switch (type) {
        case ODS:
            result = new ODSReportWriter(localTemplate, templateCopy);
            break;
        case XLS:
            result = new XLSReportWriter(localTemplate, templateCopy);
            break;
        case XLSX:
            result = new XLSXReportWriter(localTemplate, templateCopy);
            break;
        default:
            // This will never happen
            return null;
        }
        result.output = output;
        return result;
    }

    /**
     * Создаёт новый лист электронной таблицы, копирует ширины столбцов из
     * исходного листа и переводит позицию вывода в левый верхний угол.
     *
     * @param sheetName
     *            Имя создаваемого листа
     * @param sourceSheet
     *            Имя листа шаблона, с которого будут копироваться ширины
     *            столбцов
     * @param startRepeatingColumn
     *            Стартовая колонка сквозных колонок
     * @param endRepeatingColumn
     *            Конечная колонка сквозных колонок
     * @param startRepeatingRow
     *            Стартовая колонка сквозных строк
     * @param endRepeatingRow
     *            Конечная колонка сквозных строк
     * @throws XylophoneError
     *             В случае ошибок операции
     */
    public void sheet(String sheetName, String sourceSheet,
            int startRepeatingColumn, int endRepeatingColumn,
            int startRepeatingRow, int endRepeatingRow)
            throws XylophoneError {
        newSheet(sheetName, sourceSheet, startRepeatingColumn,
                endRepeatingColumn, startRepeatingRow, endRepeatingRow);

        blocks.clear();
        blocks.push(new LayoutBlock());

        growthPoint.setAddress("A1");
    }

    /**
     * Начинает группу элементов, запоминая текущую позицию и режим вывода
     * (сверху вниз или слева направо).
     *
     * @param horizontal
     *            true для режима слева направо, false для режима сверху вниз.
     */
    public void startSequence(boolean horizontal) {
        // На вершине стека должен находиться как минимум диапазон листа. Если
        // пользователь не создал листа, делаем это за него.
        // if (blocks.isEmpty())
        // sheet("Sheet1", null);

        LayoutBlock b = new LayoutBlock();
        b.horizontal = horizontal;

        blocks.push(b);
    }

    /**
     * Оканчивает группу элементов, устанавливает текущую позицию вывода в
     * зависимости от режима охватывающей группы: например, если до этого группа
     * элементов выводилась слева направо, а охватывающая группа --- сверху
     * вниз, то позиция смещается в столбец A, и строку, определяемую высотой
     * самого высокого элемента в группе.
     *
     */
    public void endSequence(int merge, String regionName) {
        if (blocks.isEmpty())
            return;

        LayoutBlock b2 = blocks.pop();
        // Если в группе так ничего и не было выведено или лист кончился ---
        // ничего и не делаем
        if (b2.borders == null || blocks.isEmpty())
            return;

        RangeAddress borders = b2.borders;
        CellAddress tl = new CellAddress(borders.topLeft().getCol(), borders
                .topLeft().getRow());
        CellAddress br = b2.horizontal ? new CellAddress(borders.bottomRight()
                .getCol(), tl.getRow() + merge - 1) : new CellAddress(
                tl.getCol() + merge - 1, borders.bottomRight().getRow());

        // Делаем мёрдж тех ячеек, которые необходимы
        if (merge > 0)
            mergeUp(tl, br);

        // Добавляем именованный диапазон, если необходим
        if (regionName != null)
            addNamedRegion(regionName, tl, br);

        // Иначе, мы вставили прямоугольник, и его следует оприходовать, как
        // большую секцию:
        LayoutBlock b = blocks.peek();

        // Устанавливаем или расширяем границы текущего блока
        if (b.borders == null)
            b.borders = b2.borders;
        else {
            RangeAddress curRange = b.borders;
            curRange.setRight(Math.max(b2.borders.right(), curRange.right()));
            curRange.setBottom(Math.max(b2.borders.bottom(), curRange.bottom()));
        }

        // И устанавливаем точку роста в зависимости от режима раскладки
        updateGrowthPoint(b);
    }

    /**
     * Выводит новую секцию отчёта, в данном контексте и базируясь на данном
     * диапазоне ячеек шаблона. Позиция новой секции определяется режимом
     * вывода, заданном вызовом startSequence.
     *
     * @param context
     *            Контекст данных
     * @param sourceSheet
     *            Лист шаблона, с которого скопировать заполненный данными
     *            диапазон
     * @param range
     *            диапазон листа шаблона, который надо скопировать, заполнив
     *            данными
     * @throws XylophoneError
     *             Если что-то пошло не так.
     */
    public void section(XMLContext context, String sourceSheet,
            RangeAddress range, boolean pagebreak) throws XylophoneError {
        // На вершине стека должен находиться как минимум диапазон листа. Если
        // пользователь не создал листа, делаем это за него.
        if (blocks.isEmpty())
            sheet("Sheet1", null, -1, -1, -1, -1);

        // Укладываем секцию на лист
        putSection(context, growthPoint, sourceSheet, range);

        // Мы только что вставили прямоугольник.
        // Его ЛВУ является growthPoint, а ПНУ сейчас вычислим:
        CellAddress bottomRight = new CellAddress(growthPoint.getCol()
                + range.right() - range.left(), growthPoint.getRow()
                + range.bottom() - range.top());

        LayoutBlock b = blocks.peek();
        // Устанавливаем page break, если это необходимо
        processPageBreak(pagebreak, b);

        // Устанавливаем или расширяем границы текущего блока
        if (b.borders == null) {
            b.borders = new RangeAddress(growthPoint.getAddress() + ":"
                    + bottomRight.getAddress());
        } else {
            RangeAddress curRange = b.borders;
            curRange.setRight(Math.max(bottomRight.getCol(), curRange.right()));
            curRange.setBottom(Math.max(bottomRight.getRow(), curRange.bottom()));
        }

        // И устанавливаем точку роста в зависимости от режима раскладки
        updateGrowthPoint(b);
    }

    private void processPageBreak(boolean pagebreak, LayoutBlock b) {
        if (pagebreak) {
            if (b.horizontal) {
                int num = growthPoint.getCol() - 2;
                if (num >= 0)
                    putColBreak(num);
            } else {
                int num = growthPoint.getRow() - 2;
                if (num >= 0)
                    putRowBreak(num);
            }
        }
    }

    private void updateGrowthPoint(LayoutBlock b) {
        RangeAddress curRange = b.borders;
        if (b.horizontal) {
            growthPoint.setCol(curRange.right() + 1);
            growthPoint.setRow(curRange.top());
        } else {
            growthPoint.setCol(curRange.left());
            growthPoint.setRow(curRange.bottom() + 1);
        }
    }

    final OutputStream getOutput() {
        return output;
    }

    abstract void newSheet(String sheetName, String sourceSheet,
            int startRepeatingColumn, int endRepeatingColumn,
            int startRepeatingRow, int endRepeatingRow)
            throws XylophoneError;

    abstract void putSection(XMLContext context, CellAddress growthPoint2,
            String sourceSheet, RangeAddress range) throws XylophoneError;

    abstract void mergeUp(CellAddress a1, CellAddress a2);

    abstract void addNamedRegion(String name, CellAddress a1, CellAddress a2);

    abstract void putRowBreak(int rowNumber);

    abstract void putColBreak(int colNumber);

    /**
     * Сбрасывает результат создания документа в поток.
     *
     * @throws XylophoneError
     *             при возникновении ошибки сохранения
     */
    public abstract void flush() throws XylophoneError;

    abstract void applyMergedRegions(Stream<CellRangeAddress> mergedRegions);
}
