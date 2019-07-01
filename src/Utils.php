<?php

namespace ryunosuke\Excelate;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\ReferenceHelper;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * 主に Worksheet を対象にしたユーティリティクラス
 */
class Utils
{
    /**
     * セル範囲を新行として追加する
     *
     * @param Worksheet $sheet
     * @param int $left
     * @param int $top
     * @param int $right
     * @param int $bottom
     * @param int $targetRow
     * @param int $length
     * @return int
     */
    public static function insertDuplicateRows(Worksheet $sheet, $left, $top, $right, $bottom, $targetRow, $length)
    {
        $height = $bottom - $top + 1;
        $size = $height * $length;

        if ($size === 0) {
            return 0;
        }

        if ($size < 0) {
            self::unmergeCells($sheet, $left, $top, $right, $bottom);
        }

        $sheet->insertNewRowBefore($targetRow, $size);

        for ($n = 0; $n < $length; $n++) {
            self::copyCells($sheet, $left, $top, $right, $bottom, null, $targetRow + $n * $height);
        }

        return $size;
    }

    /**
     * セル範囲を新列として追加する
     *
     * @param Worksheet $sheet
     * @param int $left
     * @param int $top
     * @param int $right
     * @param int $bottom
     * @param int $targetCol
     * @param int $length
     * @return int
     */
    public static function insertDuplicateCols(Worksheet $sheet, $left, $top, $right, $bottom, $targetCol, $length)
    {
        $width = $right - $left + 1;
        $size = $width * $length;

        if ($size === 0) {
            return 0;
        }

        if ($size < 0) {
            self::unmergeCells($sheet, $left, $top, $right, $bottom);
        }

        $sheet->insertNewColumnBefore(Coordinate::stringFromColumnIndex($targetCol), $size);

        for ($n = 0; $n < $length; $n++) {
            self::copyCells($sheet, $left, $top, $right, $bottom, $targetCol + $n * $width, null);
        }

        return $size;
    }

    /**
     * セル範囲を下へシフトする
     *
     * @param Worksheet $sheet
     * @param int $left
     * @param int $top
     * @param int $right
     * @param int $bottom
     * @param int $length
     * @param int $bottomLimit
     * @return int
     */
    public static function shiftDuplicateRows(Worksheet $sheet, $left, $top, $right, $bottom, $length, $bottomLimit = null)
    {
        if ($bottomLimit === null) {
            $bottomLimit = $sheet->getHighestRow();
        }

        if ($length < 0) {
            self::unmergeCells($sheet, $left, $top, $right, $bottom);
        }
        for ($i = 0; $i < $length; $i++) {
            self::copyCells($sheet, $left, $top, $right, $bottomLimit, null, $bottom + 1);
        }

        return ($bottom - $top + 1) * $length;
    }

    /**
     * セル範囲を右へシフトする
     *
     * @param Worksheet $sheet
     * @param int $left
     * @param int $top
     * @param int $right
     * @param int $bottom
     * @param int $length
     * @param int $rightLimit
     * @return int
     */
    public static function shiftDuplicateCols(Worksheet $sheet, $left, $top, $right, $bottom, $length, $rightLimit = null)
    {
        if ($rightLimit === null) {
            $rightLimit = Coordinate::columnIndexFromString($sheet->getHighestColumn());
        }
        if ($length < 0) {
            self::unmergeCells($sheet, $left, $top, $right, $bottom);
        }
        for ($i = 0; $i < $length; $i++) {
            self::copyCells($sheet, $left, $top, $rightLimit, $bottom, $right + 1, null);
        }

        return ($right - $left + 1) * $length;
    }

    /**
     * セル範囲内のセル結合を削除する
     *
     * @param Worksheet $sheet
     * @param int $left
     * @param int $top
     * @param int $right
     * @param int $bottom
     * @return int
     */
    public static function unmergeCells(Worksheet $sheet, $left, $top, $right, $bottom)
    {
        $count = 0;
        /** @var string $mergeCell */
        foreach ($sheet->getMergeCells() as $mergeCell) {
            $boundary = Coordinate::rangeBoundaries($mergeCell);
            list($bLeft, $bTop) = $boundary[0];
            list($bRight, $bBottom) = $boundary[1];
            if ($left <= $bLeft && $bRight <= $right && $top <= $bTop && $bBottom <= $bottom) {
                $sheet->unmergeCells($mergeCell);
                $count++;
            }
        }
        return $count;
    }

    /**
     * セル範囲をコピーする
     *
     * @param Worksheet $sheet
     * @param int $left
     * @param int $top
     * @param int $right
     * @param int $bottom
     * @param int $targetLeft
     * @param int $targetTop
     */
    public static function copyCells(Worksheet $sheet, $left, $top, $right, $bottom, $targetLeft = null, $targetTop = null)
    {
        // 指定されなければ同列・同行とする
        $targetLeft = $targetLeft ?? $left;
        $targetTop = $targetTop ?? $top;

        // 幅・高さ
        $width = $right - $left;
        $height = $bottom - $top;

        // 向きやサイズでループ方向を制御しないと重複範囲で値が死ぬ場合がある
        $cols = $left > $targetLeft ? range(0, $width, 1) : range($width, 0, -1);
        $rows = $top > $targetTop ? range(0, $height, 1) : range($height, 0, -1);

        // セルの書式と値の複製
        $helper = ReferenceHelper::getInstance();
        foreach ($cols as $col) {
            foreach ($rows as $row) {
                $srcCell = $sheet->getCellByColumnAndRow($left + $col, $top + $row);
                $dstCell = $sheet->getCellByColumnAndRow($targetLeft + $col, $targetTop + $row);

                $value = $srcCell->getValue();
                if ($srcCell->isFormula()) {
                    $value = $helper->updateFormulaReferences($value, 'A1', $targetLeft - $left, $targetTop - $top);
                    $dstCell->setValueExplicit($value, DataType::TYPE_FORMULA);
                }
                else {
                    $dstCell->setValue($value);
                }
                $style = $sheet->getStyleByColumnAndRow($left + $col, $top + $row);
                $sheet->duplicateStyle($style, $dstCell->getCoordinate());
            }
        }

        // セル結合の複製
        /** @var string $mergeCell */
        foreach ($sheet->getMergeCells() as $mergeCell) {
            $boundary = Coordinate::rangeBoundaries($mergeCell);
            list($bLeft, $bTop) = $boundary[0];
            list($bRight, $bBottom) = $boundary[1];
            if ($left <= $bLeft && $bRight <= $right && $top <= $bTop && $bBottom <= $bottom) {
                $leftIndex = $targetLeft + $bLeft - $left;
                $topIndex = $targetTop + $bTop - $top;
                $bWidth = $bRight - $bLeft;
                $bHeight = $bBottom - $bTop;
                $leftTop = Coordinate::stringFromColumnIndex($leftIndex) . $topIndex;
                $rightBottom = Coordinate::stringFromColumnIndex($leftIndex + $bWidth) . ($topIndex + $bHeight);
                $sheet->mergeCells("$leftTop:$rightBottom");
            }
        }
    }

    public static function dumpCellValues(Worksheet $sheet, $left, $top, $right, $bottom)
    {
        $values = [];
        for ($row = $top; $row <= $bottom; $row++) {
            for ($col = $left; $col <= $right; $col++) {
                $cell = $sheet->getCellByColumnAndRow($col, $row);
                $values[$cell->getCoordinate()] = $cell->getValue();
            }
        }
        return $values;
    }
}
