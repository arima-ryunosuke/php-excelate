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
            $asize = -$size;
            self::unmergeCells($sheet, $left, $top, $right, $bottom);
            $sheet->removeRow($targetRow - $asize, $asize);
        }
        else {
            $sheet->insertNewRowBefore($targetRow, $size);
            for ($n = 0; $n < $length; $n++) {
                self::copyCells($sheet, $left, $top, $right, $bottom, null, $targetRow + $n * $height);
            }
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
            $asize = -$size;
            self::unmergeCells($sheet, $left, $top, $right, $bottom);
            $sheet->removeColumn(Coordinate::stringFromColumnIndex($targetCol - $asize), $asize);
        }
        else {
            $sheet->insertNewColumnBefore(Coordinate::stringFromColumnIndex($targetCol), $size);
            for ($n = 0; $n < $length; $n++) {
                self::copyCells($sheet, $left, $top, $right, $bottom, $targetCol + $n * $width, null);
            }
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
        $height = $bottom - $top + 1;
        $size = $height * $length;

        if ($size === 0) {
            return 0;
        }

        if ($bottomLimit === null) {
            $bottomLimit = $sheet->getHighestRow();
        }

        if ($size < 0) {
            self::unmergeCells($sheet, $left, $top, $right, $bottom);
            self::copyCells($sheet, $left, $bottom + 1, $right, $bottomLimit, null, $top, 'move');
        }
        else {
            self::copyCells($sheet, $left, $bottom + 1, $right, $bottomLimit, null, $bottom + 1 + $size, 'move');
            for ($n = 0; $n < $length; $n++) {
                self::copyCells($sheet, $left, $top, $right, $bottom, null, $bottom + 1 + $n * $height);
            }
        }

        return $size;
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
        $width = $right - $left + 1;
        $size = $width * $length;

        if ($size === 0) {
            return 0;
        }

        if ($rightLimit === null) {
            $rightLimit = Coordinate::columnIndexFromString($sheet->getHighestColumn());
        }

        if ($size < 0) {
            self::unmergeCells($sheet, $left, $top, $right, $bottom);
            self::copyCells($sheet, $right + 1, $top, $rightLimit, $bottom, $left, null, 'move');
        }
        else {
            self::copyCells($sheet, $right + 1, $top, $rightLimit, $bottom, $right + 1 + $size, null, 'move');
            for ($n = 0; $n < $length; $n++) {
                self::copyCells($sheet, $left, $top, $right, $bottom, $right + 1 + $n * $width, null);
            }
        }

        return $size;
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
        foreach ($sheet->getMergeCells() as $mergeCell) {
            $boundary = Coordinate::rangeBoundaries($mergeCell);
            [$bLeft, $bTop] = $boundary[0];
            [$bRight, $bBottom] = $boundary[1];
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
     * @param string $mergedCellMethod
     */
    public static function copyCells(Worksheet $sheet, $left, $top, $right, $bottom, $targetLeft = null, $targetTop = null, $mergedCellMethod = 'copy')
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
                $srcCell = $sheet->getCell([$left + $col, $top + $row]);
                $dstCell = $sheet->getCell([$targetLeft + $col, $targetTop + $row]);

                $value = $srcCell->getValue();
                if ($srcCell->isFormula()) {
                    $value = $helper->updateFormulaReferences($value, 'A1', $targetLeft - $left, $targetTop - $top);
                    $dstCell->setValueExplicit($value, DataType::TYPE_FORMULA);
                }
                else {
                    $dstCell->setValue($value);
                }
                $dstCell->setXfIndex($srcCell->getXfIndex());
            }
        }

        // セル結合の移動・複製
        foreach ($sheet->getMergeCells() as $mergeCell) {
            $boundary = Coordinate::rangeBoundaries($mergeCell);
            [$bLeft, $bTop] = $boundary[0];
            [$bRight, $bBottom] = $boundary[1];
            if ($left <= $bLeft && $bRight <= $right && $top <= $bTop && $bBottom <= $bottom) {
                $leftIndex = $targetLeft + $bLeft - $left;
                $topIndex = $targetTop + $bTop - $top;
                $bWidth = $bRight - $bLeft;
                $bHeight = $bBottom - $bTop;
                $leftTop = Coordinate::stringFromColumnIndex($leftIndex) . $topIndex;
                $rightBottom = Coordinate::stringFromColumnIndex($leftIndex + $bWidth) . ($topIndex + $bHeight);
                $sheet->mergeCells("$leftTop:$rightBottom");
                if ($mergedCellMethod === 'move') {
                    $sheet->unmergeCells($mergeCell);
                }
            }
        }
    }

    public static function dumpCellValues(Worksheet $sheet, $left, $top, $right, $bottom)
    {
        $mb_str_pad = function (string $string, int $width, string $pad_string = " ", int $pad_type = STR_PAD_RIGHT): string {
            $padlength = $width - mb_strwidth($string);
            if ($padlength <= 0) {
                return $string;
            }
            if ($pad_type === STR_PAD_BOTH) {
                $padlength = $padlength / 2;
            }

            if (in_array($pad_type, [STR_PAD_BOTH, STR_PAD_LEFT], true)) {
                $string = str_repeat($pad_string, floor($padlength)) . $string;
            }
            if (in_array($pad_type, [STR_PAD_BOTH, STR_PAD_RIGHT], true)) {
                $string = $string . str_repeat($pad_string, ceil($padlength));
            }
            return $string;
        };

        $MINCOLUMN = 3;
        $SIDELINE = '│';
        $OVERLINE = '‾';
        $UNDERLINE = '_';
        $SPACE = ' ';

        $tb = range($top, $bottom);
        $lr = range($left, $right);

        $values = [-1 => [-1 => '#'] + array_combine($lr, array_map(fn($v) => Coordinate::stringFromColumnIndex($v), $lr))];
        foreach ($tb as $rowNo) {
            $values[$rowNo][-1] = $rowNo;
            foreach ($lr as $colNo) {
                $values[$rowNo][$colNo] = (string) $sheet->getCell([$colNo, $rowNo])->getValue();
            }
        }
        $widths = [-1 => $MINCOLUMN];
        foreach ($lr as $colNo) {
            $widths[$colNo] = max($MINCOLUMN, ...array_map('mb_strwidth', array_column($values, $colNo)));
        }
        $styles = [
            -1  => ['align' => STR_PAD_BOTH, 'delimiter' => $UNDERLINE],
            '*' => ['align' => STR_PAD_RIGHT, 'delimiter' => $SPACE],
        ];
        $lines = [];
        foreach ($values as $rowNo => $cols) {
            $style = $styles[$rowNo] ?? $styles['*'];
            $line = [];
            foreach ($cols as $colNo => $value) {
                $line[] = $mb_str_pad($value, $widths[$colNo], $style['delimiter'], $colNo === -1 ? STR_PAD_LEFT : $style['align']);
            }
            $lines[] = "$SIDELINE{$style['delimiter']}" . implode("{$style['delimiter']}$SIDELINE{$style['delimiter']}", $line) . "{$style['delimiter']}$SIDELINE";
        }

        $linesize = mb_strwidth($lines[0]);
        $V = fn($v) => $v;
        echo <<<TABLE
        {$V(str_repeat($UNDERLINE, $linesize))}
        {$V(implode("\n", $lines))}
        {$V(str_repeat($OVERLINE, $linesize))}
        
        TABLE;
    }

    public static function dumpRangeValues(Worksheet $sheet, $range)
    {
        $boundaries = Coordinate::rangeBoundaries($range);
        return self::dumpCellValues($sheet, $boundaries[0][0], $boundaries[0][1], $boundaries[1][0], $boundaries[1][1]);
    }
}
