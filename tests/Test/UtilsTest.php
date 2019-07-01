<?php

namespace ryunosuke\Test\Excelate;

use ryunosuke\Excelate\Utils;

class UtilsTest extends \ryunosuke\Test\Excelate\AbstractTestCase
{
    function test_insertDuplicateRows()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::insertDuplicateRows($sheet, 1, 2, 1, 2, 30, 1);
        $this->assertEquals(1, $delta);
        $delta = Utils::insertDuplicateRows($sheet, 1, 2, 3, 4, 32, 2);
        $this->assertEquals(6, $delta);
    }

    function test_insertDuplicateRows_0()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::insertDuplicateRows($sheet, 1, 0, 1, 0, 4, 0);
        $this->assertEquals(0, $delta);
    }

    function test_insertDuplicateRows_merged()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::insertDuplicateRows($sheet, 1, 7, 6, 7, 39, 2);
        $this->assertEquals(2, $delta);
    }

    function test_insertDuplicateCols()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::insertDuplicateCols($sheet, 1, 2, 1, 2, 30, 1);
        $this->assertEquals(1, $delta);
        $delta = Utils::insertDuplicateCols($sheet, 1, 2, 3, 4, 32, 2);
        $this->assertEquals(6, $delta);
    }

    function test_insertDuplicateCols_0()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::insertDuplicateCols($sheet, 1, 0, 1, 0, 4, 0);
        $this->assertEquals(0, $delta);
    }

    function test_insertDuplicateCols_merged()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::insertDuplicateCols($sheet, 1, 7, 6, 7, 39, 1);
        $this->assertEquals(6, $delta);
    }

    function test_shiftDuplicateCols()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::shiftDuplicateCols($sheet, 1, 14, 3, 14, 3);
        $this->assertEquals(9, $delta);
        $this->assertEquals('a', $sheet->getCell('D14')->getValue());
        $this->assertEquals('b', $sheet->getCell('E14')->getValue());
        $this->assertEquals('c', $sheet->getCell('F14')->getValue());
        $this->assertEquals('a', $sheet->getCell('G14')->getValue());
        $this->assertEquals('b', $sheet->getCell('H14')->getValue());
        $this->assertEquals('c', $sheet->getCell('I14')->getValue());
        $this->assertEquals('a', $sheet->getCell('J14')->getValue());
        $this->assertEquals('b', $sheet->getCell('K14')->getValue());
        $this->assertEquals('c', $sheet->getCell('L14')->getValue());
        $this->assertEquals('', $sheet->getCell('M14')->getValue());
    }

    function test_shiftDuplicateRows()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::shiftDuplicateRows($sheet, 1, 14, 1, 16, 3);
        $this->assertEquals(9, $delta);
        $this->assertEquals('a', $sheet->getCell('A17')->getValue());
        $this->assertEquals('b', $sheet->getCell('A18')->getValue());
        $this->assertEquals('c', $sheet->getCell('A19')->getValue());
        $this->assertEquals('a', $sheet->getCell('A20')->getValue());
        $this->assertEquals('b', $sheet->getCell('A21')->getValue());
        $this->assertEquals('c', $sheet->getCell('A22')->getValue());
        $this->assertEquals('a', $sheet->getCell('A23')->getValue());
        $this->assertEquals('b', $sheet->getCell('A24')->getValue());
        $this->assertEquals('c', $sheet->getCell('A25')->getValue());
        $this->assertEquals('', $sheet->getCell('A26')->getValue());
    }

    function test_copyCells()
    {
        $sheet = self::$testBook->getSheetByName('util');
        Utils::copyCells($sheet, 1, 10, 4, 12, 30, 10);
        $sheet->setCellValue('AD10', 10);
        $sheet->setCellValue('AD11', 20);
        $sheet->setCellValue('AD12', 30);
        $this->assertEquals(10, $sheet->getCell('AE10')->getFormattedValue());
        $this->assertEquals(20, $sheet->getCell('AF10')->getFormattedValue());
        $this->assertEquals(30, $sheet->getCell('AG10')->getFormattedValue());
        $this->assertEquals(20, $sheet->getCell('AE11')->getFormattedValue());
        $this->assertEquals(40, $sheet->getCell('AF11')->getFormattedValue());
        $this->assertEquals(60, $sheet->getCell('AG11')->getFormattedValue());
        $this->assertEquals(30, $sheet->getCell('AE12')->getFormattedValue());
        $this->assertEquals(60, $sheet->getCell('AF12')->getFormattedValue());
        $this->assertEquals(90, $sheet->getCell('AG12')->getFormattedValue());
    }

    function test_dumpCellValues()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $this->assertEquals([
            'B2' => 'a',
            'C2' => 'b',
            'D2' => 'c',
            'B3' => 'd',
            'C3' => 'e',
            'D3' => 'f',
            'B4' => 'g',
            'C4' => 'h',
            'D4' => 'i',
        ], Utils::dumpCellValues($sheet, 2, 2, 4, 4));
    }
}
