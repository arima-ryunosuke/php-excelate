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

    function test_shiftDuplicateRows()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::shiftDuplicateRows($sheet, 1, 14, 1, 16, 3);
        $this->assertEquals(9, $delta);
        $this->assertRangeValues(<<<EXPECTED
        a
        b
        c
        a
        b
        c
        a
        b
        c
        
        EXPECTED, $sheet, 'A17:A26');
    }

    function test_shiftDuplicateRows_0()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::shiftDuplicateRows($sheet, 1, 14, 1, 16, 0);
        $this->assertEquals(0, $delta);
    }

    function test_shiftDuplicateCols()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::shiftDuplicateCols($sheet, 1, 14, 3, 14, 3);
        $this->assertEquals(9, $delta);
        $this->assertRangeValues(<<<EXPECTED
        a|b|c|a|b|c|a|b|c|
        EXPECTED, $sheet, 'D14:M14');
    }

    function test_shiftDuplicateCols_0()
    {
        $sheet = self::$testBook->getSheetByName('util');
        $delta = Utils::shiftDuplicateCols($sheet, 1, 14, 3, 14, 0);
        $this->assertEquals(0, $delta);
    }

    function test_copyCells()
    {
        $sheet = self::$testBook->getSheetByName('util');
        Utils::copyCells($sheet, 1, 10, 4, 12, 30, 10);
        $sheet->setCellValue('AD10', 10);
        $sheet->setCellValue('AD11', 20);
        $sheet->setCellValue('AD12', 30);
        $this->assertRangeValues(<<<EXPECTED
        10 | 20 | 30
        20 | 40 | 60
        30 | 60 | 90
        EXPECTED, $sheet, 'AE10:AG12', true);
    }

    function test_dumpRangeValues()
    {
        $sheet = self::$testBook->getSheetByName('util');
        ob_start();
        Utils::dumpRangeValues($sheet, 'B18:D21');
        $this->assertEquals(<<<TABLE
        _____________________________________
        │___#_│______B_______│__C__│___D____│
        │  18 │ longlonglong │ s   │ middle │
        │  19 │ a1           │ b1  │ c1     │
        │  20 │ a2           │ b2  │ c2     │
        │  21 │ a3           │ b3  │ c3     │
        ‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾
        
        TABLE, ob_get_clean());
    }
}
