<?php

namespace ryunosuke\Test\Excelate;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use ryunosuke\Excelate\Utils;

abstract class AbstractTestCase extends \PHPUnit\Framework\TestCase
{
    /** @var Spreadsheet */
    protected static $testBook;

    public static function setUpBeforeClass(): void
    {
        parent::setUpBeforeClass();

        if (static::$testBook !== false) {
            static::$testBook = IOFactory::load(__DIR__ . '/../test.xlsx');
        }
    }

    public static function tearDownAfterClass(): void
    {
        parent::tearDownAfterClass();

        if (static::$testBook !== false) {
            $filename = str_replace(__NAMESPACE__ . '\\', '', get_called_class());
            $filename = str_replace('\\', '_', $filename);
            IOFactory::createWriter(static::$testBook, 'Xlsx')->save(__DIR__ . "/../result/$filename.xlsx");
        }
    }

    public static function assertException(\Exception $e, callable $callback)
    {
        try {
            call_user_func_array($callback, array_slice(func_get_args(), 2));
        }
        catch (\Throwable $ex) {
            self::assertInstanceOf(get_class($e), $ex);
            self::assertEquals($e->getCode(), $ex->getCode());
            if (strlen($e->getMessage()) > 0) {
                self::assertStringContainsString($e->getMessage(), $ex->getMessage());
            }
            return;
        }
        self::fail(get_class($e) . ' is not thrown.');
    }

    public static function assertRangeValues($expected, Worksheet $sheet, $range, $formattedValue = false)
    {
        if (!is_array($expected)) {
            $expected = preg_split('#\\R#u', $expected);
        }
        $expected = array_map(function ($v) {
            if (!is_array($v)) {
                $v = array_map('trim', explode('|', $v));
            }
            return $v;
        }, $expected);

        $actual = [];
        $boundaries = Coordinate::rangeBoundaries($range);
        foreach (range($boundaries[0][1], $boundaries[1][1]) as $y) {
            $line = [];
            foreach (range($boundaries[0][0], $boundaries[1][0]) as $x) {
                $cell = $sheet->getCellByColumnAndRow($x, $y);
                $line[] = trim((string) ($formattedValue ? $cell->getFormattedValue() : $cell->getValue()));
            }
            $actual[] = $line;
        }
        foreach ($actual as $n => $line) {
            self::assertEquals(implode(' | ', $expected[$n] ?? []), implode(' | ', $line), "failed " . ($n + 1) . " row");
        }
        self::assertCount(count($actual), $expected);
    }
}
