<?php

namespace ryunosuke\Test\Excelate;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

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
}
