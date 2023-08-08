<?php

namespace ryunosuke\Test\Excelate;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use ryunosuke\Excelate\Renderer;

class RendererTest extends \ryunosuke\Test\Excelate\AbstractTestCase
{
    function test_renderBook()
    {
        $renderer = new Renderer();
        $bookFile = $renderer->renderBook(__DIR__ . '/../test.xlsx', [
            // @formatter:off
            'sheet1'          => [
                'title' => 'X',
                'A1' => 'x-a1', 'B1' => 'x-b1', 'C1' => 'x-c1',
                'A2' => 'x-a2', 'B2' => 'x-b2', 'C2' => 'x-c2',
                'A3' => 'x-a3', 'B3' => 'x-b3', 'C3' => 'x-c3',
            ],
            '1'               => [
                'title' => 'Y',
                'A1' => 'y-a1', 'B1' => 'y-b1', 'C1' => 'y-c1',
                'A2' => 'y-a2', 'B2' => 'y-b2', 'C2' => 'y-c2',
                'A3' => 'y-a3', 'B3' => 'y-b3', 'C3' => 'y-c3',
            ],
            'undefined-sheet' => [],
            // @formatter:on
        ]);

        $book = IOFactory::load($bookFile);
        $this->assertRangeValues(<<<EXPECTED
         x-a1| x-b1| x-c1
         x-a2| x-b2| x-c2
         x-a3| x-b3| x-c3
        EXPECTED, $book->getSheetByName('sheet1'), 'A1:C3');
        $this->assertRangeValues(<<<EXPECTED
         y-a1| y-b1| y-c1
         y-a2| y-b2| y-c2
         y-a3| y-b3| y-c3
        EXPECTED, $book->getSheetByName('sheet2Y'), 'A1:C3');
    }

    function test_template()
    {
        $renderer = new Renderer();
        $sheet = self::$testBook->getSheet(2);
        $renderer->renderSheet($sheet, [
            'value' => 'tValue',
            'st'    => 'hogera',
        ]);
        $this->assertEquals('aHOGERAz', $sheet->getTitle());
        // A1 が問題なくレンダリングされているはず
        $this->assertEquals('tValuevalue', $sheet->getCell('A1')->getValue());
        // B2 は $value の値のはず
        $this->assertEquals('tValue', $sheet->getCell('B2')->getValue());
        // 範囲外はレンダリングされていないはず
        $this->assertEquals('{$value}', $sheet->getCell('C1')->getValue());
        $this->assertEquals('{$value}', $sheet->getCell('A3')->getValue());
    }

    function test_rowif()
    {
        $renderer = new Renderer();
        $sheet = self::$testBook->getSheetByName('rowif');
        $delta = $renderer->renderSheet($sheet, [
            'true'   => true,
            'false'  => false,
            'string' => 'hello',
        ]);
        $this->assertEquals([0, -2], $delta);
        // $true はレンダリングされる
        $this->assertEquals('1', $sheet->getCell('A2')->getValue());
        $this->assertEquals('10', $sheet->getCell('E3')->getValue());
        // $false はレンダリングされない（直下の $string がレンダリングされる）
        $this->assertEquals('hello', $sheet->getCell('A4')->getValue());
        // テンプレートぴったりでないとシフトされる
        $this->assertEquals('shift', $sheet->getCell('A8')->getValue());
        $this->assertEquals('shift', $sheet->getCell('B5')->getValue());
        $this->assertEquals('shift', $sheet->getCell('C5')->getValue());
        $this->assertEquals('shift', $sheet->getCell('D5')->getValue());
        $this->assertEquals('shift', $sheet->getCell('E8')->getValue());
    }

    function test_colif()
    {
        $renderer = new Renderer();
        $sheet = self::$testBook->getSheetByName('colif');
        $delta = $renderer->renderSheet($sheet, [
            'true'   => true,
            'false'  => false,
            'string' => 'hello',
        ]);
        $this->assertEquals([-2, 0], $delta);
        // $false はレンダリングされない（直右の $string がレンダリングされる）
        $this->assertEquals('hello', $sheet->getCell('B5')->getValue());
        // $true はレンダリングされる
        $this->assertEquals('1', $sheet->getCell('C5')->getValue());
        $this->assertEquals('2', $sheet->getCell('D5')->getValue());
        // テンプレートぴったりでないとシフトされる
        $this->assertEquals('shift', $sheet->getCell('G1')->getValue());
        $this->assertEquals('shift', $sheet->getCell('G2')->getValue());
        $this->assertEquals('shift', $sheet->getCell('G3')->getValue());
        $this->assertEquals('shift', $sheet->getCell('G4')->getValue());
        $this->assertEquals('shift', $sheet->getCell('E5')->getValue());
    }

    function test_ifmisc()
    {
        $renderer = new Renderer();
        $sheet = self::$testBook->getSheetByName('ifmisc');
        $delta = $renderer->renderSheet($sheet, [
            'true'   => true,
            'false'  => false,
            'items'  => [
                ['flag' => true, 'value1' => 1, 'value2' => 2, 'value3' => 3, 'value4' => 4, 'value5' => 5],
                ['flag' => false, 'value1' => -1, 'value2' => -2, 'value3' => -3, 'value4' => -4, 'value5' => -5],
            ],
            'string' => 'hello',
        ]);
        $this->assertEquals([0, -3], $delta);

        $this->assertRangeValues(<<<EXPECTED
         1| 2| 3| 4| 5
         1| 2| 3| 4| 5
        -1|-2|-3|-4|-5
        EXPECTED, $sheet, 'A2:E4');

        $this->assertEquals('hello', $sheet->getCell('A5')->getValue());

        $this->assertEquals('true && true', $sheet->getCell('B6')->getValue());
        $this->assertEquals('true || true', $sheet->getCell('B7')->getValue());
        $this->assertEquals('true || false', $sheet->getCell('B8')->getValue());

        $this->assertEquals('hello', $sheet->getCell('A9')->getValue());

        $this->assertEquals('1', $sheet->getCell('B11')->getValue());
        $this->assertEquals('2', $sheet->getCell('C11')->getValue());
        $this->assertEquals('3', $sheet->getCell('D11')->getValue());
        $this->assertEquals('4', $sheet->getCell('E11')->getValue());

        $this->assertEquals('hello', $sheet->getCell('A12')->getValue());
    }

    function test_roweach()
    {
        $renderer = new Renderer();
        $sheet = self::$testBook->getSheetByName('roweach');
        $delta = $renderer->renderSheet($sheet, [
            'dummys1' => [
                ['hoge' => 'HOGE1', 'fuga' => 'FUGA1', 'piyo' => 'PIYO1'],
                ['hoge' => 'HOGE2', 'fuga' => 'FUGA2', 'piyo' => 'PIYO2'],
                ['hoge' => 'HOGE3', 'fuga' => 'FUGA3', 'piyo' => 'PIYO3'],
            ],
            'dummys2' => [
                ['title' => 'A', 'children' => [1]],
                ['title' => 'B', 'children' => [2, 3]],
                ['title' => 'C', 'children' => [4, 5, 6]],
            ],
            'dummys3' => [
                [
                    'title'    => 'A',
                    'children' => [
                        ['hoge' => 'HOGE_A1', 'fuga' => 'FUGA_A1', 'piyo' => 'PIYO_A1'],
                        ['hoge' => 'HOGE_A2', 'fuga' => 'FUGA_A2', 'piyo' => 'PIYO_A2'],
                        ['hoge' => 'HOGE_A3', 'fuga' => 'FUGA_A3', 'piyo' => 'PIYO_A3'],
                    ],
                ],
                [
                    'title'    => 'B',
                    'children' => [
                        ['hoge' => 'HOGE_B1', 'fuga' => 'FUGA_B1', 'piyo' => 'PIYO_B1'],
                        ['hoge' => 'HOGE_B2', 'fuga' => 'FUGA_B2', 'piyo' => 'PIYO_B2'],
                    ],
                ],
                [
                    'title'    => 'C',
                    'children' => [
                        ['hoge' => 'HOGE_C1', 'fuga' => 'FUGA_C1', 'piyo' => 'PIYO_C1'],
                    ],
                ],
            ],
        ]);
        $this->assertEquals(14, $delta[1]);

        $this->assertRangeValues(<<<EXPECTED
        0first | | HOGE1 | FUGA1 | PIYO1 | |
        1      | | HOGE2 | FUGA2 | PIYO2 | |
        2last  | | HOGE3 | FUGA3 | PIYO3 | |
        EXPECTED, $sheet, 'B2:H4');

        $this->assertRangeValues(<<<EXPECTED
        0first | A | 0firstlast | 1 | | A |
        1      | B | 0first     | 2 | | B |
               |   | 1last      | 3 | |   |
        2last  | C | 0first     | 4 | | C |
               |   | 1          | 5 | |   |
               |   | 2last      | 6 | |   |
        EXPECTED, $sheet, 'B6:H11');

        $this->assertRangeValues(<<<EXPECTED
        0first | A |         |         |         | |
               |   | HOGE_A1 | FUGA_A1 | PIYO_A1 | |
               |   | HOGE_A2 | FUGA_A2 | PIYO_A2 | |
               |   | HOGE_A3 | FUGA_A3 | PIYO_A3 | |
        1      | B |         |         |         | |
               |   | HOGE_B1 | FUGA_B1 | PIYO_B1 | |
               |   | HOGE_B2 | FUGA_B2 | PIYO_B2 | |
        2last  | C |         |         |         | |
               |   | HOGE_C1 | FUGA_C1 | PIYO_C1 | |
        EXPECTED, $sheet, 'B13:H21');
    }

    function test_coleach()
    {
        $renderer = new Renderer();
        $delta = $renderer->renderSheet(self::$testBook->getSheetByName('coleach'), [
            'dummys1' => [
                ['hoge' => 'HOGE1', 'fuga' => 'FUGA1', 'piyo' => 'PIYO1'],
                ['hoge' => 'HOGE2', 'fuga' => 'FUGA2', 'piyo' => 'PIYO2'],
                ['hoge' => 'HOGE3', 'fuga' => 'FUGA3', 'piyo' => 'PIYO3'],
            ],
            'dummys2' => [
                ['title' => 'A', 'children' => [1]],
                ['title' => 'B', 'children' => [2, 3]],
                ['title' => 'C', 'children' => [4, 5, 6]],
            ],
        ]);
        $this->assertEquals(7, $delta[0]);
    }

    function test_rowshift()
    {
        $renderer = new Renderer();
        $delta = $renderer->renderSheet(self::$testBook->getSheetByName('rowshift'), [
            'values' => [
                'hoge',
                'fuga',
                'piyo',
            ],
        ]);
        $this->assertEquals(6, $delta[1]);
    }

    function test_colshift()
    {
        $renderer = new Renderer();
        $delta = $renderer->renderSheet(self::$testBook->getSheetByName('colshift'), [
            'values' => [
                'hoge',
                'fuga',
                'piyo',
            ],
        ]);
        $this->assertEquals(6, $delta[0]);
    }

    function test_merge()
    {
        $renderer = new Renderer();
        $sheet = self::$testBook->getSheetByName('merge');
        $renderer->renderSheet($sheet, [
            'empty' => [],
        ]);
        $this->assertCount(4, $sheet->getMergeCells());
    }

    function test_effector()
    {
        $renderer = new Renderer();
        $renderer->registerEffector('BGColor', function (Cell $cell, $color) {
            $cell->getStyle()->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB($color);
        });
        $sheet = self::$testBook->getSheetByName('effector');
        $renderer->renderSheet($sheet, [
            'Name'   => 'hoge',
            'Path'   => __DIR__ . '/../test.png',
            'Attrs1' => [
                'path'        => __DIR__ . '/../test.png',
                'description' => 'descriptionです',
                'sizeToFit'   => true,
                'width'       => 160,
                'height'      => 32,
            ],
            'Attrs2' => [
                'path'        => __DIR__ . '/../test.png',
                'description' => 'descriptionです',
                'sizeToFit'   => true,
                'width'       => 32,
                'height'      => 160,
            ],
        ]);
        $this->assertEquals('sheet://util!B2', $sheet->getCell('A1')->getHyperlink()->getUrl());
        $this->assertEquals('http://example.com', $sheet->getCell('A2')->getHyperlink()->getUrl());
        $this->assertEquals('http://example.com', $sheet->getCell('A2')->getValue());
        $this->assertEquals('http://example.com', $sheet->getCell('B2')->getHyperlink()->getUrl());
        $this->assertEquals('aaalink textzzz', $sheet->getCell('B2')->getValue());

        $this->assertEquals('0000FF', $sheet->getCell('A3')->getStyle()->getFont()->getColor()->getRGB());
        $this->assertEquals('0000FF', $sheet->getCell('A4')->getStyle()->getFill()->getStartColor()->getRGB());

        $border = $sheet->getCell('A5')->getStyle()->getBorders();
        $this->assertEquals('FF0000', $border->getTop()->getColor()->getRGB());
        $this->assertEquals('FF0000', $border->getRight()->getColor()->getRGB());
        $this->assertEquals('FF0000', $border->getBottom()->getColor()->getRGB());
        $this->assertEquals('FF0000', $border->getLeft()->getColor()->getRGB());

        $border = $sheet->getCell('B5')->getStyle()->getBorders();
        $this->assertEquals('FF0000', $border->getTop()->getColor()->getRGB());
        $this->assertEquals('00FF00', $border->getRight()->getColor()->getRGB());
        $this->assertEquals('FF0000', $border->getBottom()->getColor()->getRGB());
        $this->assertEquals('00FF00', $border->getLeft()->getColor()->getRGB());

        $border = $sheet->getCell('C5')->getStyle()->getBorders();
        $this->assertEquals('FF0000', $border->getTop()->getColor()->getRGB());
        $this->assertEquals('00FF00', $border->getRight()->getColor()->getRGB());
        $this->assertEquals('0000FF', $border->getBottom()->getColor()->getRGB());
        $this->assertEquals('00FF00', $border->getLeft()->getColor()->getRGB());

        $border = $sheet->getCell('D5')->getStyle()->getBorders();
        $this->assertEquals('FF0000', $border->getTop()->getColor()->getRGB());
        $this->assertEquals('00FF00', $border->getRight()->getColor()->getRGB());
        $this->assertEquals('0000FF', $border->getBottom()->getColor()->getRGB());
        $this->assertEquals('F0000F', $border->getLeft()->getColor()->getRGB());

        $border = $sheet->getCell('E5')->getStyle()->getBorders();
        $this->assertEquals('000000', $border->getTop()->getColor()->getRGB());
        $this->assertEquals('000000', $border->getRight()->getColor()->getRGB());
        $this->assertEquals('000000', $border->getBottom()->getColor()->getRGB());
        $this->assertEquals('000000', $border->getLeft()->getColor()->getRGB());
    }

    function test_misc()
    {
        $renderer = new Renderer();

        $this->assertException(new \InvalidArgumentException(), function () use ($renderer) {
            $renderer->setErrorMode(99);
        });

        $this->assertException(new \DomainException(), function () use ($renderer) {
            $misc = self::$testBook->getSheetByName('misc')->copy();
            $renderer->renderSheet($misc, ['notfound' => null], 'C3:C3');
        });
    }

    function test_delim()
    {
        $renderer = new Renderer();
        $renderer->registerVariable('globalValue', 'hogera');
        $misc = self::$testBook->getSheetByName('misc')->copy();
        $renderer->renderSheet($misc, ['notfound' => null], 'A2:A3');
        $this->assertEquals('{hoge}', $misc->getCell('A2')->getValue());
        $this->assertEquals('hogera', $misc->getCell('A3')->getValue());
    }

    function test_error()
    {
        $renderer = new Renderer();
        $renderer->registerVariable('globalValue', 'hogera');

        error_clear_last();
        $misc = self::$testBook->getSheetByName('misc')->copy();
        $renderer->setErrorMode(Renderer::ERROR_MODE_SILENT);
        $renderer->renderSheet($misc, ['Name' => 'hoge']);
        $this->assertEquals('', $misc->getCell('A1')->getValue());

        error_clear_last();
        $misc = self::$testBook->getSheetByName('misc')->copy();
        $renderer->setErrorMode(Renderer::ERROR_MODE_RENDERING);
        $renderer->renderSheet($misc, ['Name' => 'hoge']);
        $this->assertEquals('Undefined variable: notfound', $misc->getCell('A1')->getValue());

        error_clear_last();
        $misc = self::$testBook->getSheetByName('misc')->copy();
        $renderer->setErrorMode(Renderer::ERROR_MODE_WARNING);
        @$renderer->renderSheet($misc, ['Name' => 'hoge']);
        $this->assertEquals('', $misc->getCell('A1')->getValue());
        $this->assertEquals('$notfound', error_get_last()['message']);

        error_clear_last();
        $this->expectException(get_class(new \ErrorException()));
        $misc = self::$testBook->getSheetByName('misc')->copy();
        $renderer->setErrorMode(Renderer::ERROR_MODE_EXCEPTION);
        $renderer->renderSheet($misc, ['Name' => 'hoge']);
    }
}
