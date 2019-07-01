<?php

namespace ryunosuke\Test\Excelate;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use ryunosuke\Excelate\Renderer;

class RendererTest extends \ryunosuke\Test\Excelate\AbstractTestCase
{
    function test_template()
    {
        $renderer = new Renderer();
        $sheet = self::$testBook->getSheet(0);
        $renderer->render($sheet, [
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
        $delta = $renderer->render($sheet, [
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
        $delta = $renderer->render($sheet, [
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
        $delta = $renderer->render($sheet, [
            'true'   => true,
            'false'  => false,
            'items'  => [
                ['flag' => true, 'value1' => 1, 'value2' => 2, 'value3' => 3, 'value4' => 4, 'value5' => 5],
                ['flag' => false, 'value1' => -1, 'value2' => -2, 'value3' => -3, 'value4' => -4, 'value5' => -5],
            ],
            'string' => 'hello',
        ]);
        $this->assertEquals([0, -3], $delta);
        $this->assertEquals('1', $sheet->getCell('A2')->getValue());
        $this->assertEquals('2', $sheet->getCell('B2')->getValue());
        $this->assertEquals('3', $sheet->getCell('C2')->getValue());
        $this->assertEquals('4', $sheet->getCell('D2')->getValue());
        $this->assertEquals('5', $sheet->getCell('E2')->getValue());

        $this->assertEquals('1', $sheet->getCell('A3')->getValue());
        $this->assertEquals('2', $sheet->getCell('B3')->getValue());
        $this->assertEquals('3', $sheet->getCell('C3')->getValue());
        $this->assertEquals('4', $sheet->getCell('D3')->getValue());
        $this->assertEquals('5', $sheet->getCell('E3')->getValue());
        $this->assertEquals('-1', $sheet->getCell('A4')->getValue());
        $this->assertEquals('-2', $sheet->getCell('B4')->getValue());
        $this->assertEquals('-3', $sheet->getCell('C4')->getValue());
        $this->assertEquals('-4', $sheet->getCell('D4')->getValue());
        $this->assertEquals('-5', $sheet->getCell('E4')->getValue());

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
        $delta = $renderer->render($sheet, [
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
        $this->assertEquals(7, $delta[1]);

        $this->assertEquals('0first', $sheet->getCell('B2')->getValue());
        $this->assertEquals('1', $sheet->getCell('B3')->getValue());
        $this->assertEquals('2last', $sheet->getCell('B4')->getValue());

        $this->assertEquals('0first', $sheet->getCell('B6')->getValue());
        $this->assertEquals('0firstlast', $sheet->getCell('D6')->getValue());
        $this->assertEquals('1', $sheet->getCell('B7')->getValue());
        $this->assertEquals('0first', $sheet->getCell('D7')->getValue());
        $this->assertEquals('', $sheet->getCell('B8')->getValue());
        $this->assertEquals('1last', $sheet->getCell('D8')->getValue());
        $this->assertEquals('2last', $sheet->getCell('B9')->getValue());
        $this->assertEquals('0first', $sheet->getCell('D9')->getValue());
        $this->assertEquals('', $sheet->getCell('B10')->getValue());
        $this->assertEquals('1', $sheet->getCell('D10')->getValue());
        $this->assertEquals('', $sheet->getCell('B11')->getValue());
        $this->assertEquals('2last', $sheet->getCell('D11')->getValue());
    }

    function test_coleach()
    {
        $renderer = new Renderer();
        $delta = $renderer->render(self::$testBook->getSheetByName('coleach'), [
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
        $delta = $renderer->render(self::$testBook->getSheetByName('rowshift'), [
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
        $delta = $renderer->render(self::$testBook->getSheetByName('colshift'), [
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
        $renderer->render($sheet, [
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
        $renderer->render($sheet, [
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
            $renderer->render($misc, ['notfound' => null], 'C3:C3');
        });
    }

    function test_delim()
    {
        $renderer = new Renderer();
        $renderer->registerVariable('globalValue', 'hogera');
        $misc = self::$testBook->getSheetByName('misc')->copy();
        $renderer->render($misc, ['notfound' => null], 'A2:A3');
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
        $renderer->render($misc, ['Name' => 'hoge']);
        $this->assertEquals('', $misc->getCell('A1')->getValue());

        error_clear_last();
        $misc = self::$testBook->getSheetByName('misc')->copy();
        $renderer->setErrorMode(Renderer::ERROR_MODE_RENDERING);
        $renderer->render($misc, ['Name' => 'hoge']);
        $this->assertEquals('Undefined variable: notfound', $misc->getCell('A1')->getValue());

        error_clear_last();
        $misc = self::$testBook->getSheetByName('misc')->copy();
        $renderer->setErrorMode(Renderer::ERROR_MODE_WARNING);
        @$renderer->render($misc, ['Name' => 'hoge']);
        $this->assertEquals('', $misc->getCell('A1')->getValue());
        $this->assertEquals('$notfound', error_get_last()['message']);

        error_clear_last();
        $this->expectException(get_class(new \ErrorException()));
        $misc = self::$testBook->getSheetByName('misc')->copy();
        $renderer->setErrorMode(Renderer::ERROR_MODE_EXCEPTION);
        $renderer->render($misc, ['Name' => 'hoge']);
    }
}
