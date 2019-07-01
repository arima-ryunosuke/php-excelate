<?php

namespace ryunosuke\Test\Excelate;

use ryunosuke\Excelate\Variable;

class VariableTest extends \ryunosuke\Test\Excelate\AbstractTestCase
{
    protected static $testBook = false;

    function test_arrayize()
    {
        $actual = Variable::arrayize([
            'A' => ['a' => 'A'],
            'B' => ['b' => 'B'],
            'C' => ['c' => 'C'],
            'X' => ['c' => 'C'],
        ], function ($v, $n, $k) {
            if ($k !== 'X') {
                $v['No'] = $n + 1;
                $v['Key'] = $k;
                return $v;
            }
        });

        $this->assertCount(3, $actual);
        $this->assertInstanceOf(Variable::class, $actual[0]);
        $this->assertInstanceOf(Variable::class, $actual[1]);
        $this->assertInstanceOf(Variable::class, $actual[2]);
        $this->assertEquals(2, $actual[1]->No);
        $this->assertEquals('B', $actual[1]->Key);
        $this->assertEquals('B', $actual[1]->b);
    }

    function test_all()
    {
        $object = new Variable(['hoge' => 'HOGE', 'fuga' => 'FUGA', 'piyo' => 'PIYO']);

        $this->assertFalse(isset($object['foo']));
        $this->assertTrue(isset($object['hoge']));

        $this->assertEquals('HOGE', $object['hoge']);

        $object['fuga'] = 'FUGAex';
        $this->assertEquals('FUGAex', $object['fuga']);

        $this->assertTrue(isset($object['piyo']));
        unset($object['piyo']);
        $this->assertFalse(isset($object['piyo']));

        $this->assertCount(2, $object);
    }

    function test___toString()
    {
        $object = new Variable(['hoge' => 'HOGE', 'fuga' => 'FUGA', 'piyo' => 'PIYO']);
        $this->assertEquals(print_r($object, 1), "$object");
    }
}
