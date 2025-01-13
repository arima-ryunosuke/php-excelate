<?php

namespace demo;

use ryunosuke\Excelate\Renderer;

require_once __DIR__ . '/../vendor/autoload.php';

$renderer = new Renderer();
$outfile = $renderer->renderBook(__DIR__ . '/template.xlsx', [
    0 => [
        'title' => 'example',
        'rows'  => [
            ['no' => 1, 'name' => 'hoge', 'attrs' => ['attr1', 'attr2']],
            ['no' => 2, 'name' => 'fuga', 'attrs' => ['attr1', 'attr2', 'attr3']],
        ],
    ],
    1 => [
        'False'   => false,
        'True'    => true,
        'Values1' => [
            ['val10' => 10, 'val11' => 11, 'val12' => 12, 'val13' => 13, 'val14' => 14, 'val15' => 15, 'val16' => 16, 'val17' => 17, 'val18' => 18, 'val19' => 19, 'val20' => 20, 'val21' => 21, 'val22' => 22, 'val23' => 23, 'val24' => 24, 'val25' => 25, 'val26' => 26, 'val27' => 27, 'val28' => 28, 'val29' => 29,],
            ['val10' => 10, 'val11' => 11, 'val12' => 12, 'val13' => 13, 'val14' => 14, 'val15' => 15, 'val16' => 16, 'val17' => 17, 'val18' => 18, 'val19' => 19, 'val20' => 20, 'val21' => 21, 'val22' => 22, 'val23' => 23, 'val24' => 24, 'val25' => 25, 'val26' => 26, 'val27' => 27, 'val28' => 28, 'val29' => 29,],
        ],
        'Values2' => [
            ['valFlag' => true, 'val10' => 10, 'val11' => 11, 'val12' => 12, 'val13' => 13, 'val14' => 14, 'val15' => 15, 'val16' => 16, 'val17' => 17, 'val18' => 18, 'val19' => 19, 'val20' => 20, 'val21' => 21, 'val22' => 22, 'val23' => 23, 'val24' => 24, 'val25' => 25, 'val26' => 26, 'val27' => 27, 'val28' => 28, 'val29' => 29,],
            ['valFlag' => false, 'val10' => 10, 'val11' => 11, 'val12' => 12, 'val13' => 13, 'val14' => 14, 'val15' => 15, 'val16' => 16, 'val17' => 17, 'val18' => 18, 'val19' => 19, 'val20' => 20, 'val21' => 21, 'val22' => 22, 'val23' => 23, 'val24' => 24, 'val25' => 25, 'val26' => 26, 'val27' => 27, 'val28' => 28, 'val29' => 29,],
        ],
    ],
    2 => [
        'Values1' => [
            ['val10' => 10, 'val11' => 11, 'val12' => 12, 'val13' => 13, 'val14' => 14, 'val15' => 15, 'val16' => 16, 'val17' => 17, 'val18' => 18, 'val19' => 19, 'val20' => 20, 'val21' => 21, 'val22' => 22, 'val23' => 23, 'val24' => 24, 'val25' => 25, 'val26' => 26, 'val27' => 27, 'val28' => 28, 'val29' => 29,],
            ['val10' => 10, 'val11' => 11, 'val12' => 12, 'val13' => 13, 'val14' => 14, 'val15' => 15, 'val16' => 16, 'val17' => 17, 'val18' => 18, 'val19' => 19, 'val20' => 20, 'val21' => 21, 'val22' => 22, 'val23' => 23, 'val24' => 24, 'val25' => 25, 'val26' => 26, 'val27' => 27, 'val28' => 28, 'val29' => 29,],
        ],
        'Values2' => [
            ['no' => 1, 'name' => 'hoge', 'attrs' => ['attr1', 'attr2']],
            ['no' => 2, 'name' => 'fuga', 'attrs' => ['attr1', 'attr2', 'attr3']],
        ],
        'Values3' => [
            ['val' => 10],
            ['val' => 20],
            ['val' => 30],
        ],
    ],
    3 => [
        'String' => 'hello workd',
    ],
    4 => [
        'imagepath' => __DIR__ . '/test.png',
        'loop'      => [
            ['color' => '00FF00', 'imagepath' => __DIR__ . '/test1.png'],
            ['color' => '0000FF', 'imagepath' => __DIR__ . '/test2.png'],
        ],
    ],
]);

rename($outfile, __DIR__ . '/template-out.xlsx');
