<?php

namespace ryunosuke\Excelate;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * 埋め込みテンプレートをレンダリングするクラス
 */
class Renderer
{
    // エラーモード定数
    const ERROR_MODE_DEFAULT   = 0; // 余計なことは一切しないで生のままの実行する
    const ERROR_MODE_SILENT    = 1; // 一切の報告をしない。ログにも出さないで握りつぶす
    const ERROR_MODE_RENDERING = 2; // エラー文字列を変数値としてエクセルに埋め込む
    const ERROR_MODE_WARNING   = 3; // 通常の WARNING として発生させる
    const ERROR_MODE_EXCEPTION = 4; // 例外として送出する

    /** @var Cell */
    private $currentCell;

    /** @var mixed[] */
    private $variables = [];

    /** @var callable[] */
    private $effectors = [];

    /** * @var int */
    private $errorMode = self::ERROR_MODE_DEFAULT;

    public function __construct()
    {
        // 別シートセルへのリンクを貼る effector
        $this->registerEffector('Link', function (Cell $cell, $link) {
            $cell->getHyperlink()->setUrl("sheet://$link");
        });
        // ハイパーリンクを貼る effector
        $this->registerEffector('HyperLink', function (Cell $cell, $url, $value = null) {
            $cell->getHyperlink()->setUrl($url);
            return $value === null ? $url : $value;
        });
        // 文字色を変える effector
        $this->registerEffector('Color', function (Cell $cell, $color) {
            $cell->getStyle()->getFont()->setColor(new Color($color));
        });
        // 罫線をつける effector
        $this->registerEffector('Border', function (Cell $cell, $border) {
            if (!$border) {
                return;
            }
            switch (count($border)) {
                case 1:
                    $border = [$border[0], $border[0], $border[0], $border[0]];
                    break;
                case 2:
                    $border = [$border[0], $border[1], $border[0], $border[1]];
                    break;
                case 3:
                    $border = [$border[0], $border[1], $border[2], $border[1]];
                    break;
                case 4:
                    $border = [$border[0], $border[1], $border[2], $border[3]];
                    break;
            }
            $style = $cell->getWorksheet()->getStyle($cell->getCoordinate());
            foreach (['top', 'right', 'bottom', 'left'] as $n => $pos) {
                /** @var Border $b */
                $b = $style->getBorders()->{"get$pos"}();
                if (isset($border[$n][0])) {
                    $b->setBorderStyle($border[$n][0]);
                }
                if (isset($border[$n][1])) {
                    $b->getColor()->setRGB($border[$n][1]);
                }
            }
        });
        // 入力規則 effector @see https://phpspreadsheet.readthedocs.io/en/latest/topics/recipes/#setting-data-validation-on-a-cell
        $validation = function (Cell $cell, $attrs) {
            // array_change_key_case が遅いわけではないが超頻繁に呼ばれる可能性があるので無駄なことはしない
            if (!isset($attrs['__from__'])) {
                $attrs = array_change_key_case($attrs);
            }
            unset($attrs['__from__']);

            // メッセージ系の簡易設定（指定されていたら自動で true にしたり）
            $types = [
                'prompt' => ['switch' => 'showinputmessage', 'message' => 'prompt', 'title' => 'prompttitle', 'style' => null],
                'error'  => ['switch' => 'showerrormessage', 'message' => 'error', 'title' => 'errortitle', 'style' => DataValidation::STYLE_STOP],
                'warn'   => ['switch' => 'showerrormessage', 'message' => 'error', 'title' => 'errortitle', 'style' => DataValidation::STYLE_WARNING],
                'info'   => ['switch' => 'showerrormessage', 'message' => 'error', 'title' => 'errortitle', 'style' => DataValidation::STYLE_INFORMATION],
            ];
            foreach ($types as $type => $config) {
                if (isset($attrs[$type])) {
                    $messages = ((array) $attrs[$type]) + [1 => ''];
                    $attrs[$config['switch']] = true;
                    $attrs[$config['message']] = $messages[0];
                    $attrs[$config['title']] = $messages[1];
                    if ($config['style']) {
                        $attrs['errorstyle'] = $config['style'];
                    }
                    unset($attrs[$type]);
                }
            }

            // デフォルト系（未設定のみ）
            $attrs['showerrormessage'] ??= true; // エラーを出さないと入力規則の旨味がほとんどない
            if ($attrs['type'] === DataValidation::TYPE_LIST) {
                $attrs['showdropdown'] ??= true; // ドロップダウンを出さないリストはほとんどない
            }

            $validation = $cell->getDataValidation();
            foreach ($attrs as $name => $value) {
                $validation->{"set$name"}(...(array) $value);
            }
        };
        $this->registerEffector('Validation', $validation); // 汎用
        $this->registerEffector('ValidationList', function (Cell $cell, $attrs) use ($validation) { // 経験上、最も多い入力規則はリスト型
            // デフォルト設定で良ければ単にリストを与えた場合に選択肢とする
            $indexarray = true;
            foreach ($attrs as $name => $dummy) {
                if (is_string($name)) {
                    $indexarray = false;
                    break;
                }
            }
            if ($indexarray) {
                $attrs = ['formula1' => $attrs];
            }

            $attrs = array_change_key_case($attrs);
            $attrs['__from__'] = DataValidation::TYPE_LIST;
            $attrs['type'] = DataValidation::TYPE_LIST;
            if (is_array($attrs['formula1'])) {
                $attrs['formula1'] = '"' . implode(',', array_map(fn($v) => strtr($v, ['"' => '""']), $attrs['formula1'])) . '"';
            }
            return $validation($cell, $attrs);
        });
        // 画像を埋め込む effector
        $this->registerEffector('Image', function (Cell $cell, $attrs) {
            if (is_string($attrs)) {
                $attrs = ['path' => $attrs];
            }
            $drawing = new Drawing();
            $drawing->setPath($attrs['path']);
            $drawing->setCoordinates($cell->getCoordinate());
            $keys = [
                'name'               => [],
                'description'        => [],
                'resizeProportional' => [],
                'offsetX'            => [],
                'offsetY'            => [],
                'width'              => [],
                'height'             => [],
                'rotation'           => [],
            ];
            // sizeToFit だけ特別扱いで処理する（width/height を超えないようにリサイズする）
            if (isset($attrs['sizeToFit'], $attrs['width'], $attrs['height']) && $attrs['sizeToFit']) {
                $ratio = $drawing->getWidth() / $drawing->getHeight();
                if ($ratio > $attrs['width'] / $attrs['height']) {
                    $width = $attrs['width'];
                    $height = intval($attrs['width'] / $ratio);
                }
                else {
                    $width = intval($attrs['height'] * $ratio);
                    $height = $attrs['height'];
                }
                $drawing->setResizeProportional(false);
                $drawing->setWidth($width);
                $drawing->setHeight($height);
                unset($keys['resizeProportional'], $keys['width'], $keys['height']);
            }
            foreach ($keys as $key => $arg) {
                if (array_key_exists($key, $attrs)) {
                    $drawing->{"set$key"}($attrs[$key]);
                }
            }
            $drawing->setWorksheet($cell->getWorksheet());
        });
    }

    /**
     * Variable を登録
     *
     * registerVariable('Hoge', 'hoge'){});のように登録するとテンプレート側で {$Hoge} のように使用できるようになる。
     *
     * @param string $name
     * @param mixed $variable
     */
    public function registerVariable($name, $variable)
    {
        $this->variables[$name] = $variable;
    }

    /**
     * Effector を登録
     *
     * registerEffector('Hoge', function($cell, $arg1){});のように登録するとテンプレート側で $Hoge('arg1') のように呼べるようになる。
     *
     * @param string $name
     * @param callable $effector
     */
    public function registerEffector($name, $effector)
    {
        // テンプレート側で $cell を指定しなくて済むようにラップする
        $this->effectors[$name] = function () use ($effector) {
            $args = func_get_args();
            array_unshift($args, $this->currentCell);
            return call_user_func_array($effector, $args);
        };
    }

    public function setErrorMode($errorMode)
    {
        $valids = [
            self::ERROR_MODE_DEFAULT,
            self::ERROR_MODE_SILENT,
            self::ERROR_MODE_WARNING,
            self::ERROR_MODE_RENDERING,
            self::ERROR_MODE_EXCEPTION,
        ];
        if (!in_array($errorMode, $valids)) {
            throw new \InvalidArgumentException("$errorMode is invalid error mode.");
        }
        $this->errorMode = $errorMode;
    }

    public function renderBook(string $filename, array $sheetsVars, callable $done = null)
    {
        $typeMap = [
            'xlsx' => 'Xlsx',
            'xlsm' => 'Xlsx',
            'xltx' => 'Xlsx',
            'xltm' => 'Xlsx',
            'xls'  => 'Xls',
            'xlt'  => 'Xls',
            'ods'  => 'Ods',
            'htm'  => 'Html',
            'html' => 'Html',
            'csv'  => 'Csv',
        ];

        $extension = pathinfo($filename, PATHINFO_EXTENSION);
        $type = $typeMap[strtolower($extension)];

        $book = IOFactory::createReader($type)->load($filename);

        foreach ($sheetsVars as $eitherNameOrIndex => $vars) {
            if ($eitherNameOrIndex === "") {
                $sheet = $book->getActiveSheet();
            }
            elseif (is_string($eitherNameOrIndex)) {
                $sheet = $book->getSheetByName($eitherNameOrIndex);
            }
            else {
                $sheet = $book->getSheet($eitherNameOrIndex);
            }

            if ($sheet !== null) {
                $this->renderSheet($sheet, $vars);
            }
        }

        if ($done) {
            $done($book);
        }

        $tmpfile = tempnam(sys_get_temp_dir(), 'excelate');
        IOFactory::createWriter($book, $type)->save($tmpfile);
        return $tmpfile;
    }

    public function renderSheet(Worksheet $sheet, $vars, $range = null)
    {
        $this->currentCell = $sheet->getCell('A1');

        $title = $sheet->getTitle();
        $this->parse($title, $vars);
        $sheet->setTitle($title);

        if ($range === null) {
            $cell = $sheet->getCell('A1');
            $cellvalue = $cell->getValue();
            $tokens = $this->parse($cellvalue, $vars);
            foreach ($tokens as $token) {
                if (is_array($token)) {
                    switch ($token['type']) {
                        case 'template':
                            $range = $token['args']['range'];
                            break;
                    }
                }
            }
        }
        if ($range === null) {
            $highests = $sheet->getHighestRowAndColumn();
            $highests['row']++;
            $highests['column']++;
            $range = 'A1:' . $highests['column'] . $highests['row'];
        }

        [$lt, $rb] = Coordinate::rangeBoundaries($range);
        [$left, $top] = $lt;
        [$right, $bottom] = $rb;

        error_clear_last();
        return $this->_render($sheet, $vars, $left, $top, $right, $bottom);
    }

    private function _render(Worksheet $sheet, $vars, $left, $top, $right, $bottom)
    {
        $left = (int) $left;
        $top = (int) $top;
        $right = (int) $right;
        $bottom = (int) $bottom;

        $initRight = $right;
        $initBottom = $bottom;

        $nest = 0;
        $if = [];
        $foreach = [];

        for ($row = $top; $row <= $bottom; $row++) {
            for ($col = $left; $col <= $right; $col++) {
                $cell = $sheet->getCell([$col, $row]);
                $cellvalue = $cellvalue2 = $cell->getValue();
                $this->currentCell = $cell;
                $tokens = $this->parse($cellvalue, $vars, $nest);
                if ($cellvalue !== $cellvalue2) {
                    $cell->setValue($cellvalue);
                }

                foreach ($tokens as $token) {
                    switch ($token['type']) {
                        case 'template':
                            if (!($row === 1 && $col === 1)) {
                                throw new \DomainException('{template} tag is permitted only A1 cell.');
                            }
                            break;
                        case 'row':
                            $varss = $this->placeholder($token['args']['values'], $vars);
                            foreach (array_values($varss) as $dc => $value) {
                                $sheet->getCell([$col + $dc, $row])->setValue($value)->setXfIndex($cell->getXfIndex());
                            }
                            break;
                        case 'rowcol':
                            $varss = $this->placeholder($token['args']['values'], $vars);
                            if ($varss) {
                                if ($token['args']['header']) {
                                    $varss = array_merge([array_keys(reset($varss))], $varss);
                                }

                                $varslength = count($varss);

                                if ($varslength > 1) {
                                    $sheet->insertNewRowBefore($row + 1, $varslength - 1);
                                }

                                foreach (array_values($varss) as $dr => $cols) {
                                    foreach (array_values($cols) as $dc => $value) {
                                        $sheet->getCell([$col + $dc, $row + $dr])->setValue($value);
                                    }
                                }

                                $bottom += $varslength;
                            }
                            break;
                        case 'rowif':
                        case 'colif':
                        case 'if':
                            $if = $token['args'] + [
                                    'left' => $col,
                                    'top'  => $row,
                                ];
                            break;
                        case '/rowif':
                        case '/colif':
                        case '/if':
                            $varss = $this->placeholder($if['cond'], $vars);
                            if ($varss) {
                                [$dCol, $dRow] = $this->_render($sheet, $vars, $if['left'], $if['top'], $col, $row);
                                $right += $dCol;
                                $bottom += $dRow;
                                break;
                            }

                            $ifmode = $token['type'] === '/if';
                            Utils::unmergeCells($sheet, $if['left'], $if['top'], $col, $row);
                            if (strpos($token['type'], 'row') !== false || $ifmode) {
                                if (($if['left'] === $left && $col === $right) || $ifmode) {
                                    $delta = $row - $if['top'] + 1;
                                    $sheet->removeRow($if['top'], $delta);
                                    $bottom -= $delta;
                                    $row -= $delta;
                                }
                                else {
                                    $target = max($sheet->getHighestRow(), $row + 1 + $row - $if['top'] + 1);
                                    Utils::copyCells($sheet, $if['left'], $row + 1, $col, $target, null, $if['top']);
                                }
                            }
                            else {
                                if (($if['top'] === $top && $row === $bottom)) {
                                    $delta = $col - $if['left'] + 1;
                                    $sheet->removeColumnByIndex($if['left'], $delta);
                                    $right -= $delta;
                                    $col -= $delta;
                                }
                                else {
                                    $target = max(Coordinate::columnIndexFromString($sheet->getHighestColumn()), $col + 1 + $col - $if['left'] + 1);
                                    Utils::copyCells($sheet, $col + 1, $if['top'], $target, $row, $if['left'], null);
                                }
                            }
                            break;
                        case 'roweach':
                        case 'coleach':
                        case 'rowshift':
                        case 'colshift':
                            $foreach = $token['args'] + [
                                    'left' => $col,
                                    'top'  => $row,
                                ];
                            break;
                        case '/roweach':
                        case '/coleach':
                        case '/rowshift':
                        case '/colshift':
                            $rowmode = strpos($token['type'], 'row') !== false;
                            $eachmode = strpos($token['type'], 'each') !== false;

                            $varss = $this->placeholder($foreach['values'], $vars);
                            $varslength = count($varss);

                            if ($rowmode) {
                                if ($eachmode) {
                                    $delta = Utils::insertDuplicateRows($sheet, $foreach['left'], $foreach['top'], $col, $row, $row + 1, $varslength - 1);
                                }
                                else {
                                    $delta = Utils::shiftDuplicateRows($sheet, $foreach['left'], $foreach['top'], $col, $row, $varslength - 1, $bottom);
                                }
                                $bottom += $delta;
                                $dimRow = $row - $foreach['top'] + 1;
                                $dimCol = 0;
                            }
                            else {
                                if ($eachmode) {
                                    $delta = Utils::insertDuplicateCols($sheet, $foreach['left'], $foreach['top'], $col, $row, $col + 1, $varslength - 1);
                                }
                                else {
                                    $delta = Utils::shiftDuplicateCols($sheet, $foreach['left'], $foreach['top'], $col, $row, $varslength - 1, $right);
                                }
                                $right += $delta;
                                $dimRow = 0;
                                $dimCol = $col - $foreach['left'] + 1;
                            }

                            $dc = 0;
                            $dr = 0;
                            $n = 0;
                            foreach ($varss as $k => $var) {
                                $context = [
                                    $foreach['k'] => $k,
                                    $foreach['v'] => $var,
                                    'index'       => $n,
                                    'first'       => $n === 0,
                                    'last'        => $n === $varslength - 1,
                                ];
                                $context += (array) $var + (array) $vars;
                                $l = $foreach['left'] + $dc;
                                $t = $foreach['top'] + $dr;
                                $r = $col + $dc;
                                $b = $row + $dr;
                                [$dCol, $dRow] = $this->_render($sheet, $context, $l, $t, $r, $b);

                                $dc += $dCol + $dimCol;
                                $dr += $dRow + $dimRow;
                                $bottom += $dRow;
                                $right += $dCol;
                                $n++;
                            }
                            break;
                    }
                }
            }
        }

        return [$right - $initRight, $bottom - $initBottom];
    }

    private function parse(&$cellvalue, $vars, &$nest = 0)
    {
        if (strlen($cellvalue ?? '') === 0) {
            return [];
        }
        $tokens = [];
        preg_match_all('#(\{ (?: [^{}]+ | (?1) )* \})|([^{}]+)#x', $cellvalue, $m);
        $cellvalue = '';
        foreach ($m[0] as $token) {
            if ($token[0] === '{') {
                $_token = substr($token, 1, strlen($token) - 2);
                if (preg_match('#^(template)\s+([a-zA-Z]+[0-9]+:[a-zA-Z]+[0-9]+)$#', $_token, $matches)) {
                    $tokens[] = [
                        'type' => $matches[1],
                        'args' => [
                            'range' => $matches[2],
                        ],
                    ];
                }
                elseif (preg_match('#^(row)\s+([^:}]+)$#', $_token, $matches)) {
                    if ($nest === 0) {
                        $tokens[] = [
                            'type' => $matches[1],
                            'args' => [
                                'values' => trim($matches[2]),
                            ],
                        ];
                    }
                    else {
                        $cellvalue .= $token;
                    }
                }
                elseif (preg_match('#^(rowcol)\s+([^:}]+):?(true|false)?$#', $_token, $matches)) {
                    if ($nest === 0) {
                        $tokens[] = [
                            'type' => $matches[1],
                            'args' => [
                                'values' => trim($matches[2]),
                                'header' => filter_var(trim($matches[3] ?? false), FILTER_VALIDATE_BOOLEAN),
                            ],
                        ];
                    }
                    else {
                        $cellvalue .= $token;
                    }
                }
                elseif (preg_match('#^(rowif|colif|if)\s+([^}]+)$#', $_token, $matches)) {
                    if ($nest++ === 0) {
                        $tokens[] = [
                            'type' => $matches[1],
                            'args' => [
                                'cond' => $matches[2],
                            ],
                        ];
                    }
                    else {
                        $cellvalue .= $token;
                    }
                }
                elseif (preg_match('#^(/rowif|/colif|/if)$#', $_token, $matches)) {
                    if (--$nest === 0) {
                        $tokens[] = [
                            'type' => $matches[1],
                            'args' => [],
                        ];
                    }
                    else {
                        $cellvalue .= $token;
                    }
                }
                elseif (preg_match('#^(roweach|coleach|rowshift|colshift)\s+(\$[^}\s]+)(\s+(\$[^:}]+)(:(\$[^}]+))?)?$#', $_token, $matches)) {
                    if ($nest++ === 0) {
                        $matches += [4 => '$k', 6 => '$v'];
                        $tokens[] = [
                            'type' => $matches[1],
                            'args' => [
                                'values' => trim($matches[2]),
                                'k'      => ltrim(trim($matches[4]), '$'),
                                'v'      => ltrim(trim($matches[6]), '$'),
                            ],
                        ];
                    }
                    else {
                        $cellvalue .= $token;
                    }
                }
                elseif (preg_match('#^(/roweach|/coleach|/rowshift|/colshift)$#', $_token, $matches)) {
                    if (--$nest === 0) {
                        $tokens[] = [
                            'type' => $matches[1],
                            'args' => [],
                        ];
                    }
                    else {
                        $cellvalue .= $token;
                    }
                }
                elseif (preg_match('#^(ldelim|rdelim)$#', $_token, $matches)) {
                    if ($matches[1] === 'ldelim') {
                        $cellvalue .= '{';
                    }
                    else {
                        $cellvalue .= '}';
                    }
                }
                else {
                    if ($nest === 0) {
                        $cellvalue .= $this->placeholder($_token, $vars);
                    }
                    else {
                        $cellvalue .= $token;
                    }
                }
            }
            else {
                $cellvalue .= $token;
            }
        }

        return $tokens;
    }

    private function placeholder($statement)
    {
        $context = [];
        $context += $this->effectors;
        $context += $this->variables;
        foreach (array_slice(func_get_args(), 1) as $ctx) {
            $context += (array) $ctx;
        }

        $render = function () {
            extract(func_get_arg(0));
            return eval("return " . func_get_arg(1) . ";");
        };

        // eval
        if ($this->errorMode === self::ERROR_MODE_DEFAULT) {
            return $render($context, $statement);
        }
        $current = error_get_last();
        $return = @$render($context, $statement);
        $error = error_get_last();
        if ($current !== $error) {
            switch ($this->errorMode) {
                case self::ERROR_MODE_SILENT;
                    break;
                case self::ERROR_MODE_RENDERING;
                    return $error['message'];
                case self::ERROR_MODE_WARNING;
                    trigger_error($statement, E_USER_WARNING);
                    break;
                case self::ERROR_MODE_EXCEPTION;
                    throw new \ErrorException($error['message'], 0, $error['type'], $error['file'], $error['line']);
            }
        }
        return $return;
    }
}
