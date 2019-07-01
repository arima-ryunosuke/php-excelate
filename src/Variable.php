<?php

namespace ryunosuke\Excelate;

/**
 * 配列のようにもオブジェクトのようにもアクセスできるクラス
 */
class Variable implements \ArrayAccess, \Countable
{
    /**
     * コールバックを適用してこのクラスの配列を作成
     *
     * 仕様上、値だけを指定してループしたいことが多いので、キーや連番を包含するような配列が作れると便利。
     *
     * @param array $array
     * @param callable $callback
     * @return self[]
     */
    public static function arrayize($array, $callback)
    {
        $result = [];
        $n = 0;
        foreach ($array as $key => $value) {
            $row = $callback($value, $n, $key);
            if (is_array($row)) {
                $result[] = new self($row);
                $n++;
            }
        }
        return $result;
    }

    public function __construct($source)
    {
        foreach ($source as $key => $value) {
            $this->$key = $value;
        }
    }

    public function offsetExists($offset)
    {
        return isset($this->$offset);
    }

    public function offsetGet($offset)
    {
        return $this->$offset;
    }

    public function offsetSet($offset, $value)
    {
        $this->$offset = $value;
    }

    public function offsetUnset($offset)
    {
        unset($this->$offset);
    }

    public function count()
    {
        return count((array) $this);
    }

    public function __toString()
    {
        return (string) print_r($this, 1);
    }
}
