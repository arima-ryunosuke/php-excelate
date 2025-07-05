<?php

if (getenv('PHPVERSION')) {
    require_once __DIR__ . '/versions/' . getenv('PHPVERSION') . '/vendor/autoload.php';
}
require_once __DIR__ . '/../vendor/autoload.php';
