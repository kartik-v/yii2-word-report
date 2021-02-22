<?php


namespace kartik\wordreport\utils;

use yii\console\Application as ConsoleApplication;
use Yii;

trait LogTrait
{
    /**
     * @param string $message
     */
    public static function log($message)
    {
        $classPart = explode('\\', static::class);
        $class = end($classPart);
        $msg = "{$class}::{$message}";
        Yii::debug("[DEBUG]: {$msg}", 'word-report');
        if (Yii::$app instanceof ConsoleApplication) {
            $now = date('Y-m-d H:i:s');
            echo "[$now]: {$msg}\n";
        }
    }

}