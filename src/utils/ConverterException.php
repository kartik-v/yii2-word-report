<?php

/**
 * @package   yii2-word-report
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2021
 * @version   1.0.0
 */
namespace kartik\wordreport\utils;

use yii\base\Exception;

/**
 * ConverterException represents an exception caused due to error in document format conversion.
 */
class ConverterException extends Exception
{
    /**
     * @return string the user-friendly name of this exception
     */
    public function getName()
    {
        return 'Converter Exception';
    }
}