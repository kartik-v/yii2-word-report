<?php

/**
 * @package   yii2-word-report
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2021
 * @version   1.0.0
 */

namespace kartik\wordreport\utils;

use Yii;
use yii\base\BaseObject;

/**
 * Abstract Converter class manages Document format conversions
 * @package kartik\wordreport
 */
abstract class Converter extends BaseObject
{
    use LogTrait;

    /**
     * @var string the name of the binary
     */
    public $binary;

    /**
     * @var array the arguments to be passed to command as key value pairs
     */
    public $args;

    /**
     * @var string the input file path
     */
    public $input;

    /**
     * @var string the output file path
     */
    public $output;

    /**
     * @var string the user profile directory on the server
     */
    public $profile;

    /**
     * @var Command
     */
    protected $_command;

    /**
     * Convert the document to the target format
     */
    abstract public function convert();

    /**
     * Runs the binary command on the server
     */
    public function runBinary()
    {
        $command = $this->getCommand();
        if ($command->execute()) {
            static::log('runBinary(): Success: ' . $command->getOutput());
            return;
        }
        $code = $command->getExitCode();
        $error = $command->getError();
        static::log("runBinary(): Error [{$code}]: {$error}");
        throw new ConverterException($error);
    }

    /**
     * @return Command
     */
    public function getCommand()
    {
        if (empty($this->_command)) {
            $this->setCommand();
        }
        return $this->_command;
    }

    /**
     * Sets command
     * @param array $config the command configuration
     */
    public function setCommand($config = [])
    {
        if (empty($config['command'])) {
            $config['command'] = $this->binary;
        }
        $this->_command = new Command($config);
        foreach ($this->args as $key => $value) {
            $this->_command->addArg($key, $value);
        }
    }

    /**
     * Validation of input file, output file and profile location
     * @throws ConverterException
     */
    protected function validateArgs()
    {
        if (empty($this->input)) {
            throw new ConverterException("Input file not provided");
        }
        $input = Yii::getAlias($this->input);
        if (!is_readable($input)) {
            throw new ConverterException("Input file '{$input}' is not readable!");
        }
        if (empty($this->output)) {
            throw new ConverterException("Output file not provided");
        }
        if (!empty($this->profile)) {
            $profile = Yii::getAlias($this->profile);
            if (!is_dir($profile)) {
                $out = mkdir($profile, 0777, true);
            }
            if (!$out || !is_writable($profile)) {
                $name = $this->getCommand()->getCommand();
                $error = "'{$name}' does not have permissions to the user profile directory ('{$this->profile}')!";
                throw new ConverterException($error);
            }
        }
    }
}