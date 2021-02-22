<?php

/**
 * @package   yii2-word-report
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2021
 * @version   1.0.0
 */

namespace kartik\wordreport\pdf;

use kartik\wordreport\utils\Converter;
use Yii;

/**
 * Libre Office PDF conversion library
 * @package kartik\wordreport
 */
class LibreOffice extends Converter
{
    /**
     * @var string the name of the binary
     */
    public $binary = '/usr/bin/libreoffice';

    /**
     * @var string the name of the profile folder location
     */
    public $profile = '/tmp/kv-pdf-libreoffice';

    /**
     * Convert the document to PDF using `libreoffice` executable
     */
    public function convert()
    {
        $this->validateArgs();
        $this->args = ['--headless' => ''];
        if (!empty($this->profile)) {
            $profile = trim(Yii::getAlias($this->profile), "/\\");
            $this->args['-env:UserInstallation='] = "file://{$profile}";
        }
        $this->args['--convert-to'] = 'pdf:writer_pdf_Export';
        $this->args['--outdir'] = Yii::getAlias($this->output);
        $this->args[' '] = Yii::getAlias($this->input);
        $this->runBinary();
    }
}