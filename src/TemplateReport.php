<?php

/**
 * @package   yii2-word-report
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2021
 * @version   1.0.0
 */

namespace kartik\wordreport;

use DOMDocument;
use kartik\wordreport\pdf\LibreOffice;
use kartik\wordreport\utils\Converter;
use kartik\wordreport\utils\ConverterException;
use kartik\wordreport\utils\LogTrait;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\Exception\Exception as PhpWordException;
use yii\base\Component;
use Yii;
use yii\base\InvalidConfigException;

/**
 * TemplateReport class is a library for generating reports in Word/PDF from Microsoft Word Templates.
 *
 * Usage:
 * ```
 * // initialize report
 * $report = new TemplateReport([
 *    'format' => TemplateReport::FORMAT_BOTH,
 *    'inputFile' => 'Invoice_Template_01.docx',
 *    'outputFile' => 'Invoice_Report_' . date('Y-m-d'),
 *    'values' => ['invoice_no' => 2001, 'invoice_date' => '2020-02-21'],
 *    'images' => ['company_logo' => '@webroot/images/company.jpg', 'customer_logo' => '@webroot/images/company.jpg'],
 *    'rows' => [
 *      'item' => [
 *          ['item' => 1, 'name' => 'Potato', 'price' => '$10.00'],
 *          ['item' => 2, 'name' => 'Tomato', 'price' => '$20.00'],
 *      ]
 *    ],
 *    'blocks' => [
 *      'customer_block' => [
 *          ['customer_name' => 'John', 'customer_address' => 'Address for John'],
 *          ['customer_name' => 'Bill', 'customer_address' => 'Address for Bill'],
 *      ],
 *    ]
 * ]);
 * // Generate the report
 * $report->generate();
 * ```
 * @author Kartik Visweswaran <kartikv2@gmail.com>
 * @since  1.0
 */
class TemplateReport extends Component
{
    use LogTrait;

    const FORMAT_WORD = 'docx';
    const FORMAT_PDF = 'pdf';
    const FORMAT_BOTH = 'both';

    /**
     * @var string the output format whether 'docx' or 'pdf' or 'both'
     */
    public $format = self::FORMAT_WORD;

    /**
     * @var Converter the pdf converter class to use - defaults to 'kartik\wordreport\pdf\LibreOffice' if not provided
     */
    public $pdfConverter;

    /**
     * @var string the path where the Microsoft Word template files exist
     */
    public $inputPath = '@webroot/templates';

    /**
     * @var string the name of the template file with the extension (docx, odt, rtf)
     */
    public $inputFile;

    /**
     * @var string the path where the Microsoft Word output files will be saved
     */
    public $outputPath = '@webroot/reports';

    /**
     * @var string the name of the output file without the PDF or DOCX extension
     */
    public $outputFile;

    /**
     * @var string path to any custom XSL style sheet to apply if applicable
     */
    public $xslStyleSheet;

    /**
     * @var array key value pairs of template variable names and their values which will be replaced. This must be
     * a two dimensional array (which means a variable value must be string/int and cannot be an array).
     * ```
     * // template setting
     * // Invoice No.: ${invoice_number}
     * // Invoice Dt: ${invoice_date}
     *
     * 'values' => [
     *      'invoice_number' => '1000',
     *      'invoice_date' => '22-Feb-2021',
     * ]
     * ```
     */
    public $values;

    /**
     * @var array key value pairs of template image variable names and their values which will be replaced. The value
     * must be full path to the image on the server.
     *
     * ```
     * // template setting
     * // ${company_logo}
     * // Company Name
     *
     * 'images' => [
     *      'company_logo' => '@webroot/uploads/company/logo.jpg'
     * ]
     * ```
     */
    public $images;

    /**
     * @var array key value pairs of template variable names in a table in the document and their values which will be
     * replaced. Each value must be an array of  [[values]] which will be repeated for number of items in the
     * array. The key should typically match the first table column's variable name.
     *
     * ```
     * // template setting
     * // ----------+-------------+-----------
     * // ${item}   |    {$name}  | ${price}
     * // ----------+-------------+-----------
     * 'rows' => [
     *      'item' => [
     *          [
     *               'item' => 1,
     *               'name' => 'Potato',
     *               'price' => '$10.00',
     *          ],
     *          [
     *               'item' => 2,
     *               'name' => 'Tomato',
     *               'price' => '$20.00',
     *          ],
     *      ]
     * ]
     * ```
     */
    public $rows;

    /**
     * @var array repeating block of values - similar to tabular values but with enhanced control.
     * ```
     * // template setting
     * // ${customer_block}
     * // Customer: ${customer_name}
     * // Address: ${customer_address}
     * // ${/customer_block}
     * // ${block_2}
     * // Some text inside the block
     * // @{/block_2}
     * // ${block_3}
     * // @{/block_3}
     *
     * 'blocks' => [
     *      // loops array and replaces variables inside the block
     *      'customer_block' => [
     *          [
     *               'customer_name' => 'John',
     *               'customer_address' => 'Address for John',
     *          ],
     *          [
     *               'customer_name' => 'Bill',
     *               'customer_address' => 'Address for Bill',
     *          ],
     *      ],
     *
     *      // replaces the block with the block text mentioned
     *      'block_2' => 'Replace this text inside the block',
     *
     *      // deletes the block if the block variable is set to false
     *      'block_3' => false
     * ]
     * ```
     */
    public $blocks;

    /**
     * @var array key value pairs of template variable chart names and their values which will be replaced. This must be
     * a two dimensional array of chart variable names as keys and the chart object as the value. For example:
     *
     * ```
     * // template setting
     * // CHART 1:
     * // ${chart1}
     * // CHART 2:
     * // ${chart2}
     *
     * use PhpOffice\PhpWord\Element\Chart;
     * $categories = array('A', 'B', 'C', 'D', 'E');
     * $series1 = array(1, 3, 2, 5, 4);
     * $series2 = array(3, 1, 7, 2, 6);
     * $series3 = array(8, 3, 2, 5, 4);
     *
     * $chart1 = new Chart('pie', $categories, $series1);
     * $chart2 = new Chart('bar', $categories, $series1);
     * $chart2->addSeries($series2);
     * $chart3->addSeries($series2);
     *
     * 'charts' => [
     *      'chart1' => $chart1,
     *      'chart2' => $chart2,
     * ]
     * ```
     */
    public $charts;

    /**
     * @var array key value pairs of template variable names and their values which will be replaced. This must be
     * a two dimensional array of variable names as keys and the complex object as the value. For example:
     *
     * ```
     * // template setting
     * // Content: ${inline}
     *
     *  $object = new \PhpOffice\PhpWord\Element\TextRun();
     *  $object->addText('by a red italic text', ['italic' => true, 'color' => 'red']);
     * 'complexValues' => [
     *      'inline' => $object,
     * ]
     * ```
     */
    public $complexValues;

    /**
     * @var array key value pairs of template variable block names and their values which will be replaced. This must
     *     be
     * a two dimensional array of block variable names as keys and the complex object as the value. For example:
     *
     * ```
     * // template setting
     * // ${table}
     * // ${/table}
     *
     * $table = new \PhpOffice\PhpWord\Element\Table(['borderSize' => 12, 'borderColor' => 'green', 'width' => 6000,
     *     'unit' => TblWidth::TWIP]);
     * $table->addRow();
     * $table->addCell(150)->addText('Cell A1');
     * $table->addCell(150)->addText('Cell A2');
     * $table->addRow();
     * $table->addCell(150)->addText('Cell B1');
     * $table->addCell(150)->addText('Cell B2');
     *
     * 'complexBlocks' => [
     *      'table' => $table,
     * ]
     * ```
     */
    public $complexBlocks;

    /**
     * @var TemplateProcessor the PHP Template Processor instance
     */
    private $_template;

    /**
     * Set the template processor object
     * @throws CopyFileException|CreateTemporaryFileException|InvalidConfigException|PhpWordException
     */
    public function setTemplate()
    {
        $path = Yii::getAlias($this->inputPath . '/' . $this->inputFile);
        if (!file_exists($path)) {
            throw new InvalidConfigException("Invalid input template file or path '{$path}'.");
        }
        $this->_template = new TemplateProcessor($path);
        static::log("setTemplate(): Template input read from '{$path}'.");
    }

    /**
     * Get the template processor object
     * @return TemplateProcessor
     */
    public function getTemplate()
    {
        return $this->_template;
    }

    /**
     * Generates the word report output
     * @throws CopyFileException|CreateTemporaryFileException|InvalidConfigException|PhpWordException|ConverterException
     */
    public function generate()
    {
        $doc = self::FORMAT_WORD;
        $pdf = self::FORMAT_PDF;
        $both = self::FORMAT_BOTH;
        if ($this->format !== $doc && $this->format !== $pdf && $this->format !== $both) {
            throw new InvalidConfigException("Format must be either '{$doc}' or '{$pdf}' or '{$both}'");
        }
        $this->setTemplate();
        $template = $this->getTemplate();
        static::log("generate(): Begin report output generation");
        if (!empty($this->xslStyleSheet)) {
            $path = Yii::getAlias($this->xslStyleSheet);
            if (file_exists($path)) {
                $xslDomDocument = new DOMDocument();
                $xslDomDocument->load($path);
                $template->applyXslStyleSheet($xslDomDocument);
                static::log("generate(): Applied xslStyleSheet '{$this->xslStyleSheet}'");
            }
        }
        $this->process('values');
        $this->process('images');
        $this->process('blocks');
        $this->process('charts');
        $this->process('rows', 'cloneRowAndSetValues');
        $this->process('complexValues');
        $this->process('complexBlocks');
        $file = substr($this->outputFile, 0, strrpos($this->outputFile, "."));
        $output = Yii::getAlias("{$this->outputPath}/{$file}.{$doc}");
        $template->saveAs($output);
        static::log("generate(): Generated output file '{$file}.{$doc}'");
        if ($this->format !== $doc) {
            $this->generatePdf($output, Yii::getAlias("{$this->outputPath}/{$file}.{$pdf}"));
            static::log("generate(): Generated output file '{$file}.{$pdf}'");
            if ($this->format === $pdf) {
                unlink($output);
                static::log("generate(): Deleted file '{$file}.{$doc}'");
            }
        }
        static::log("generate(): End report output generation");
    }

    /**
     * Processes output by executing appropriate template method for the selected property
     * @param string $prop
     * @param string $method
     */
    public function process($prop, $method = null)
    {
        if (empty($this->$prop)) {
            return;
        }
        $function = $method === null ? 'set' . ucfirst(rtrim($prop, 's')) : $method;
        static::log("process('{$prop}'): Processing template '{$function}' method");
        $t = $this->getTemplate();
        if ($prop === 'values') {
            $t->setValues($this->values);
            return;
        }
        foreach ($this->$prop as $key => $value) {
            if ($prop === 'images') {
                $path = Yii::getAlias($value);
                if (file_exists($path)) {
                    $t->setImageValue($key, $path);
                }
            } elseif ($prop === 'blocks') {
                if ($value === false) {
                    $t->deleteBlock($key);
                } elseif (is_array($value)) {
                    $t->cloneBlock($key, 0, true, false, $value);
                } else {
                    $t->replaceBlock($key, $value);
                }
            } else {
                $t->$function($key, $value);
            }
        }
    }

    /**
     * Generate a PDF file from input document file using libreoffice host command
     * @param string $input the input document file
     * @param string $output the output document file
     * @throws ConverterException
     */
    public function generatePdf($input, $output)
    {
        if (empty($this->pdfConverter)) {
            $this->pdfConverter = LibreOffice::class;
        }
        $class = $this->pdfConverter;
        if (!class_exists($class)) {
            throw new ConverterException("PDF Converter class '{$class}' does not exist.");
        }
        $converter = new $class(['input' => $input, 'output' => $output]);
        if (!$converter instanceof Converter) {
            $base = Converter::class;
            throw new ConverterException("PDF Converter class '{$class}' must extend from '{$base}'");
        }
        $converter->convert();
        static::log("generatePdf('{$input}', '{$output}'): Finished");
    }
}
