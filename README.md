<h1 align="center">
    <a href="http://demos.krajee.com" title="Krajee Demos" target="_blank">
        <img src="http://kartik-v.github.io/bootstrap-fileinput-samples/samples/krajee-logo-b.png" alt="Krajee Logo"/>
    </a>
    <br>
    yii2-word-report
    <hr>
    <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=DTP3NZQ6G2AYU"
       title="Donate via Paypal" target="_blank">
        <img src="http://kartik-v.github.io/bootstrap-fileinput-samples/samples/donate.png" alt="Donate"/>
    </a>
</h1>

<div align="center">

[![Stable Version](https://poser.pugx.org/kartik-v/yii2-word-report/v/stable)](https://packagist.org/packages/kartik-v/yii2-word-report)
[![Unstable Version](https://poser.pugx.org/kartik-v/yii2-word-report/v/unstable)](https://packagist.org/packages/kartik-v/yii2-word-report)
[![License](https://poser.pugx.org/kartik-v/yii2-word-report/license)](https://packagist.org/packages/kartik-v/yii2-word-report)

[![Total Downloads](https://poser.pugx.org/kartik-v/yii2-word-report/downloads)](https://packagist.org/packages/kartik-v/yii2-word-report)
[![Monthly Downloads](https://poser.pugx.org/kartik-v/yii2-word-report/d/monthly)](https://packagist.org/packages/kartik-v/yii2-word-report)
[![Daily Downloads](https://poser.pugx.org/kartik-v/yii2-word-report/d/daily)](https://packagist.org/packages/kartik-v/yii2-word-report)

</div>

A Yii2 library to generate Word / PDF reports using Microsoft Word Templates. 

Refer [detailed documentation](http://demos.krajee.com/word-report) and/or a [complete demo](http://demos.krajee.com/word-report-demo). 

### Documentation and Demo
You can see detailed [documentation](http://demos.krajee.com/word-report) and [demonstration](http://demos.krajee.com/word-report-demo) on usage of the extension.

## Installation

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

### Pre-requisites
> Note: Check the [composer.json](https://github.com/kartik-v/yii2-dropdown-x/blob/master/composer.json) for this extension's requirements and dependencies. 
You must set the `minimum-stability` to `dev` in the **composer.json** file in your application root folder before installation of this extension OR
if your `minimum-stability` is set to any other value other than `dev`, then set the following in the require section of your composer.json file

```
kartik-v/yii2-word-report: "@dev"
```

Read this [web tip /wiki](http://webtips.krajee.com/setting-composer-minimum-stability-application/) on setting the `minimum-stability` settings for your application's composer.json.

### Install

Either run

```
$ php composer.phar require kartik-v/yii2-word-report "@dev"
```

or add

```
"kartik-v/yii2-word-report": "@dev"
```

to the ```require``` section of your `composer.json` file.

## Usage
```php
use kartik\wordreport\TemplateReport;

$report = new TemplateReport([
   'format' => TemplateReport::FORMAT_BOTH,
   'inputFile' => 'Invoice_Template_01.docx',
   'outputFile' => 'Invoice_Report_' . date('Y-m-d'),
   'values' => ['invoice_no' => 2001, 'invoice_date' => '2020-02-21'],
   'images' => ['company_logo' => '@webroot/images/company.jpg', 'customer_logo' => '@webroot/images/company.jpg'],
   'rows' => [
     'item' => [
         ['item' => 1, 'name' => 'Potato', 'price' => '$10.00'],
         ['item' => 2, 'name' => 'Tomato', 'price' => '$20.00'],
     ]
   ],
   'blocks' => [
     'customer_block' => [
         ['customer_name' => 'John', 'customer_address' => 'Address for John'],
         ['customer_name' => 'Bill', 'customer_address' => 'Address for Bill'],
     ],
   ]
]);
// Generate the report
$report->generate();
```

## License

**yii2-word-report** is released under the BSD-3-Clause License. See the bundled `LICENSE.md` for details.