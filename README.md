<p align="center">
    <a href="https://github.com/yii2tech" target="_blank">
        <img src="https://avatars2.githubusercontent.com/u/12951949" height="100px">
    </a>
    <h1 align="center">Spreadsheet Data Export extension for Yii2</h1>
    <br>
</p>

This extension provides ability to export data to spreadsheet, e.g. Excel, LibreOffice etc.

For license information check the [LICENSE](LICENSE.md)-file.

[![Latest Stable Version](https://poser.pugx.org/yii2tech/spreadsheet/v/stable.png)](https://packagist.org/packages/yii2tech/spreadsheet)
[![Total Downloads](https://poser.pugx.org/yii2tech/spreadsheet/downloads.png)](https://packagist.org/packages/yii2tech/spreadsheet)
[![Build Status](https://travis-ci.org/yii2tech/spreadsheet.svg?branch=master)](https://travis-ci.org/yii2tech/spreadsheet)


Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist yii2tech/spreadsheet
```

or add

```json
"yii2tech/spreadsheet": "*"
```

to the require section of your composer.json.


Usage
-----

This extension provides ability to export data to a spreadsheet, e.g. Excel, LibreOffice etc.
It is powered by [phpoffice/phpspreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) library.
Export is performed via [[\yii2tech\spreadsheet\Spreadsheet]] instance, which provides interface similar to [[\yii\grid\GridView]] widget.

Example:

```php
use yii2tech\spreadsheet\Spreadsheet;
use yii\data\ArrayDataProvider;

$exporter = new Spreadsheet([
    'dataProvider' => new ArrayDataProvider([
        'allModels' => [
            [
                'name' => 'some name',
                'price' => '9879',
            ],
            [
                'name' => 'name 2',
                'price' => '79',
            ],
        ],
    ]),
    'columns' => [
        [
            'attribute' => 'name',
            'contentOptions' => [
                'alignment' => [
                    'horizontal' => 'center',
                    'vertical' => 'center',
                ],
            ],
        ],
        [
            'attribute' => 'price',
        ],
    ],
]);
$exporter->save('/path/to/file.xls');
```

Please, refer to [[\yii2tech\spreadsheet\Column]] class for the information about column properties and configuration specifications.

While running web application you can use [[\yii2tech\spreadsheet\Spreadsheet::send()]] method to send a result file to
the browser through download dialog:

```php
use yii2tech\spreadsheet\Spreadsheet;
use yii\data\ActiveDataProvider;
use yii\web\Controller;

class ItemController extends Controller
{
    public function actionExport()
    {
        $exporter = new Spreadsheet([
            'dataProvider' => new ActiveDataProvider([
                'query' => Item::find(),
            ]),
        ]);
        return $exporter->send('items.xls');
    }
}
```


## Multiple sheet files <span id="multiple-sheet-files"></span>

You can create an output file with multiple worksheets (tabs). For example: you may want to export data about
equipment used in the office, keeping monitors, mouses, keyboards and so on in separated listings but in the same file.
To do so you will need to manually call [[\yii2tech\spreadsheet\Spreadsheet::render()]] method with different configuration
before creating final file. For example:

```php
use yii2tech\spreadsheet\Spreadsheet;
use yii\data\ActiveDataProvider;
use app\models\Equipment;

$exporter = (new Spreadsheet([
    'title' => 'Monitors',
    'dataProvider' => new ActiveDataProvider([
        'query' => Equipment::find()->andWhere(['group' => 'monitor']),
    ]),
    'columns' => [
        [
            'attribute' => 'name',
        ],
        [
            'attribute' => 'price',
        ],
    ],
]))->render(); // call `render()` to create a single worksheet

$exporter->configure([ // update spreadsheet configuration
    'title' => 'Mouses',
    'dataProvider' => new ActiveDataProvider([
        'query' => Equipment::find()->andWhere(['group' => 'mouse']),
    ]),
])->render(); // call `render()` to create a single worksheet

$exporter->configure([ // update spreadsheet configuration
    'title' => 'Keyboards',
    'dataProvider' => new ActiveDataProvider([
        'query' => Equipment::find()->andWhere(['group' => 'keyboard']),
    ]),
])->render(); // call `render()` to create a single worksheet

$exporter->save('/path/to/file.xls');
```

As the result you will get a single *.xls file with 3 worksheets (tabs): 'Monitors', 'Mouses' and 'Keyboards'.

Using [[\yii2tech\spreadsheet\Spreadsheet::configure()]] you can reset any spreadsheet parameter, including `columns`.
Thus you are able to combine several entirely different sheets into a single file.


## Large data processing <span id="large-data-processing"></span>

[[\yii2tech\spreadsheet\Spreadsheet]] allows exporting of the [[\yii\data\DataProviderInterface]] and [[\yii\db\QueryInterface]] instances.
Export is performed via batches, which allows processing of the large data without memory overflow.

In case of [[\yii\data\DataProviderInterface]] usage, data will be split to batches using pagination mechanism.
Thus you should setup pagination with page size in order to control batch size:

```php
use yii2tech\spreadsheet\Spreadsheet;
use yii\data\ActiveDataProvider;

$exporter = new Spreadsheet([
    'dataProvider' => new ActiveDataProvider([
        'query' => Item::find(),
        'pagination' => [
            'pageSize' => 100, // export batch size
        ],
    ]),
]);
$exporter->saveAs('/path/to/file.xls');
```

> Note: if you disable pagination in your data provider - no batch processing will be performed.

In case of [[\yii\db\QueryInterface]] usage, `Spreadsheet` will attempt to use `batch()` method, if it present in the query
class (for example in case [[\yii\db\Query]] or [[\yii\db\ActiveQuery]] usage). If `batch()` method is not available -
[[yii\data\ActiveDataProvider]] instance will be automatically created around given query.
You can control batch size via [[\yii2tech\spreadsheet\Spreadsheet::$batchSize]]:

```php
use yii2tech\spreadsheet\Spreadsheet;
use yii\data\ActiveDataProvider;

$exporter = new Spreadsheet([
    'query' => Item::find(),
    'batchSize' => 200, // export batch size
]);
$exporter->saveAs('/path/to/file.xls');
```

> Note: despite batch data processing reduces amount of resources needed for spreadsheet file generation,
  your program may still easily end up with PHP memory limit error on large data. This happens because of
  large complexity of the created document, which is stored in the memory during the entire process.
  In case you need to export really large data set, consider doing so via simple CSV data format
  using [yii2tech/csv-grid](https://github.com/yii2tech/csv-grid) extension.


## Complex headers <span id="complex-headers"></span>

You may union some columns in the sheet header into a groups. For example: you may have 2 different data columns:
'Planned Revenue' and 'Actual Revenue'. In this case you may want to display them as single column 'Revenue', split
into 2 sub columns: 'Planned' and 'Actual'.
This can be achieved using [[\yii2tech\spreadsheet\Spreadsheet::$headerColumnUnions]]. Its each entry
should specify 'offset', which determines the amount of columns to be skipped, and 'length', which determines
the amount of columns to be united. Other options of the union are the same as for regular column.
For example:

```php
use yii2tech\spreadsheet\Spreadsheet;
use yii\data\ArrayDataProvider;

$exporter = new Spreadsheet([
    'dataProvider' => new ArrayDataProvider([
        'allModels' => [
            [
                'column1' => '1.1',
                'column2' => '1.2',
                'column3' => '1.3',
                'column4' => '1.4',
                'column5' => '1.5',
                'column6' => '1.6',
                'column7' => '1.7',
            ],
            [
                'column1' => '2.1',
                'column2' => '2.2',
                'column3' => '2.3',
                'column4' => '2.4',
                'column5' => '2.5',
                'column6' => '2.6',
                'column7' => '2.7',
            ],
        ],
    ]),
    'headerColumnUnions' => [
        [
            'header' => 'Skip 1 column and group 2 next',
            'offset' => 1,
            'length' => 2,
        ],
        [
            'header' => 'Skip 2 columns and group 2 next',
            'offset' => 2,
            'length' => 2,
        ],
    ],
]);
$exporter->saveAs('/path/to/file.xls');
```

> Note: only single level of header column unions is supported. You will need to deal with more complex
  cases on your own.


## Custom cell rendering <span id="custom-cell-rendering"></span>

Before `save()` or `send()` method is invoked, you are able to edit generated spreadsheet, making some
final adjustments to it. Several methods exist to facilitate this process:

 - [[\yii2tech\spreadsheet\Spreadsheet::renderCell()]] - renders specified cell with given content and style.
 - [[\yii2tech\spreadsheet\Spreadsheet::applyCellStyle()]] - applies specified style to the cell.
 - [[\yii2tech\spreadsheet\Spreadsheet::mergeCells()]] - merges sell range into single one.

You may use these methods, after document has been composed via [[\yii2tech\spreadsheet\Spreadsheet::render()]],
to override or add some content. For example:

```php
use yii2tech\spreadsheet\Spreadsheet;
use yii2tech\spreadsheet\SerialColumn;
use yii\data\ArrayDataProvider;

$exporter = new Spreadsheet([
    'dataProvider' => new ArrayDataProvider([
        'allModels' => [
            [
                'id' => 1,
                'name' => 'first',
            ],
            [
                'id' => 2,
                'name' => 'second',
            ],
        ],
    ]),
    'columns' => [
        [
            'class' => SerialColumn::class,
        ],
        [
            'attribute' => 'id',
        ],
        [
            'attribute' => 'name',
        ],
    ],
])->render(); // render the document

// override serial column header :
$grid->renderCell('A1', 'Overridden serial column header');

// add custom footer :
$grid->renderCell('A4', 'Custom A4', [
    'font' => [
        'color' => [
            'rgb' => '#FF0000',
        ],
    ],
]);

// merge footer cells :
$grid->mergeCells('A4:B4');

$exporter->saveAs('/path/to/file.xls');
```

> Tip: you can use [[\yii2tech\spreadsheet\Spreadsheet::$rowIndex]] to get number of the row, which is next
  to the last rendered one.
