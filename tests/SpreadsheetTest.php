<?php

namespace yii2tech\tests\unit\spreadsheet;

use yii\data\ActiveDataProvider;
use yii\db\Query;
use yii2tech\spreadsheet\DataColumn;
use yii2tech\spreadsheet\SerialColumn;
use yii2tech\spreadsheet\Spreadsheet;
use Yii;
use yii\data\ArrayDataProvider;
use yii\i18n\Formatter;

class SpreadsheetTest extends TestCase
{
    /**
     * Setup tables for test ActiveRecord
     */
    protected function setupTestDbData()
    {
        $db = Yii::$app->getDb();

        // Structure :

        $table = 'Item';
        $columns = [
            'id' => 'pk',
            'name' => 'string',
            'number' => 'integer',
        ];
        $db->createCommand()->createTable($table, $columns)->execute();

        $db->createCommand()->batchInsert($table, ['name', 'number'], [
            ['first', 1],
            ['second', 2],
            ['third', 3],
        ])->execute();
    }

    /**
     * @param array $config Excel grid configuration.
     * @return Spreadsheet Excel grid instance.
     */
    protected function createSpreadsheet(array $config = [])
    {
        if (!isset($config['dataProvider']) && !isset($config['query'])) {
            $config['dataProvider'] = new ArrayDataProvider();
        }
        return new Spreadsheet($config);
    }

    // Tests :

    public function testSetupFormatter()
    {
        $grid = $this->createSpreadsheet();

        $formatter = new Formatter();
        $grid->setFormatter($formatter);
        $this->assertSame($formatter, $grid->getFormatter());
    }

    public function testInitColumns()
    {
        $grid = $this->createSpreadsheet([
            'dataProvider' => new ArrayDataProvider([
                'allModels' => [
                    [
                        'id' => 1,
                        'name' => 'first',
                        'description' => 'first description',
                    ],
                    [
                        'id' => 2,
                        'name' => 'second',
                        'description' => 'second description',
                    ],
                ],
            ]),
            'columns' => [
                ['class' => SerialColumn::class],
                'id',
                'name:text',
                [
                    'attribute' => 'description'
                ],
            ],
        ]);
        $grid->render();

        $this->assertCount(4, $grid->columns);
        list($serialColumn, $idColumn, $nameColumn, $descriptionColumn) = $grid->columns;

        $this->assertTrue($serialColumn instanceof SerialColumn);
        /* @var $idColumn DataColumn */
        /* @var $nameColumn DataColumn */
        /* @var $descriptionColumn DataColumn */
        $this->assertTrue($idColumn instanceof DataColumn);
        $this->assertSame('id', $idColumn->attribute);
        $this->assertSame('raw', $idColumn->format);
        $this->assertTrue($nameColumn instanceof DataColumn);
        $this->assertSame('name', $nameColumn->attribute);
        $this->assertSame('text', $nameColumn->format);
        $this->assertTrue($descriptionColumn instanceof DataColumn);
        $this->assertSame('description', $descriptionColumn->attribute);
        $this->assertSame('raw', $descriptionColumn->format);
    }

    public function testExport()
    {
        $grid = $this->createSpreadsheet([
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
            ])
        ]);

        $fileName = $this->getTestFilePath() . '/basic.xls';
        $grid->save($fileName);

        $this->assertTrue(file_exists($fileName));
        $this->assertSame(4, $grid->rowIndex);
    }

    /**
     * @depends testExport
     */
    public function testExportColumnOptions()
    {
        $grid = $this->createSpreadsheet([
            'dataProvider' => new ArrayDataProvider([
                'allModels' => [
                    [
                        'id' => 1,
                        'name' => 'first',
                        'number' => 10,
                    ],
                    [
                        'id' => 2,
                        'name' => 'second',
                        'number' => 20,
                    ],
                ],
            ]),
            'columns' => [
                [
                    'attribute' => 'id',
                    'dimensionOptions' => [
                        'autoSize' => true
                    ],
                ],
                [
                    'attribute' => 'name',
                    'headerOptions' => [
                        'borders' => [
                            'allBorders' => [
                                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOTTED,
                            ],
                        ],
                    ],
                    'contentOptions' => [
                        'borders' => [
                            'allBorders' => [
                                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOTTED,
                            ],
                        ],
                        'alignment' => [
                            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                        ],
                    ],
                ],
                [
                    'attribute' => 'number',
                    'contentOptions' => [
                        'font' => [
                            'color' => [
                                'rgb' => '#FF0000',
                            ],
                        ],
                    ],
                ],
            ],
        ]);

        $fileName = $this->getTestFilePath() . '/column-options.xls';
        $grid->save($fileName);

        $this->assertTrue(file_exists($fileName));
    }

    /**
     * @depends testExport
     */
    public function testExportMultipleSheets()
    {
        $grid = $this->createSpreadsheet([
            'title' => 'items page 1',
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
                    'attribute' => 'id',
                    'header' => 'ID',
                ],
                [
                    'attribute' => 'name',
                    'header' => 'Name',
                ],
            ],
        ])
        ->render()
        ->configure([
            'title' => 'items page 2',
            'dataProvider' => new ArrayDataProvider([
                'allModels' => [
                    [
                        'id' => 3,
                        'name' => 'third',
                    ],
                    [
                        'id' => 4,
                        'name' => 'fourth',
                    ],
                ],
            ])
        ])
        ->render();

        $fileName = $this->getTestFilePath() . '/multiple-sheet.xls';
        $grid->save($fileName);

        $this->assertTrue(file_exists($fileName));
    }

    /**
     * @depends testExport
     */
    public function testExportColumnUnions()
    {
        $grid = $this->createSpreadsheet([
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

        $fileName = $this->getTestFilePath() . '/column-unions.xls';
        $grid->save($fileName);

        $this->assertTrue(file_exists($fileName));
        $this->assertSame(5, $grid->rowIndex);
    }

    /**
     * @depends testExport
     */
    public function testCustomCellRender()
    {
        $grid = $this->createSpreadsheet([
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
            ])
        ])->render();

        $grid->renderCell('A4', 'Custom A4', [
            'font' => [
                'color' => [
                    'rgb' => '#FF0000',
                ],
            ],
        ]);
        $grid->mergeCells('A4:B4');

        $fileName = $this->getTestFilePath() . '/custom-render.xls';
        $grid->save($fileName);

        $this->assertTrue(file_exists($fileName));
    }

    /**
     * @depends testExport
     */
    public function testExportSerialColumn()
    {
        $grid = $this->createSpreadsheet([
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
        ]);

        $fileName = $this->getTestFilePath() . '/serial-column.xls';
        $grid->save($fileName);

        $this->assertTrue(file_exists($fileName));
        $this->assertSame(4, $grid->rowIndex);
    }

    /**
     * @depends testExport
     */
    public function testExportQuery()
    {
        $this->setupTestDbData();

        $query = (new Query())->from('Item');

        $grid = $this->createSpreadsheet([
            'query' => $query,
            'batchSize' => 2
        ]);

        $fileName = $this->getTestFilePath() . '/query.xls';
        $grid->save($fileName);

        $this->assertTrue(file_exists($fileName));
        $this->assertSame(5, $grid->rowIndex);
    }

    /**
     * @depends testExport
     */
    public function testExportDataProviderIterate()
    {
        $this->setupTestDbData();

        $query = (new Query())->from('Item');

        $grid = $this->createSpreadsheet([
            'dataProvider' => new ActiveDataProvider([
                'query' => $query,
                'pagination' => [
                    'pageSize' => 2
                ],
            ]),
        ]);

        $fileName = $this->getTestFilePath() . '/data-provider-iterator.xls';
        $grid->save($fileName);

        $this->assertTrue(file_exists($fileName));
        $this->assertSame(5, $grid->rowIndex);
    }
}