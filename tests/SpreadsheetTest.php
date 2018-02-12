<?php

namespace yii2tech\tests\unit\spreadsheet;

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
}