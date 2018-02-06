<?php

namespace yii2tech\tests\unit\spreadsheet;

use yii2tech\spreadsheet\ExcelGrid;
use Yii;
use yii\data\ArrayDataProvider;
use yii\i18n\Formatter;

class ExcelGridTest extends TestCase
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
     * @return ExcelGrid Excel grid instance.
     */
    protected function createExcelGrid(array $config = [])
    {
        if (!isset($config['dataProvider']) && !isset($config['query'])) {
            $config['dataProvider'] = new ArrayDataProvider();
        }
        return new ExcelGrid($config);
    }

    // Tests :

    public function testSetupFormatter()
    {
        $grid = $this->createExcelGrid();

        $formatter = new Formatter();
        $grid->setFormatter($formatter);
        $this->assertSame($formatter, $grid->getFormatter());
    }

    public function testExport()
    {
        $grid = $this->createExcelGrid([
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

        $fileName = $this->getTestFilePath() . '/save.xls';
        $grid->render()->save($fileName);

        $this->assertTrue(file_exists($fileName));
    }
}