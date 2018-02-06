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

        $fileName = $this->getTestFilePath() . '/save.xls';
        $grid->render()->save($fileName);

        $this->assertTrue(file_exists($fileName));
    }
}