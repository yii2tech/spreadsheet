<?php
/**
 * @link https://github.com/yii2tech
 * @copyright Copyright (c) 2015 Yii2tech
 * @license [New BSD License](http://www.opensource.org/licenses/bsd-license.php)
 */

namespace yii2tech\spreadsheet;

use PhpOffice\PhpSpreadsheet\IOFactory;
use yii\data\ActiveDataProvider;
use yii\helpers\FileHelper;
use yii\i18n\Formatter;
use Yii;
use yii\base\Component;
use yii\base\InvalidConfigException;
use yii\di\Instance;
use yii\web\Response;

/**
 * Spreadsheet allows export of data provider into Excel document via {@see \PhpOffice\PhpSpreadsheet\Spreadsheet} library.
 * It provides interface, which is similar to {@see \yii\grid\GridView} widget.
 *
 * Example:
 *
 * ```php
 * use yii2tech\spreadsheet\Spreadsheet;
 * use yii\data\ActiveDataProvider;
 *
 * $exporter = new Spreadsheet([
 *     'dataProvider' => new ArrayDataProvider([
 *         'allModels' => [
 *             [
 *                 'name' => 'some name',
 *                 'price' => '9879',
 *             ],
 *             [
 *                 'name' => 'name 2',
 *                 'price' => '79',
 *             ],
 *         ],
 *     ]),
 *     'columns' => [
 *         [
 *             'attribute' => 'name',
 *             'contentOptions' => [
 *                 'alignment' => [
 *                     'horizontal' => 'center',
 *                     'vertical' => 'center',
 *                 ],
 *             ],
 *         ],
 *         [
 *             'attribute' => 'price',
 *         ],
 *     ],
 * ]);
 * $exporter->save('/path/to/file.xls');
 * ```
 *
 * @see https://phpspreadsheet.readthedocs.io/
 * @see \PhpOffice\PhpSpreadsheet\Spreadsheet
 *
 * @property array|Formatter $formatter the formatter used to format model attribute values into displayable texts.
 * @property \PhpOffice\PhpSpreadsheet\Spreadsheet $document spreadsheet document representation instance.
 *
 * @author Paul Klimov <klimov.paul@gmail.com>
 * @since 1.0
 */
class Spreadsheet extends Component
{
    /**
     * @var \yii\data\DataProviderInterface the data provider for the view. This property is required.
     */
    public $dataProvider;
    /**
     * @var \yii\db\QueryInterface the data source query.
     * Note: this field will be ignored in case {@see dataProvider} is set.
     */
    public $query;
    /**
     * @var int the number of records to be fetched in each batch.
     * This property takes effect only in case of {@see query} usage.
     */
    public $batchSize = 100;
    /**
     * @var array|Column[] spreadsheet column configuration. Each array element represents the configuration
     * for one particular column. For example:
     *
     * ```php
     * [
     *     ['class' => SerialColumn::class],
     *     [
     *         'class' => DataColumn::class, // this line is optional
     *         'attribute' => 'name',
     *         'format' => 'text',
     *         'header' => 'Name',
     *     ],
     * ]
     * ```
     *
     * If a column is of class {@see DataColumn}, the "class" element can be omitted.
     */
    public $columns = [];
    /**
     * @var bool whether to show the header section of the sheet.
     */
    public $showHeader = true;
    /**
     * @var bool whether to show the footer section of the sheet.
     */
    public $showFooter = false;
    /**
     * @var string|null sheet title.
     */
    public $title;
    /**
     * @var string the HTML display when the content of a cell is empty.
     * This property is used to render cells that have no defined content,
     * e.g. empty footer or filter cells.
     *
     * Note that this is not used by the {@see DataColumn} if a data item is `null`. In that case
     * the {@see nullDisplay} property will be used to indicate an empty data value.
     */
    public $emptyCell = '';
    /**
     * @var string the text to be displayed when formatting a `null` data value.
     */
    public $nullDisplay = '';
    /**
     * @var string writer type (format type). If not set, it will be determined automatically.
     * Supported values:
     *
     * - 'Xls'
     * - 'Xlsx'
     * - 'Ods'
     * - 'Csv'
     * - 'Html'
     * - 'Tcpdf'
     * - 'Dompdf'
     * - 'Mpdf'
     *
     * @see IOFactory
     */
    public $writerType;
    /**
     * @var callable|null a PHP callback, which should create spreadsheet writer instance.
     * The signature of this callback should be following: `function(\PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet, string $writerType): \PhpOffice\PhpSpreadsheet\Writer\IWriter`
     * @see \PhpOffice\PhpSpreadsheet\Writer\IWriter
     * @since 1.0.5
     */
    public $writerCreator;
    /**
     * @var array[] list of header column unions.
     * For example:
     *
     * ```php
     * [
     *     [
     *         'header' => 'Skip one column and group 3 next',
     *         'offset' => 1,
     *         'length' => 3,
     *     ],
     *     [
     *         'header' => 'Skip two column and group 5 next',
     *         'offset' => 2,
     *         'length' => 5,
     *     ],
     * ]
     * ```
     */
    public $headerColumnUnions = [];
    /**
     * @var int|null current sheet row index.
     * Value of this field automatically changes during spreadsheet rendering. After rendering is complete,
     * it will contain the number of the row next to the latest fill-up one.
     * Note: be careful while manually manipulating value of this field as it may cause unexpected results.
     */
    public $rowIndex;
    /**
     * @var int index of the sheet row, from which rendering should start.
     * This field can be used to skip some lines at the sheet beginning for the further manual fill up.
     * @since 1.0.4
     */
    public $startRowIndex = 1;

    /**
     * @var bool whether spreadsheet has been already rendered or not.
     */
    protected $isRendered = false;

    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet|null spreadsheet document representation instance.
     */
    private $_document;
    /**
     * @var array|Formatter the formatter used to format model attribute values into displayable texts.
     * This can be either an instance of {@see Formatter} or an configuration array for creating the {@see Formatter}
     * instance. If this property is not set, the "formatter" application component will be used.
     */
    private $_formatter;
    /**
     * @var array|null internal iteration information.
     */
    private $batchInfo;


    /**
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet spreadsheet document representation instance.
     */
    public function getDocument()
    {
        if (!is_object($this->_document)) {
            $this->_document = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        }
        return $this->_document;
    }

    /**
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet|null $document spreadsheet document representation instance.
     */
    public function setDocument($document)
    {
        $this->_document = $document;
    }

    /**
     * @return Formatter formatter instance.
     */
    public function getFormatter()
    {
        if (!is_object($this->_formatter)) {
            if ($this->_formatter === null) {
                $this->_formatter = Yii::$app->getFormatter();
            } else {
                $this->_formatter = Instance::ensure($this->_formatter, Formatter::class);
            }
        }
        return $this->_formatter;
    }

    /**
     * @param array|Formatter $formatter formatter instance.
     */
    public function setFormatter($formatter)
    {
        $this->_formatter = $formatter;
    }

    /**
     * Creates column objects and initializes them.
     */
    protected function initColumns()
    {
        foreach ($this->columns as $i => $column) {
            if (is_string($column)) {
                $column = $this->createDataColumn($column);
            } elseif (is_array($column)) {
                $column = Yii::createObject(array_merge([
                    'class' => DataColumn::class,
                    'grid' => $this,
                ], $column));
            }
            if (!$column->visible) {
                unset($this->columns[$i]);
                continue;
            }
            $this->columns[$i] = $column;
        }
    }

    /**
     * This function tries to guess the columns to show from the given data
     * if {@see columns} are not explicitly specified.
     * @param \yii\base\Model|array $model model to be used for column information source.
     */
    protected function guessColumns($model)
    {
        if (is_array($model) || is_object($model)) {
            foreach ($model as $name => $value) {
                $this->columns[] = (string) $name;
            }
        }
    }

    /**
     * Creates a {@see DataColumn} object based on a string in the format of "attribute:format:label".
     * @param string $text the column specification string
     * @return DataColumn the column instance
     * @throws InvalidConfigException if the column specification is invalid
     */
    protected function createDataColumn($text)
    {
        if (!preg_match('/^([^:]+)(:(\w*))?(:(.*))?$/', $text, $matches)) {
            throw new InvalidConfigException('The column must be specified in the format of "attribute", "attribute:format" or "attribute:format:label"');
        }

        return Yii::createObject([
            'class' => DataColumn::class,
            'grid' => $this,
            'attribute' => $matches[1],
            'format' => isset($matches[3]) ? $matches[3] : 'raw',
            'label' => isset($matches[5]) ? $matches[5] : null,
        ]);
    }

    /**
     * Sets spreadsheet document properties.
     * @param array $properties list of document properties in format: name => value
     * @return $this self reference.
     * @see \PhpOffice\PhpSpreadsheet\Document\Properties
     */
    public function properties($properties)
    {
        $documentProperties = $this->getDocument()->getProperties();
        foreach ($properties as $name => $value) {
            $method = 'set' . ucfirst($name);
            call_user_func([$documentProperties, $method], $value);
        }
        return $this;
    }

    /**
     * Configures (re-configures) this spreadsheet with the property values.
     * This method is useful for rendering multisheet documents. For example:
     *
     * ```php
     * (new Spreadsheet([
     *     'title' => 'Monitors',
     *     'dataProvider' => $monitorDataProvider,
     * ]))
     * ->render()
     * ->configure([
     *     'title' => 'Mouses',
     *     'dataProvider' => $mouseDataProvider,
     * ])
     * ->render()
     * ->configure([
     *     'title' => 'Keyboards',
     *     'dataProvider' => $keyboardDataProvider,
     * ])
     * ->save('/path/to/export/files/office-equipment.xls');
     * ```
     *
     * @param array $properties the property initial values given in terms of name-value pairs.
     * @return $this self reference.
     */
    public function configure($properties)
    {
        Yii::configure($this, $properties);
        return $this;
    }

    /**
     * Performs actual document composition.
     * @return $this self reference.
     */
    public function render()
    {
        if ($this->dataProvider === null) {
            if ($this->query !== null) {
                $this->dataProvider = new ActiveDataProvider([
                    'query' => $this->query,
                    'pagination' => [
                        'pageSize' => $this->batchSize,
                    ],
                ]);
            }
        }

        $document = $this->getDocument();

        if ($this->isRendered) {
            // second run
            $document->createSheet();
            $document->setActiveSheetIndex($document->getActiveSheetIndex() + 1);
        }

        if ($this->title !== null) {
            $document->getActiveSheet()->setTitle($this->title);
        }

        $this->rowIndex = $this->startRowIndex;

        $columnsInitialized = false;
        $modelIndex = 0;
        while (($data = $this->batchModels()) !== false) {
            list($models, $keys) = $data;

            if (!$columnsInitialized) {
                if (empty($this->columns)) {
                    $this->guessColumns(reset($models));
                }

                $this->initColumns();
                $this->applyColumnOptions();
                $columnsInitialized = true;

                if ($this->showHeader) {
                    $this->renderHeader();
                }
            }

            $this->renderBody($models, $keys, $modelIndex);
            $this->gc();
        }

        if ($this->showFooter) {
            $this->renderFooter();
        }

        $this->isRendered = true;

        return $this;
    }

    /**
     * Renders sheet table body batch.
     * This method will be invoked several times, one per each model batch.
     * @param array $models batch of models.
     * @param array $keys batch of model keys.
     * @param int $modelIndex model iteration index.
     */
    protected function renderBody($models, $keys, &$modelIndex)
    {
        foreach ($models as $index => $model) {
            $key = isset($keys[$index]) ? $keys[$index] : $index;
            $columnIndex = 'A';
            foreach ($this->columns as $column) {
                /* @var $column Column */
                $column->renderDataCell($columnIndex . $this->rowIndex, $model, $key, $modelIndex);
                $columnIndex++;
            }
            $this->rowIndex++;
            $modelIndex++;
        }
    }

    /**
     * Renders sheet table header
     */
    protected function renderHeader()
    {
        if (empty($this->headerColumnUnions)) {
            $columnIndex = 'A';
            foreach ($this->columns as $column) {
                /* @var $column Column */
                $column->renderHeaderCell($columnIndex . $this->rowIndex);
                $columnIndex++;
            }
            $this->rowIndex++;
            return;
        }

        $sheet = $this->getDocument()->getActiveSheet();

        $columns = $this->columns;

        $columnIndex = 'A';
        foreach ($this->headerColumnUnions as $columnUnion) {
            if (isset($columnUnion['offset'])) {
                $offset = (int)$columnUnion['offset'];
                unset($columnUnion['offset']);
            } else {
                $offset = 0;
            }

            if (isset($columnUnion['length'])) {
                $length = (int)$columnUnion['length'];
                unset($columnUnion['length']);
            } else {
                $length = 1;
            }

            while ($offset > 0) {
                /* @var $column Column */
                $column = array_shift($columns);
                $column->renderHeaderCell($columnIndex . $this->rowIndex);

                $sheet->mergeCells($columnIndex . ($this->rowIndex) . ':' . $columnIndex . ($this->rowIndex + 1));
                $columnIndex++;
                $offset--;
            }

            $column = new Column($columnUnion);
            $column->grid = $this;
            $column->renderHeaderCell($columnIndex . $this->rowIndex);

            $startColumnIndex = $columnIndex;
            while (true) {
                /* @var $column Column */
                $column = array_shift($columns);
                $column->renderHeaderCell($columnIndex . ($this->rowIndex + 1));
                $length--;
                if (($length < 1)) {
                    break;
                }
                $columnIndex++;
            }

            $sheet->mergeCells($startColumnIndex . $this->rowIndex . ':' . $columnIndex . $this->rowIndex);

            $columnIndex++;
        }

        foreach ($columns as $column) {
            /* @var $column Column */
            $column->renderHeaderCell($columnIndex . $this->rowIndex);
            $sheet->mergeCells($columnIndex . ($this->rowIndex) . ':' . $columnIndex . ($this->rowIndex + 1));
            $columnIndex++;
        }

        $this->rowIndex++;
        $this->rowIndex++;
    }

    /**
     * Renders sheet table footer
     */
    protected function renderFooter()
    {
        $columnIndex = 'A';
        foreach ($this->columns as $column) {
            /* @var $column Column */
            $column->renderFooterCell($columnIndex . $this->rowIndex);
            $columnIndex++;
        }
        $this->rowIndex++;
    }

    /**
     * Applies column overall options, such as dimension options.
     */
    protected function applyColumnOptions()
    {
        $sheet = $this->getDocument()->getActiveSheet();
        $columnIndex = 'A';
        foreach ($this->columns as $column) {
            /* @var $column Column */
            if (!empty($column->dimensionOptions)) {
                $columnDimension = $sheet->getColumnDimension($columnIndex);
                foreach ($column->dimensionOptions as $name => $value) {
                    $method = 'set' . ucfirst($name);
                    call_user_func([$columnDimension, $method], $value);
                }
            }

            $columnIndex++;
        }
    }

    /**
     * Iterates over {@see query} or {@see dataProvider} returning data by batches.
     * @return array|false data batch: first element - models list, second model keys list.
     */
    protected function batchModels()
    {
        if ($this->batchInfo === null) {
            if ($this->query !== null && method_exists($this->query, 'batch')) {
                $this->batchInfo = [
                    'queryIterator' => $this->query->batch($this->batchSize)
                ];
            } else {
                $this->batchInfo = [
                    'pagination' => $this->dataProvider->getPagination(),
                    'page' => 0
                ];
            }
        }

        if (isset($this->batchInfo['queryIterator'])) {
            /* @var $iterator \Iterator */
            $iterator = $this->batchInfo['queryIterator'];
            $iterator->next();

            if ($iterator->valid()) {
                return [$iterator->current(), []];
            }

            $this->batchInfo = null;
            return false;
        }

        if (isset($this->batchInfo['pagination'])) {
            /* @var $pagination \yii\data\Pagination|bool */
            $pagination = $this->batchInfo['pagination'];
            $page = $this->batchInfo['page'];

            if ($pagination === false || $pagination->pageCount === 0) {
                if ($page === 0) {
                    $this->batchInfo['page']++;
                    return [
                        $this->dataProvider->getModels(),
                        $this->dataProvider->getKeys()
                    ];
                }
            } else {
                if ($page < $pagination->pageCount) {
                    $pagination->setPage($page);
                    $this->dataProvider->prepare(true);
                    $this->batchInfo['page']++;
                    return [
                        $this->dataProvider->getModels(),
                        $this->dataProvider->getKeys()
                    ];
                }
            }

            $this->batchInfo = null;
            return false;
        }

        return false;
    }

    /**
     * Renders cell with given coordinates.
     * @param string $cell cell coordinates, e.g. 'A1', 'B4' etc.
     * @param string $content cell raw content.
     * @param array $style cell style options.
     * @return $this self reference.
     */
    public function renderCell($cell, $content, $style = [])
    {
        $sheet = $this->getDocument()->getActiveSheet();
        $sheet->setCellValue($cell, $content);
        $this->applyCellStyle($cell, $style);
        return $this;
    }

    /**
     * Applies cell style from configuration.
     * @param string $cell cell coordinates, e.g. 'A1', 'B4' etc.
     * @param array $style style configuration.
     * @return $this self reference.
     * @throws \PhpOffice\PhpSpreadsheet\Exception on failure.
     */
    public function applyCellStyle($cell, $style)
    {
        if (empty($style)) {
            return $this;
        }

        $cellStyle = $this->getDocument()->getActiveSheet()->getStyle($cell);
        if (isset($style['alignment'])) {
            $cellStyle->getAlignment()->applyFromArray($style['alignment']);
            unset($style['alignment']);
            if (empty($style)) {
                return $this;
            }
        }
        $cellStyle->applyFromArray($style);

        return $this;
    }

    /**
     * Merges sell range into single one.
     * @param string $cellRange cell range (e.g. 'A1:E1').
     * @return $this self reference.
     * @throws \PhpOffice\PhpSpreadsheet\Exception on failure.
     */
    public function mergeCells($cellRange)
    {
        $this->getDocument()->getActiveSheet()->mergeCells($cellRange);
        return $this;
    }

    /**
     * Saves the document into a file.
     * @param string $filename name of the output file.
     */
    public function save($filename)
    {
        if (!$this->isRendered) {
            $this->render();
        }

        $filename = Yii::getAlias($filename);

        $writerType = $this->writerType;
        if ($writerType === null) {
            $fileExtension = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
            $writerType = ucfirst($fileExtension);
        }

        $fileDir = pathinfo($filename, PATHINFO_DIRNAME);
        FileHelper::createDirectory($fileDir);

        $writer = $this->createWriter($writerType);
        $writer->save($filename);
    }

    /**
     * Sends the rendered content as a file to the browser.
     *
     * Note that this method only prepares the response for file sending. The file is not sent
     * until {@see \yii\web\Response::send()} is called explicitly or implicitly.
     * The latter is done after you return from a controller action.
     *
     * @param string $attachmentName the file name shown to the user.
     * @param array $options additional options for sending the file. The following options are supported:
     *
     *  - `mimeType`: the MIME type of the content. Defaults to 'application/octet-stream'.
     *  - `inline`: bool, whether the browser should open the file within the browser window. Defaults to false,
     *    meaning a download dialog will pop up.
     *
     * @return \yii\web\Response the response object.
     */
    public function send($attachmentName, $options = [])
    {
        if (!$this->isRendered) {
            $this->render();
        }

        $writerType = $this->writerType;
        if ($writerType === null) {
            $fileExtension = strtolower(pathinfo($attachmentName, PATHINFO_EXTENSION));
            $writerType = ucfirst($fileExtension);
        }

        $tmpResource = tmpfile();
        if ($tmpResource === false) {
            throw new \RuntimeException('Unable to create temporary file.');
        }

        $tmpResourceMetaData = stream_get_meta_data($tmpResource);
        $tmpFileName = $tmpResourceMetaData['uri'];

        $writer = $this->createWriter($writerType);
        $writer->save($tmpFileName);
        unset($writer);

        $tmpFileStatistics = fstat($tmpResource);
        if ($tmpFileStatistics['size'] > 0) {
            return Yii::$app->getResponse()->sendStreamAsFile($tmpResource, $attachmentName, $options);
        }

        // some writers, like 'Xlsx', may delete target file during the process, making temporary file resource invalid
        $response = Yii::$app->getResponse();
        $response->on(Response::EVENT_AFTER_SEND, function() use ($tmpResource) {
            // with temporary file resource closing file matching its URI will be deleted, even if resource is invalid
            fclose($tmpResource);
        });

        return $response->sendFile($tmpFileName, $attachmentName, $options);
    }

    /**
     * Performs PHP memory garbage collection.
     */
    protected function gc()
    {
        if (!gc_enabled()) {
            gc_enable();
        }
        gc_collect_cycles();
    }

    /**
     * Creates a spreadsheet writer for the given type.
     * @param string $writerType spreadsheet writer type.
     * @return \PhpOffice\PhpSpreadsheet\Writer\IWriter
     * @since 1.0.5
     */
    protected function createWriter($writerType)
    {
        if ($this->writerCreator === null) {
            return IOFactory::createWriter($this->getDocument(), $writerType);
        }

        return call_user_func($this->writerCreator, $this->getDocument(), $writerType);
    }
}