<?php
/**
 * @link https://github.com/yii2tech
 * @copyright Copyright (c) 2015 Yii2tech
 * @license [New BSD License](http://www.opensource.org/licenses/bsd-license.php)
 */

namespace yii2tech\spreadsheet;

use PhpOffice\PhpSpreadsheet\IOFactory;
use yii\helpers\FileHelper;
use yii\i18n\Formatter;
use Yii;
use yii\base\Component;
use yii\base\InvalidConfigException;
use yii\di\Instance;

/**
 * Spreadsheet allows export of data provider into Excel document via [[\PhpOffice\PhpSpreadsheet\Spreadsheet]] library.
 * It provides interface, which is similar to [[\yii\grid\GridView]] widget.
 *
 * Example:
 *
 * ```php
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
 * $exporter->render()->save('/path/to/file.xls');
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
     * @var array|Column[]
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
     * Note that this is not used by the [[DataColumn]] if a data item is `null`. In that case
     * the [[nullDisplay]] property will be used to indicate an empty data value.
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
     * @var int current sheet row index.
     */
    protected $rowIndex;
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
     * This can be either an instance of [[Formatter]] or an configuration array for creating the [[Formatter]]
     * instance. If this property is not set, the "formatter" application component will be used.
     */
    private $_formatter;


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
     * @return Formatter formatter instance
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
     * @param array|Formatter $formatter
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
        if (empty($this->columns)) {
            $this->guessColumns();
        }

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
     * if [[columns]] are not explicitly specified.
     */
    protected function guessColumns()
    {
        $models = $this->dataProvider->getModels();
        $model = reset($models);
        if (is_array($model) || is_object($model)) {
            foreach ($model as $name => $value) {
                $this->columns[] = (string) $name;
            }
        }
    }

    /**
     * Creates a [[DataColumn]] object based on a string in the format of "attribute:format:label".
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
            'format' => isset($matches[3]) ? $matches[3] : 'text',
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
     * @param \yii\data\DataProviderInterface $dataProvider the data provider for the document.
     * @return $this self reference.
     */
    public function dataProvider($dataProvider)
    {
        $this->dataProvider = $dataProvider;
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
        $this->initColumns();

        $document = $this->getDocument();

        if ($this->rowIndex !== null) {
            // second run
            $document->createSheet();
            $document->setActiveSheetIndex($document->getActiveSheetIndex() + 1);
        }

        if ($this->title !== null) {
            $document->getActiveSheet()->setTitle($this->title);
        }

        $this->rowIndex = 1;

        $this->applyColumnOptions();

        if ($this->showHeader) {
            $this->renderHeader();
        }

        $this->renderBody();

        if ($this->showFooter) {
            $this->renderFooter();
        }

        $this->isRendered = true;

        return $this;
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
     * Renders sheet table body
     */
    protected function renderBody()
    {
        $models = array_values($this->dataProvider->getModels());
        $keys = $this->dataProvider->getKeys();
        foreach ($models as $index => $model) {
            $key = $keys[$index];
            $columnIndex = 'A';
            foreach ($this->columns as $column) {
                /* @var $column Column */
                $column->renderDataCell($columnIndex . $this->rowIndex, $model, $key, $index);
                $columnIndex++;
            }
            $this->rowIndex++;
        }
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

        $fileDir = strtolower(pathinfo($filename, PATHINFO_DIRNAME));
        FileHelper::createDirectory($fileDir);

        $objWriter = IOFactory::createWriter($this->getDocument(), $writerType);
        $objWriter->save($filename);
    }

    /**
     * Sends the rendered content as a file to the browser.
     *
     * Note that this method only prepares the response for file sending. The file is not sent
     * until [[\yii\web\Response::send()]] is called explicitly or implicitly.
     * The latter is done after you return from a controller action.
     *
     * @param string $attachmentName the file name shown to the user.
     * @param array $options additional options for sending the file. The following options are supported:
     *
     *  - `mimeType`: the MIME type of the content. Defaults to 'application/octet-stream'.
     *  - `inline`: bool, whether the browser should open the file within the browser window. Defaults to false,
     *    meaning a download dialog will pop up.
     *
     * @return \yii\web\Response the response object
     */
    public function send($attachmentName, $options = [])
    {
        $tempFileName = tempnam(Yii::getAlias('@runtime'), 'SpreadsheetTemp_');
        $fileExtension = strtolower(pathinfo($attachmentName, PATHINFO_EXTENSION));
        if (!empty($fileExtension)) {
            $tempFileName .= '.' . $fileExtension;
        }
        $this->save($tempFileName);
        $content = file_get_contents($tempFileName);
        $response = Yii::$app->getResponse()->sendContentAsFile($content, $attachmentName, $options);
        unlink($tempFileName);
        return $response;
    }
}