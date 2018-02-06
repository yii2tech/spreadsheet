<?php
/**
 * @link https://github.com/yii2tech
 * @copyright Copyright (c) 2015 Yii2tech
 * @license [New BSD License](http://www.opensource.org/licenses/bsd-license.php)
 */

namespace yii2tech\excelgrid;

use yii\helpers\FileHelper;
use yii\i18n\Formatter;
use PHPExcel;
use PHPExcel_IOFactory;
use Yii;
use yii\base\Component;
use yii\base\InvalidConfigException;
use yii\di\Instance;

/**
 * ExcelGrid allows export of data provider into Excel document via [[PHPExcel]] library.
 * It provides interface, which is similar to [[\yii\grid\GridView]] widget.
 *
 * Example:
 *
 * ```php
 * $exporter = new ExcelGrid([
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
 * @see http://phpexcel.codeplex.com/documentation
 * @see PHPExcel
 *
 * @property array|Formatter $formatter the formatter used to format model attribute values into displayable texts.
 *
 * @author Paul Klimov <klimov.paul@gmail.com>
 * @since 1.0
 */
class ExcelGrid extends Component
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
     * @var string writer type (format type). If not set will be determined automatically.
     * Supported values:
     *
     * - 'Excel5'
     * - 'Excel2007'
     * - 'OpenDocument'
     * - 'CSV'
     * - 'PDF'
     * - 'HTML'
     */
    public $writerType;

    /**
     * @var int current sheet row index.
     */
    protected $rowIndex;

    /**
     * @var PHPExcel|null Excel document representation instance.
     */
    private $_document;
    /**
     * @var array|Formatter the formatter used to format model attribute values into displayable texts.
     * This can be either an instance of [[Formatter]] or an configuration array for creating the [[Formatter]]
     * instance. If this property is not set, the "formatter" application component will be used.
     */
    private $_formatter;


    /**
     * @return PHPExcel Excel document representation instance.
     */
    public function getDocument()
    {
        if (!is_object($this->_document)) {
            $this->_document = new PHPExcel();
        }
        return $this->_document;
    }

    /**
     * @param PHPExcel|null $document Excel document representation instance.
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
                $this->_formatter = Instance::ensure($this->_formatter, Formatter::className());
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
                    'class' => DataColumn::className(),
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
            'class' => DataColumn::className(),
            'grid' => $this,
            'attribute' => $matches[1],
            'format' => isset($matches[3]) ? $matches[3] : 'text',
            'label' => isset($matches[5]) ? $matches[5] : null,
        ]);
    }

    /**
     * @param array $properties list of document properties in format: name => value
     * @return $this self reference
     * @see \PHPExcel_DocumentProperties
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
     * @param \yii\data\DataProviderInterface $dataProvider the data provider for the document
     * @return $this self reference
     */
    public function dataProvider($dataProvider)
    {
        $this->dataProvider = $dataProvider;
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
        $columnIndex = 'A';
        foreach ($this->columns as $column) {
            /* @var $column Column */
            $column->renderHeaderCell($columnIndex . $this->rowIndex);
            $columnIndex++;
        }
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
     * Saves the document into a file.
     * @param string $filename name of the output file.
     */
    public function save($filename)
    {
        $filename = Yii::getAlias($filename);

        $writerType = $this->writerType;
        if ($writerType === null) {
            $fileExtension = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
            if ($fileExtension === 'xlsx') {
                $writerType = 'Excel2007';
            } else {
                $writerType = 'Excel5';
            }
        }

        $fileDir = strtolower(pathinfo($filename, PATHINFO_DIRNAME));
        FileHelper::createDirectory($fileDir);

        $objWriter = PHPExcel_IOFactory::createWriter($this->getDocument(), $writerType);
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
        $tempFileName = tempnam(Yii::getAlias('@runtime'), 'ExcelGridTemp_');
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