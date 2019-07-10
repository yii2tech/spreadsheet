<?php
/**
 * @link https://github.com/yii2tech
 * @copyright Copyright (c) 2015 Yii2tech
 * @license [New BSD License](http://www.opensource.org/licenses/bsd-license.php)
 */

namespace yii2tech\spreadsheet;

use Closure;
use yii\base\BaseObject;

/**
 * Column is the base class of all {@see Spreadsheet}} column classes.
 *
 * @author Paul Klimov <klimov.paul@gmail.com>
 * @since 1.0
 */
class Column extends BaseObject
{
    /**
     * @var Spreadsheet the exporter object that owns this column.
     */
    public $grid;
    /**
     * @var string the header cell content.
     */
    public $header;
    /**
     * @var string the footer cell content.
     */
    public $footer;
    /**
     * @var callable This is a callable that will be used to generate the content of each cell.
     * The signature of the function should be the following: `function ($model, $key, $index, $column)`.
     * Where `$model`, `$key`, and `$index` refer to the model, key and index of the row currently being rendered
     * and `$column` is a reference to the {@see Column} object.
     */
    public $content;
    /**
     * @var bool whether this column is visible. Defaults to true.
     */
    public $visible = true;
    /**
     * @var array the column dimension options. Each option name will be converted into a 'setter' method of {@see \PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension}.
     * @see \PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension for details on how style configuration is processed.
     */
    public $dimensionOptions = [];
    /**
     * @var array the style for the header cell.
     * @see \PhpOffice\PhpSpreadsheet\Style\Style::applyFromArray() for details on how style configuration is processed.
     */
    public $headerOptions = [];
    /**
     * @var array|\Closure the style for the data cell This can either be an array of style
     * configuration or an anonymous function ({@see Closure}) that returns such an array.
     * The signature of the function should be the following: `function ($model, $key, $index, $column)`.
     * Where `$model`, `$key`, and `$index` refer to the model, key and index of the row currently being rendered
     * and `$column` is a reference to the {@see Column} object.
     * A function may be used to assign different attributes to different rows based on the data in that row.
     *
     * @see \PhpOffice\PhpSpreadsheet\Style\Style::applyFromArray() for details on how style configuration is processed.
     * @see \PhpOffice\PhpSpreadsheet\Style\Alignment::applyFromArray() for details on how 'alignment' configuration is processed.
     */
    public $contentOptions = [];
    /**
     * @var array the style for the footer cell.
     * @see \PhpOffice\PhpSpreadsheet\Style\Style::applyFromArray() for details on how style configuration is processed.
     * @see \PhpOffice\PhpSpreadsheet\Style\Alignment::applyFromArray() for details on how 'alignment' configuration is processed.
     */
    public $footerOptions = [];
    /**
     * @var array the style for the filter cell.
     * @see \PhpOffice\PhpSpreadsheet\Style\Style::applyFromArray() for details on how style configuration is processed.
     * @see \PhpOffice\PhpSpreadsheet\Style\Alignment::applyFromArray() for details on how 'alignment' configuration is processed.
     */
    public $filterOptions = [];


    /**
     * Renders the header cell.
     * @param string $cell cell coordinates.
     */
    public function renderHeaderCell($cell)
    {
        $this->grid->renderCell($cell, $this->renderHeaderCellContent(), $this->headerOptions);
    }

    /**
     * Renders the footer cell.
     * @param string $cell cell coordinates.
     */
    public function renderFooterCell($cell)
    {
        $this->grid->renderCell($cell, $this->renderFooterCellContent(), $this->footerOptions);
    }

    /**
     * Renders a data cell.
     * @param string $cell cell coordinates.
     * @param mixed $model the data model being rendered
     * @param mixed $key the key associated with the data model
     * @param int $index the zero-based index of the data item among the item array returned by {@see GridView::dataProvider}.
     */
    public function renderDataCell($cell, $model, $key, $index)
    {
        if ($this->contentOptions instanceof Closure) {
            $style = call_user_func($this->contentOptions, $model, $key, $index, $this);
        } else {
            $style = $this->contentOptions;
        }

        $this->grid->renderCell($cell, $this->renderDataCellContent($model, $key, $index), $style);
    }

    /**
     * Renders the filter cell.
     * @param string $cell cell coordinates.
     */
    public function renderFilterCell($cell)
    {
        $this->grid->renderCell($cell, $this->renderFilterCellContent(), $this->filterOptions);
    }

    /**
     * Renders the header cell content.
     * The default implementation simply renders {@see Column::$header}.
     * This method may be overridden to customize the rendering of the header cell.
     * @return string the rendering result
     */
    public function renderHeaderCellContent()
    {
        return trim($this->header) !== '' ? $this->header : $this->grid->emptyCell;
    }

    /**
     * Renders the footer cell content.
     * The default implementation simply renders {@see Column::$footer}.
     * This method may be overridden to customize the rendering of the footer cell.
     * @return string the rendering result
     */
    public function renderFooterCellContent()
    {
        return trim($this->footer) !== '' ? $this->footer : $this->grid->emptyCell;
    }

    /**
     * Renders the data cell content.
     * @param mixed $model the data model
     * @param mixed $key the key associated with the data model
     * @param int $index the zero-based index of the data model among the models array returned by {@see Spreadsheet::$dataProvider}.
     * @return string the rendering result
     */
    public function renderDataCellContent($model, $key, $index)
    {
        if ($this->content === null) {
            return $this->grid->emptyCell;
        }
        return call_user_func($this->content, $model, $key, $index, $this);
    }

    /**
     * Renders the filter cell content.
     * The default implementation simply renders a space.
     * This method may be overridden to customize the rendering of the filter cell (if any).
     * @return string the rendering result
     */
    public function renderFilterCellContent()
    {
        return $this->grid->emptyCell;
    }
}