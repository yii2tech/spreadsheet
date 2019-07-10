<?php
/**
 * @link https://github.com/yii2tech
 * @copyright Copyright (c) 2015 Yii2tech
 * @license [New BSD License](http://www.opensource.org/licenses/bsd-license.php)
 */

namespace yii2tech\spreadsheet;

use yii\base\Model;
use yii\data\ActiveDataProvider;
use yii\db\ActiveQueryInterface;
use yii\helpers\ArrayHelper;
use yii\helpers\Inflector;

/**
 * DataColumn is the default column type for the {@see Spreadsheet}.
 *
 * @author Paul Klimov <klimov.paul@gmail.com>
 * @since 1.0
 */
class DataColumn extends Column
{
    /**
     * @var string the attribute name associated with this column. When neither {@see Column::$content} nor {@see value}
     * is specified, the value of the specified attribute will be retrieved from each data model and displayed.
     *
     * Also, if {@see label} is not specified, the label associated with the attribute will be displayed.
     */
    public $attribute;
    /**
     * @var string label to be displayed in the {@see Column::$header} and also to be used as the sorting
     * link label when sorting is enabled for this column.
     * If it is not set and the models provided by the GridViews data provider are instances
     * of {@see \yii\db\ActiveRecord}, the label will be determined using {@see \yii\db\ActiveRecord::getAttributeLabel()}.
     * Otherwise {@see \yii\helpers\Inflector::camel2words()} will be used to get a label.
     */
    public $label;
    /**
     * @var string|\Closure an anonymous function or a string that is used to determine the value to display in the current column.
     *
     * If this is an anonymous function, it will be called for each row and the return value will be used as the value to
     * display for every data model. The signature of this function should be: `function ($model, $key, $index, $column)`.
     * Where `$model`, `$key`, and `$index` refer to the model, key and index of the row currently being rendered
     * and `$column` is a reference to the {@see DataColumn} object.
     *
     * You may also set this property to a string representing the attribute name to be displayed in this column.
     * This can be used when the attribute to be displayed is different from the {@see attribute} that is used for
     * sorting and filtering.
     *
     * If this is not set, `$model[$attribute]` will be used to obtain the value, where `$attribute` is the value of {@see attribute}.
     */
    public $value;
    /**
     * @var string|array in which format should the value of each data model be displayed as (e.g. `"raw"`, `"text"`, `"html"`,
     * `['date', 'php:Y-m-d']`). Supported formats are determined by the {@see Spreadsheet::$formatter} used by
     * the {@see Spreadsheet}. Default format is "raw" which will display value as it is.
     */
    public $format = 'raw';
    /**
     * @var string|array|bool the HTML code representing a filter input (e.g. a text field, a dropdown list)
     * that is used for this data column. This property is effective only when {@see Spreadsheet::$filterModel} is set.
     *
     * - If this property is not set, a text field will be generated as the filter input;
     * - If this property is an array, a dropdown list will be generated that uses this property value as
     *   the list options.
     * - If you don't want a filter for this data column, set this value to be false.
     */
    public $filter;


    /**
     * {@inheritdoc}
     */
    public function renderHeaderCellContent()
    {
        if ($this->header !== null || $this->label === null && $this->attribute === null) {
            return parent::renderHeaderCellContent();
        }

        $provider = $this->grid->dataProvider;

        if ($this->label === null) {
            if ($provider instanceof ActiveDataProvider && $provider->query instanceof ActiveQueryInterface) {
                /* @var $model Model */
                $model = new $provider->query->modelClass;
                $label = $model->getAttributeLabel($this->attribute);
            } else {
                $models = $provider->getModels();
                if (($model = reset($models)) instanceof Model) {
                    /* @var $model Model */
                    $label = $model->getAttributeLabel($this->attribute);
                } else {
                    $label = Inflector::camel2words($this->attribute);
                }
            }
        } else {
            $label = $this->label;
        }

        return $label;
    }

    /**
     * {@inheritdoc}
     */
    public function renderFilterCellContent()
    {
        if (is_string($this->filter)) {
            return $this->filter;
        }

        return parent::renderFilterCellContent();
    }

    /**
     * Returns the data cell value.
     * @param mixed $model the data model
     * @param mixed $key the key associated with the data model
     * @param int $index the zero-based index of the data model among the models array returned by {@see Spreadsheet::$dataProvider}.
     * @return string the data cell value
     */
    public function getDataCellValue($model, $key, $index)
    {
        if ($this->value !== null) {
            if (is_string($this->value)) {
                return ArrayHelper::getValue($model, $this->value);
            }
            return call_user_func($this->value, $model, $key, $index, $this);
        } elseif ($this->attribute !== null) {
            return ArrayHelper::getValue($model, $this->attribute);
        }

        return null;
    }

    /**
     * {@inheritdoc}
     */
    public function renderDataCellContent($model, $key, $index)
    {
        if ($this->content === null) {
            $value = $this->getDataCellValue($model, $key, $index);
            if ($value === null) {
                return $this->grid->nullDisplay;
            }
            return $this->grid->formatter->format($value, $this->format);
        }

        return parent::renderDataCellContent($model, $key, $index);
    }
}