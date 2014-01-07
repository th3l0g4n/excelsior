<?php

namespace Excelsior;

use \PHPExcel_Cell;

class Cell extends Base {

    public function __construct(PHPExcel_Cell $cell, $parent) {
        $this->component = $cell;
        $this->parent = $parent;
    }

    /**
     * Select Cell left to this one
     *
     * @param int $offset
     * @return Cell
     */
    public function left($offset = 1) {
        $component = $this->getParent()->getComponent();
        return new self($component->getCellByColumnAndRow($this->getColumnIndex() - $offset, $this->getRow()), $this->getParent());
    }

    /**
     * Select Cell right to this one
     *
     * @param int $offset
     * @return Cell
     */
    public function right($offset = 1) {
        $component = $this->getParent()->getComponent();
        return new self($component->getCellByColumnAndRow($this->getColumnIndex() + $offset, $this->getRow()), $this->getParent());
    }

    /**
     * Select Cell above this one
     *
     * @param int $offset
     * @return Cell
     */
    public function up($offset = 1) {
        $component = $this->getParent()->getComponent();
        return new self($component->getCellByColumnAndRow($this->getColumnIndex() , $this->getRow() - $offset), $this->getParent());
    }

    /**
     * Select Cell below this one
     *
     * @param int $offset
     * @return Cell
     */
    public function down($offset = 1) {
        $component = $this->getParent()->getComponent();
        return new self($component->getCellByColumnAndRow($this->getColumnIndex() , $this->getRow() + $offset), $this->getParent());
    }

    public function getColumnIndex() {
        return \PHPExcel_Cell::columnIndexFromString($this->getColumn()) - 1;
    }

    /**
     * Shortcut to set Cell-Font
     *
     * @param array $config
     */
    public function setFont(array $config) {
        $this->getStyle()->getFont()->applyFromArray($config);
    }

    /**
     * Shortcut to set Cell-Fill
     *
     * @param array $config
     */
    public function setFill(array $config) {
        $this->getStyle()->getFill()->applyFromArray($config);
    }

    /**
     * Merge given Cell with this one
     *
     * @param Cell $cell
     * @return $this
     */
    public function merge(Cell $cell) {
        $cols = array();
        $cols[] = $this->getColumnIndex();
        $cols[] = $cell->getColumnIndex();
        sort($cols);

        var_dump($cols);

        $rows = array();
        $rows[] = $this->getRow();
        $rows[] = $cell->getRow();
        sort($rows);

        var_dump($rows);

        $this->getParent()->mergeCellsByColumnAndRow($cols[0], $rows[0], $cols[1], $rows[1]);

        return $this;
    }

    /**
     * Check whether this Cell is part of a Mergerange
     *
     * @return bool
     */
    public function isMerged() {
        $cellRanges = $this->getParent()->getMergeCells();

        foreach ($cellRanges as $range) {
            if ($this->isInRange($range)) return true;
        }

        return false;
    }

    /**
     * 'Overridden' to keep fluent interface
     *
     * @param $value
     * @return $this
     */
    public function setValue($value) {
        $this->component->setValue($value);
        return $this;
    }
}