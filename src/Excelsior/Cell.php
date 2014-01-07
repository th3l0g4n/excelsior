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
        $this->getParent()->getStyle($this->getRangeCoordinates())->getFont()->applyFromArray($config);
        return $this;
    }

    /**
     * Shortcut to set Cell-Fill
     *
     * @param array $config
     */
    public function setFill(array $config) {
        $this->getParent()->getStyle($this->getRangeCoordinates())->getFill()->applyFromArray($config);
        return $this;
    }

    public function setBorders(array $config) {
        $this->getParent()->getStyle($this->getRangeCoordinates())->getBorders()->applyFromArray($config);
        return $this;
    }

    /**
     * Merge given Cell with this one
     *
     * @param Cell $cell
     * @return $this
     */
    public function merge(Cell $cell) {
        if (($range = $this->isMerged()) !== false) {
            $this->getParent()->unmergeCells($range);
        }

        $this->getParent()->mergeCells($this->generateCellRange($this, $cell));

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
            if ($this->isInRange($range)) return $range;
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
        $this->getComponent()->setValue($value);
        return $this;
    }

    protected function generateCellRange(Cell $cell1, Cell $cell2) {
        $cols = array();
        $cols[] = $cell1->getColumnIndex();
        $cols[] = $cell2->getColumnIndex();
        sort($cols);

        $rows = array();
        $rows[] = $cell1->getRow();
        $rows[] = $cell2->getRow();
        sort($rows);

        $colStart = \PHPExcel_Cell::stringFromColumnIndex($cols[0]);
        $colEnd = \PHPExcel_Cell::stringFromColumnIndex($cols[1]);

        return $colStart . $rows[0] . ':' . $colEnd . $rows[1];
    }

    protected function getRangeCoordinates() {
        if (($range = $this->isMerged()) !== false) {
         return $range;
        }

        return $this->getCoordinate();
    }
}