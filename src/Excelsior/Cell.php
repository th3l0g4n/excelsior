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
     * @param int $offset select next $offset cell
     * @return Cell
     */
    public function left($offset = 1) {
        $component = $this->getParent()->getComponent();
        return new self($component->getCellByColumnAndRow($this->getColumnIndex() - $offset, $this->getRow()), $this->getParent());
    }

    /**
     * Select Cell right to this one
     *
     * @param int $offset select next $offset cell
     * @return Cell
     */
    public function right($offset = 1) {
        $component = $this->getParent()->getComponent();
        return new self($component->getCellByColumnAndRow($this->getColumnIndex() + $offset, $this->getRow()), $this->getParent());
    }

    /**
     * Select Cell above this one
     *
     * @param int $offset select next $offset cell
     * @return Cell
     */
    public function up($offset = 1) {
        $component = $this->getParent()->getComponent();
        return new self($component->getCellByColumnAndRow($this->getColumnIndex() , $this->getRow() - $offset), $this->getParent());
    }

    /**
     * Select Cell below this one
     *
     * @param int $offset select next $offset cell
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
        $cell->unmerge();
        $this->unmerge();
        $this->getParent()->mergeCells($this->generateCellRange($this, $cell));

        return $this;
    }

    /**
     * If this Cell is part of a mergerange - unmerge it
     *
     * @return Cell
     */
    public function unmerge() {
        if (($range = $this->isMerged()) !== false) {
            $this->getParent()->unmergeCells($range);
        }

        return $this;
    }

    /**
     * Extend mergerange to coordinate of given cell if already merged or simple with it if not.
     *
     * @param Cell $cell
     * @return Cell
     */
    public function append(Cell $cell) {
        $cell->unmerge();

        if (($range = $this->isMerged()) !== false) {
            $this->unmerge();

            $boundaries = $this->sortBoundaries(\PHPExcel_Cell::getRangeBoundaries($range));

            $boundaries[0][] = $cell->getColumn();
            $boundaries[1][] = $cell->getRow();
            sort($boundaries[0]);
            sort($boundaries[1]);

            $startCoordinate = $boundaries[0][0] . $boundaries[1][0];
            $endCoordinate = $boundaries[0][count($boundaries[0]) - 1] . $boundaries[1][count($boundaries[1]) - 1];
            $this->getParent()->mergeCells($startCoordinate . ':' . $endCoordinate);

            return $this;
        }

        $this->merge($cell);
        return $this;
    }

    /**
     * Check whether this Cell is part of a Mergerange
     *
     * @return mixed false if not merged or range if is.
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

    /**
     * Generate rangecoordinates for given cells
     *
     * @param Cell $cell1
     * @param Cell $cell2
     * @return string Range coordinates
     */
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

    /**
     * Returns rangecoordinates if Cell is merged or just his own if not.
     *
     * @return string
     */
    protected function getRangeCoordinates() {
        if (($range = $this->isMerged()) !== false) {
         return $range;
        }

        return $this->getCoordinate();
    }

    /**
     * Rearranges rangecoordinates for better handling
     *
     * @param array $boundaries returned by PHPExcel_Cell::getRangeBoundaries
     * @return array Index 0 array of columns; Index 1 array of rows
     */
    protected function sortBoundaries(array $boundaries) {
        $cols = array($boundaries[0][0], $boundaries[1][0]);
        $rows = array($boundaries[0][1], $boundaries[1][1]);

        return array($cols, $rows);
    }
}