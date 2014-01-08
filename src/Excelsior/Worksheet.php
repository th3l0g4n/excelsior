<?php

namespace Excelsior;

use \PHPExcel_Worksheet;

class Worksheet extends Base {

    public function __construct(PHPExcel_Worksheet $worksheet, $parent) {
        $this->component = $worksheet;
        $this->parent = $parent;
    }

    /**
     * Selects a Cell either by coordinate or row- and columnnumber
     *
     * @param mixed $coordinate either Columnetter or Columnnumber (index 0)
     * @param mixed $row Rownumber
     * @return Cell
     */
    public function getCell($coordinate, $row = null) {
        if (is_int($coordinate) && $row !== null) {
            return new Cell($this->component->getCellByColumnAndRow($coordinate, $row), $this);
        }

        return new Cell($this->component->getCell($coordinate), $this);
    }
}