<?php

namespace Excelsior;

use \PHPExcel_Worksheet;

class Worksheet extends Base {

    public function __construct(PHPExcel_Worksheet $worksheet, $parent) {
        $this->component = $worksheet;
        $this->parent = $parent;
    }

    public function getCell($coordinate, $row = null) {
        if (is_int($coordinate) && $row !== null) {
            return new Cell($this->component->getCellByColumnAndRow($coordinate, $row), $this);
        }

        return new Cell($this->component->getCell($coordinate), $this);
    }
}