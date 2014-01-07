<?php

namespace Excelsior;
use \PHPExcel;

class Excelsior extends Base {

    public function __construct() {
        $this->component = new PHPExcel();
    }

    public function load($file) {
        $this->component = \PHPExcel_IOFactory::load($file);
        return $this;
    }
}