<?php

namespace Excelsior;
use \PHPExcel;

class Excelsior extends Base {

    public function __construct($file = null) {
        if ($file === null) {
            $this->component = new PHPExcel();
            return $this;
        }

        return $this->load($file);
    }

    public function getSheet($index = 0) {
        return new Worksheet($this->component->getSheet($index), $this);
    }


    public function load($file) {
        $this->component = \PHPExcel_IOFactory::load($file);
        return $this;
    }

    public function save($file, $format = 'Excel2007') {
        $writerClass = '\\PHPExcel_Writer_' . $format;

        if (!class_exists($writerClass)) {
            throw new \Exception('Desired Writer ' . $format . ' does not exist!');
        }

        $writer = new $writerClass($this->component);
        $writer->save($file);
    }
}