<?php

namespace Excelsior;
use \PHPExcel;

class Excelsior extends Base {

    /**
     * If file is given it will be loaded instead of creating a blank worksheet
     *
     * @param null $file file to be loaded
     */
    public function __construct($file = null) {
        if ($file === null) {
            $this->component = new PHPExcel();
            return $this;
        }

        return $this->load($file);
    }

    /**
     * Get Sheet in Workbook by index
     *
     * @see PHPExcel::getSheet
     * @param int $index
     * @return Worksheet
     */
    public function getSheet($index = 0) {
        return new Worksheet($this->component->getSheet($index), $this);
    }

    /**
     * Load file for given path. Replaces currently used Workbook.
     *
     * @param $file Path to a spreadsheet to load
     * @return $this
     */
    public function load($file) {
        $this->component = \PHPExcel_IOFactory::load($file);
        return $this;
    }

    /**
     * @param $file Path to file
     * @param string $format Format in which file will be saved (default: Excel2007)
     * @throws \Exception if given Writer for format does not exist
     */
    public function save($file, $format = 'Excel2007') {
        $writerClass = '\\PHPExcel_Writer_' . $format;

        if (!class_exists($writerClass)) {
            throw new \Exception('Desired Writer ' . $format . ' does not exist!');
        }

        $writer = new $writerClass($this->component);
        $writer->save($file);
    }
}