<?php

namespace Excelsior;

abstract class Base {
    protected $component;
    protected $parent;

    public function __call($func, $args) {
        if (is_callable(array($this->component, $func))) {
            return call_user_func_array(array($this->component, $func), $args);
        }

        throw new \Exception('Method ' . $func . ' is not available');
    }

    public function getParent() {
        return $this->parent;
    }

    public function getComponent() {
        return $this->component;
    }
}