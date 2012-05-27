<?php
abstract class DataProcessor {
	protected $_alphabet = '';
	protected $_returnData = array();

	public function __construct() {
		$this->_alphabet = range('A', 'Z');
	}
}
