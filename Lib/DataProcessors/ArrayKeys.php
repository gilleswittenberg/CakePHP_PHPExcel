<?php

App::uses('DataProcessor', 'PHPExcel.Lib/DataProcessors');
App::uses('IDataProcessor', 'PHPExcel.Lib/DataProcessors');

class ArrayKeys extends DataProcessor implements IDataProcessor {

	public function processData($data, $startColumn = 0, $startRow = 1, $direction = 'HOR') {
		$return = array();
		$column = $startColumn;
		$row = $startRow;
		foreach ($data as $key => $value) {
			$cell = $this->_alphabet[$column] . $row;
			$return[$cell] = $key;
			if ($direction == 'HOR') {
				$column++;
			} else {
				$row++;
			}
		}
		return $return;
	}
}
