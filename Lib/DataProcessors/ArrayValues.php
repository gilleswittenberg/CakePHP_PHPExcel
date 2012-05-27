<?php

App::uses('DataProcessor', 'PHPExcel.Lib/DataProcessors');
App::uses('IDataProcessor', 'PHPExcel.Lib/DataProcessors');

class ArrayValues extends DataProcessor implements IDataProcessor {

	public function processData($data, $startColumn = 0, $startRow = 1, $direction = 'HOR') {
		$return = array();
		$column = $startColumn;
		$row = $startRow;
		foreach ($data as $value) {
			$cell = $this->_alphabet[$column] . $row;
			$return[$cell] = $value;
			if ($direction == 'HOR') {
				$column++;
			} else {
				$row++;
			}
		}
		return $return;
	}
}
