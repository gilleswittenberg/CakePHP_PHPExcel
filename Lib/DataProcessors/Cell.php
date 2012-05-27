<?php
App::uses('DataProcessor', 'PHPExcel.Lib/DataProcessors');
App::uses('IDataProcessor', 'PHPExcel.Lib/DataProcessors');
class Cell extends DataProcessor implements IDataProcessor {
	public function processData($data, $startColumn = 0, $startRow = 1) {
		$cell = $this->_alphabet[$startColumn] . $startRow;
		if (is_array($data)) {
			$value = array_shift($data);
		} else {
			$value = $data;
		}
		return array($cell => $value);
	}
}
