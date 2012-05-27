<?php
App::uses('Cell', 'PHPExcel.Lib/DataProcessors');
class DataProcessorCellTest extends CakeTestCase {

	public function setUp() {
		$this->cell = new Cell();
	}

	public function tearDown() {
		unset($this->cell);
	}

	public function testProcessDataAssociatedArray() {
		$value = 'value';
		$data = $this->cell->processData(array('key' => $value));
		$this->assertEquals($value, $data['A1']);
	}

	public function testProcessDataNumericArray() {
		$value = 'value';
		$data = $this->cell->processData(array($value));
		$this->assertEquals($value, $data['A1']);
	}

	public function testProcessDataString() {
		$value = 'value';
		$data = $this->cell->processData($value);
		$this->assertEquals($value, $data['A1']);
	}

	public function testStartColumnAndRow() {
		$value = 'value';
		$data = $this->cell->processData($value, 4, 9);
		$this->assertArrayHasKey('E9', $data);
		$this->assertEquals($data['E9'], $value);
	}
}
