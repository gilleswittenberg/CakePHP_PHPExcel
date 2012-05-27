<?php
App::uses('ArrayValues', 'PHPExcel.Lib/DataProcessors');
class DataProcessorArrayValuesTest extends CakeTestCase {

	public function setUp() {
		$this->arrayValues = new ArrayValues();
	}

	public function tearDown() {
		unset($this->cell);
	}

	public function testProcessDataArray() {
		$arr = array('val1', 'valTwo', 'val3');
		$data = $this->arrayValues->processData($arr);
		$this->assertArrayHasKey('A1', $data);
		$this->assertArrayHasKey('B1', $data);
		$this->assertArrayHasKey('C1', $data);
		$this->assertEquals('val3', $data['C1']);
	}

	public function testProcessDataStartColumnAndRow() {
		$arr = array('val1', 'valTwo', 'val3');
		$data = $this->arrayValues->processData($arr, 4, 9);
		$this->assertArrayHasKey('E9', $data);
		$this->assertArrayHasKey('F9', $data);
		$this->assertArrayHasKey('G9', $data);
		$this->assertEquals('val3', $data['G9']);
	}

	public function testProcessDataStartColumnAndRowAndDirection() {
		$arr = array('val1', 'valTwo', 'val3');
		$data = $this->arrayValues->processData($arr, 4, 9, 'VER');
		$this->assertArrayHasKey('E9', $data);
		$this->assertArrayHasKey('E10', $data);
		$this->assertArrayHasKey('E11', $data);
		$this->assertEquals('val3', $data['E11']);
	}
}
