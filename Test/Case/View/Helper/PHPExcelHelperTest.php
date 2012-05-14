<?php

App::uses('Controller', 'Controller');
App::uses('View', 'View');
App::uses('PHPExcelHelper', 'PHPExcel.View/Helper');

class PHPExcelHelperTest extends CakeTestCase {
    public $PHPExcel = null;

    public function setUp() {
        parent::setUp();
        $Controller = new Controller();
        $View = new View($Controller);
        $this->PHPExcel = new PHPExcelHelper($View);
    }

    public function testConstruct() {
		$this->assertInstanceOf('PHPExcelHelper', $this->PHPExcel);
    }

	public function testGenerate() {
		$this->assertFalse($this->PHPExcel->generate(array('headers' => array(), 'rows' => array())));
		$this->assertFalse($this->PHPExcel->generate(array(), 'title'));
	}
}
