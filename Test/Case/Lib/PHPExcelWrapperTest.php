<?php

App::uses('PHPExcelWrapper', 'PHPExcel.Lib');
App::uses('File', 'Utility');

class PHPExcelWrapperTest extends CakeTestCase {

	protected $_fixtureFiles = array();

    public function setUp() {
        parent::setUp();
    }

	public function tearDown() {
		$this->_cleanUpFixtureFiles();
	}

	protected function _cpFixtureFile($filename, $dest = null) {
		$dest = $dest ? $dest : WWW_ROOT . 'files';
		$file = new File(App::pluginPath('PHPExcel') . 'Test' . DS . 'Fixture' . DS . 'File' . DS . $filename);
		if ($file->copy($dest . DS . $filename, true)) {
			$this->_fixtureFiles[] = array(
				'filename' => $filename,
				'dest' => $dest
			);
		}
	}

	protected function _cleanUpFixtureFiles() {
		foreach ($this->_fixtureFiles as $file) {
			$file = new File($file['dest'] . DS . $file['filename']);
			$file->delete();
		}
	}

    public function testConstruct() {
		$PHPExcelWrapper = new PHPExcelWrapper();
		$this->assertInstanceOf('PHPExcelWrapper', $PHPExcelWrapper);
		$this->assertInstanceOf('PHPExcel', $PHPExcelWrapper->_PHPExcelObj);
	}

    public function testConstructWithFilename() {
		$filename = '1.xls';
		$this->_cpFixtureFile($filename);
		$PHPExcelWrapper = new PHPExcelWrapper('files' . DS . $filename);
		$this->assertInstanceOf('PHPExcel', $PHPExcelWrapper->_PHPExcelObj);
    }

	/**
	 * @expectedException PHPExcelException
	 */
	public function testConstructWithNonExistingFilename() {
		$PHPExcelWrapper = new PHPExcelWrapper('non_existing_file.xls');
    }

	public function testSave() {
		$PHPExcelWrapper = new PHPExcelWrapper();
		$PHPExcelWrapper->save();
		$this->assertTrue(file_exists(WWW_ROOT . 'files' . DS . 'CakePHP_PHPExcel.xls'));
		// clean up
		$file = new File(WWW_ROOT . 'files' . DS . 'CakePHP_PHPExcel.xls');
		$file->delete();
	}

	public function testSaveWithDifferentOutput() {
		$filename = 'test';
		$PHPExcelWrapper = new PHPExcelWrapper();
		$PHPExcelWrapper->setOutput('filename', $filename);
		$PHPExcelWrapper->save();
		$this->assertTrue(file_exists(WWW_ROOT . 'files' . DS . $filename . '.xls'));
		// clean up
		$file = new File(WWW_ROOT . 'files' . DS . $filename . '.xls');
		$file->delete();
	}

	public function testSaveToNonExistingFolder() {
		$filename = 'test';
		$dir = WWW_ROOT . 'files' . DS . 'non_existing' . DS . 'folder';
		$PHPExcelWrapper = new PHPExcelWrapper();
		$PHPExcelWrapper->setOutput('filename', $filename);
		$PHPExcelWrapper->setOutput('dir', $dir);
		$PHPExcelWrapper->save();
		$this->assertTrue(file_exists($dir . DS . $filename . '.xls'));
		// clean up
		$folder = new Folder(WWW_ROOT . 'files' . DS . 'non_existing');
		$folder->delete();
	}

	public function testSaveWithOverwriteFalse() {
		$PHPExcelWrapper = new PHPExcelWrapper();
		$this->assertTrue($PHPExcelWrapper->save());
		$this->assertFalse($PHPExcelWrapper->save(false));
		// clean up
		$file = new File(WWW_ROOT . 'files' . DS . 'CakePHP_PHPExcel.xls');
		$file->delete();
	}

	public function testSetProperty() {
		$title = 'Test Excel File';
		$PHPExcelWrapper = new PHPExcelWrapper();
		$properties = $PHPExcelWrapper->setProperty('title', $title);
		$this->assertEqual($properties->getTitle(), $title);
	}

	/**
	 * @expectedException PHPExcelException
	 */
	public function testSetPropertyNonExisting() {
		$PHPExcelWrapper = new PHPExcelWrapper();
		$properties = $PHPExcelWrapper->setProperty('non_existing', 'value string');
	}

	/**
	 * @expectedException PHPExcelException
	 */
	public function testSetCellValuesWithNonExistingClass() {
		$PHPExcelWrapper = new PHPExcelWrapper();
		$PHPExcelWrapper->setCellValues('NonExistingClass');
	}

	public function testSetCellValuesChaining() {
		$PHPExcelWrapper = new PHPExcelWrapper();
		$return = $PHPExcelWrapper->setCellValues('Cell', array());
		$this->assertEquals($PHPExcelWrapper, $return);
	}

	public function testSetCurrentStyle() {
		$PHPExcelWrapper = new PHPExcelWrapper();
		$currentStyle = $PHPExcelWrapper->setCurrentStyle('font', 'bold')->getCurrentStyle();
		$this->assertInternalType('array', $currentStyle['font']);
		$this->assertTrue($currentStyle['font']['bold']);
	}

	public function clearCurrentStyle() {
		$PHPExcelWrapper = new PHPExcelWrapper();
		$currentStyle = $PHPExcelWrapper->setCurrentStyle('font', 'bold');
		$currentStyle = $PHPExcelWrapper->clearCurrentStyle()->getCurrentStyle();
		$this->assertTrue(empty($currentStyle));
	}

	public function clearCurrentStyleFamily() {
		$PHPExcelWrapper = new PHPExcelWrapper();
		$currentStyle = $PHPExcelWrapper->setCurrentStyle('font', 'bold');
		$currentStyle = $PHPExcelWrapper->setCurrentStyle('fill', 'color', 'green');
		$currentStyle = $PHPExcelWrapper->clearCurrentStyle('font')->getCurrentStyle();
		$this->assertFalse(isset($currentStyle['font']));
	}
	public function clearCurrentStyleKey() {
		$PHPExcelWrapper = new PHPExcelWrapper();
		$currentStyle = $PHPExcelWrapper->setCurrentStyle('font', 'bold');
		$currentStyle = $PHPExcelWrapper->setCurrentStyle('font', 'size', 12);
		$currentStyle = $PHPExcelWrapper->clearCurrentStyle('font', 'bold')->getCurrentStyle();
		$this->assertFalse(isset($currentStyle['font']['bold']));
	}
}
