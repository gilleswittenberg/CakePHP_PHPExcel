<?php

App::uses('Folder', 'Utility');
App::uses('PHPExcelException', 'PHPExcel.Error');
App::uses('Cell', 'PHPExcel.Lib/DataProcessors');
App::uses('ArrayKeys', 'PHPExcel.Lib/DataProcessors');
App::uses('ArrayValues', 'PHPExcel.Lib/DataProcessors');
require_once APP::pluginPath('PHPExcel') . DS . 'Vendor' . DS . 'PHPExcel' . DS . 'IOFactory.php';
require_once APP::pluginPath('PHPExcel') . DS . 'Vendor' . DS . 'PHPExcel' . DS . 'PHPExcel.php';

class PHPExcelWrapper {

	public $_PHPExcelObj;
	protected $_sheet;
	protected $_currentStyle = array();
	protected $_dataProcessors = array();
	protected $_outputDir; // make sure this directory is writable
	protected $_outputFilename = 'CakePHP_PHPExcel';
	protected $_outputExt = 'xls';

	public function __construct($filename = null) {
		// set outputDir
		$this->_outputDir = WWW_ROOT . 'files' . DS;
		// create PHPExcel instance
		if (is_string($filename)) {
			try {
				$this->_PHPExcelObj = PHPExcel_IOFactory::load($filename);
			} catch (Exception $e) {
				throw new PHPExcelException('File ' . $filename . ' does not exist.');
			}
		} else {
			$this->_PHPExcelObj = new PHPExcel();
		}
		// set activeSheet
		$this->_PHPExcelObj->setActiveSheetIndex(0);
	}

	public function setProperty($property, $value) {
		$method = 'set' . $property;
		// reference of PHPExcel_DocumentProperties instance
		$properties = $this->_PHPExcelObj->getProperties();
		if (method_exists($properties, $method)) {
			$properties->$method($value);
		} else {
			throw new PHPExcelException('Method set' . $property  . ' does not exist on PHPExcel_DocumentProperties.');
		}
		return $properties;
	}

	public function setCurrentStyle($family, $key, $value = true) {
		$this->_currentStyle[$family][$key] = $value;
		return $this;
	}

	public function getCurrentStyle() {
		return $this->_currentStyle;
	}

	public function clearCurrentStyle($family = null, $key = null) {
		if (!$family) {
			$this->_currentStyle = array();
		} else if (!$key) {
			unset($this->_currentStyle[$family]);
		} else {
			if (isset($this->_currentStyle[$family][$key])) {
				unset($this->_currentStyle[$family][$key]);
			}
		}
		return $this;
	}

	public function setCellValues($dataProcessor, $data = array(), $startColumn = 0, $startRow = 1, $direction = 'HOR') {
		// get processed data
		$processedData = $this->_getDataProcessor($dataProcessor)->processData($data, $startColumn, $startRow, $direction);
		// write data to cached sheet
		$this->_sheet = $this->_PHPExcelObj->getActiveSheet();
		foreach ($processedData as $cell => $value) {
			$this->_sheet->setCellValue($cell, $value);
			$this->_sheet->getStyle($cell)->applyFromArray($this->_currentStyle);
		}
		return $this;
	}

	protected function _getDataProcessor($dataProcessor) {
		if (!isset($this->_dataProcessors[$dataProcessor])) {
			if (class_exists($dataProcessor) && in_array('IDataProcessor', class_implements($dataProcessor))) {
				$this->_dataProcessors[$dataProcessor] = new $dataProcessor();
			} else {
				throw new PHPExcelException($dataProcessor  . ' does not exist or does not implement IDataProcessor.');
			}
		}
		return $this->_dataProcessors[$dataProcessor];
	}

	public function save($overwrite = true) {
		// check for overwriting
		if (!$overwrite && file_exists($this->_getOutput())) {
			return false;
		}
		// create dir and file
		$this->_mkOutputDir();
		$objWriter = PHPExcel_IOFactory::createWriter($this->_PHPExcelObj, 'Excel5');
		try {
			$objWriter->save($this->_getOutput());
		} catch (Exception $e) {
			throw new PHPExcelException('Failed saving File ' . $this->_getOutput() . '.');
		}
		return true;
	}

	public function setOutput($name, $value) {
		switch ($name) {
			case 'dir':
				$this->_outputDir = $value;
				if (substr($this->_outputDir, -1) != DS) {
					$this->_outputDir .= DS;
				}
				return true;
			case 'filename':
				$this->_outputFilename = $value;
				return true;
			case 'ext':
				$this->_outputExt = $value;
				return true;
			default:
				return false;
		}
	}

	protected function _getOutput() {
		return $this->_outputDir . $this->_outputFilename . '.' . $this->_outputExt;
	}

	protected function _mkOutputDir() {
		$mode = 0775;
		$folder = new Folder($this->_outputDir, true, $mode);
		$folder->chmod($mode);
	}
}
