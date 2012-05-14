<?php
App::uses('AppHelper', 'View/Helper');
require_once(APP::pluginPath('PHPExcel') . DS . 'Vendor' . DS . 'PHPExcel' . DS . 'IOFactory.php');
require_once(APP::pluginPath('PHPExcel') . DS . 'Vendor' . DS . 'PHPExcel' . DS . 'PHPExcel.php');

class PHPExcelHelper extends AppHelper {
	/**
	 * XLS object
	 *
	 * @var object
	 * @access protected
	 */
	protected $_objPHPExcel;

	/**
	 * XLS active sheet
	 *
	 * @var object
	 * @access protected
	 */
	protected $_sheet;

	/**
	 * Constructor method
	 *
	 * @param View $view
	 * @param array settings
	 * @return void
	 * @access public
	 */
	public function __construct(View $view, $settings = array()) {
		parent::__construct($view, $settings);
		$this->_objPHPExcel = new PHPExcel();
		$this->_sheet = $this->_objPHPExcel->getActiveSheet();
	}

	/**
	 * Generate xls file
	 *
	 * @param array $data
	 * @param string $title
	 * @return boolean
	 * @access public
	 */
	public function generate($data, $title = '', $filename = '', $path = null) {
		if (!$title && !$filename) {
			return false;
		}
		// PHP 5.3 $filename = $filename ?: $title;
		$filename = $filename ? $filename : $title;
		if (!array_key_exists('headers', $data) || !array_key_exists('headers', $data)) {
			return false;
		}
		$this->_headers($data['headers']);
		$this->_rows($data['rows']);
		$this->_output($filename, $path);
		return true;
	}

	/**
	 * Apply settings
	 *
	 * @param array $settings
	 * @return void
	 * @access protected
	 * @todo everything
	 */
	protected function _applySettings($settings = array()) {
		// this method could be called during construction and/or before generating
		//$this->_sheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
		//$this->_sheet->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
	}

	/**
	 * Set headers
	 *
	 * @param array $headers
	 * @return void
	 * @access protected
	 */
	protected function _headers($headers) {
		$this->_sheet->getStyle('1')->getFont()->setBold(true);
		$c = 0;
		$r = 1;
		foreach ($headers as $header) {
			$this->_sheet->setCellValueByColumnAndRow($c++, $r, $header);
		}
	}

	/**
	 * Set rows
	 *
	 * @param array $rows
	 * @return void
	 * @access protected
	 */
	protected function _rows($rows) {
		$r = 2;
		foreach ($rows as $row) {
			$c = 0;
			foreach ($row as $field => $value) {
				$this->_sheet->setCellValueByColumnAndRow($c++, $r, $value);
			}
			$r++;
		}
	}

	/**
	 * Output file
	 *
	 * @param string $filename
	 * @param string $path
	 * @return void
	 * @access protected
	 */
	protected function _output($filename, $path = null) {
		$objWriter = PHPExcel_IOFactory::createWriter($this->_objPHPExcel, 'Excel5');
		// this is buggy renders a string of pseudo random byte characters (Using OpenOffice on Mac OSX)
		//header('Content-type: application/vnd.ms-excel');
		//header('Content-Disposition: attachment;filename="' . $title . '.xls"');
		//header('Cache-Control: max-age=0');
		//$objWriter->save('php://output');
		if (substr($filename, -4) != '.xls') {
			$filename .= '.xls';
		}
		if ($path && substr($path, -1) != DS) {
			$path = $path . DS;
		}
		$filename = $path . $filename;
		$objWriter->save($filename);
	}
}
