<?php
/**
 * PHPExcel TestSuite
 *
 * PHP 5
 *
 * Copyright 2012, Gilles Wittenberg (http://www.gilleswittenberg.com)
 *
 * Licensed under The MIT License
 * Redistributions of files must retain the above copyright notice.
 *
 * @copyright     Copyright (c) 2012, Gilles Wittenberg
 * @license       MIT License (http://www.opensource.org/licenses/mit-license.php)
 */

/**
 * PHPExcelTestSuite
 *
 * @package HTMLPurifier.Test
 * @author 	Gilles Wittenberg
 */
class PHPExcelTestSuite extends CakeTestSuite {

/**
 * Suite
 *
 * @return 	CakeTestSuite
 * @static
 * @access 	public
 */
	public static function suite() {
		$suite = new CakeTestSuite('PHPExcel testsuite');
		$suite->addTestDirectoryRecursive(App::pluginPath('PHPExcel') . 'Test' . DS . 'Case');
		return $suite;
	}
}
