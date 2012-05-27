<?php
interface IDataProcessor {
	public function processData($data, $startColumn = 0, $startRow = 1);
}
