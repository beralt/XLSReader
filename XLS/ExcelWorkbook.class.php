<?php

/**
 * Excel workbook
 */

class ExcelWorkbook {

	protected
		$strings = array(),
		$stream = null;

	public function __construct(OLEStream $stream) {
		$this->stream = $stream;

		$this->loadWorksheets();
	}

	protected function loadWorksheets() {
		$sheets = array();

		/* scan for sheets */
		while(!$this->stream->feof()) {
			$record_number = $this->_ushort($this->stream->read(2));
			$record_length = $this->_ushort($this->stream->read(2));
			if($record_length > 0)
				$record = $this->stream->read($record_length);
			else
				continue;

			if($record_number == ExcelSymbols::BOF) {
				$i = 0;
				$version = $this->_ushort($this->read($record, 2, $i));
				$type = $this->_ushort($this->read($record, 2, $i));
				debug('Creating worksheet for BOF version :'.$version.' type '.sprintf('0x%x', $type));
				if($type == 0x10) {
					/* worksheet */
					$sheets[] = new ExcelWorksheet($this->stream, $this, $version);
				}
			} elseif($record_number == ExcelSymbols::SST) {
				debug('Found shared string table');
				$this->setupSharedStringTable($record, $record_length);
			}
		}

		$this->sheets = $sheets;
	}

	public function getWorksheets() {
		return $this->sheets;
	}

	protected function setupSharedStringTable($record, $record_length) {
		$i = 0;
		$j = 0;
		$total = $this->_ulong($this->read($record, 4, $i));
		$unique = $this->_ulong($this->read($record, 4, $i));
		debug('Found SST: '.$total.' ('.$unique.' unique) length '.$record_length);
		/* start reading the strings */
		for(; $i < $record_length; ) {
			$nr_chars = $this->_ushort($this->read($record, 2, $i));

			$grbit = $this->_byte($this->read($record, 1, $i));

			$is_ascii = (($grbit & 0x01) == 0);
			$rich_string_follows = (($grbit & 0x08) > 0);
			$is_extended_string = (($grbit & 0x04) > 0);

			if($is_extended_string) {
				$run_length = $this->_ulong($this->read($record, 4, $i));
				debug('Is extended string');
			}

			if($rich_string_follows) {
				$nr_runs = $this->_ushort($this->read($record, 2, $i));
				debug('Rich string with '.$nr_runs.' runs');
			}

			if($is_ascii) {
				$string = $this->read($record, $nr_chars, $i);
				$this->strings[$j++] = $string;
			} else {
				$string = $this->read($record, $nr_chars * 2, $i);
				$string = mb_convert_encoding($string, 'UTF-8', 'UTF-16LE');
				$this->strings[$j++] = $string;
				debug('Found UTF16-LE ['.$grbit.']: '.$string);
			}

			if($rich_string_follows)
				$i += $nr_runs * 4;
		}

		debug('Added '.$j.' strings');
	}

	public function getSharedString($index) {
		if(isset($this->strings[$index]))
			return $this->strings[$index];
		return '';
	}

	protected function read($str, $c, &$i) {
		$t = substr($str, $i, $c);
		$i += $c;
		return $t;
	}

	protected function _ushort($buf) {
		list(, $tmp) = unpack('v', $buf);
		return $tmp;
	}

	protected function _ulong($buf) {
		list(, $tmp) = unpack('V', $buf);
		return $tmp;
	}

	protected function _byte($buf) {
		list(, $tmp) = unpack('c', $buf);
		return $tmp;
	}

}
