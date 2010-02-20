<?php

/**
 * Excel worksheet
 */

class ExcelWorksheet {

	protected
		$strings = array(),
		$book = null,
		$stream = null;

	protected
		$cells = array(array()),
		$max_row = 0,
		$max_col = 0;

	public function __construct(OLEStream $stream, ExcelWorkbook $book, $version) {
		$this->stream = $stream;
		$this->book = $book;
		$this->parse();
	}

	protected function parse() {
		while(!$this->stream->feof()) {
			$record_number = $this->_ushort($this->stream->read(2));
			$record_length = $this->_ushort($this->stream->read(2));
			//debug('Found record '.sprintf('0x%x', $record_number).', length '.$record_length);
			if($record_length > 0)
				$record = $this->stream->read($record_length);
	
			switch($record_number) {
				case ExcelSymbols::EOF:
					return;
				break;
				case ExcelSymbols::SST:
					$this->setupSharedStringTable($record, $record_length);
				case ExcelSymbols::EXTSST:
					debug('Found external shared string table');
				case ExcelSymbols::LABEL:
					$this->parseLabel($record, $record_length);
				case ExcelSymbols::LABELSST:
					$this->parseLabel($record, $record_length, true);
				case ExcelSymbols::ROW:
					//
				break;
				default:
					continue;
				break;
			}
		}
	}

	protected function setupSharedStringTable($record, $record_length) {
		$i = 0;
		$j = 0;
		$total = $this->_ulong($this->read($record, 4, $i));
		$unique = $this->_ulong($this->read($record, 4, $i));
		//debug('Found SST: '.$total.' ('.$unique.' unique) length '.$record_length);
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

	protected function parseBOF($record, $length) {
		$i = 0;
		$version = $this->_ushort($this->read($record, 2, $i));
		$type = $this->_ushort($this->read($record, 2, $i));
		debug('Found BOF version :'.$version.' type '.sprintf('0x%x', $type));
	}

	protected function parseLabel($record, $length, $is_sst = false) {
		if($is_sst) {
			$i = 0;
			$row = $this->_ushort($this->read($record, 2, $i));
			$col = $this->_ushort($this->read($record, 2, $i));
			$ixfe = $this->_ushort($this->read($record, 2, $i));
			$isst = $this->_ulong($this->read($record, 4, $i));
			$text = $this->book->getSharedString($isst);
			debug('Found LABELSST at ('.$row.','.$col.') ixfe '.$ixfe.' isst '.$isst.' text '.$text);
			$this->setCell($row, $col, $text);
		}
	}

	public function debugCells() {
		debug('Size: '.$this->max_row.','.$this->max_col);
		/*
		for($i = 0; $i < $this->max_row; ++$i) {
			if(!isset($this->cells[$i])) {
				echo str_repeat('|   ', $this->max_col)."|\n";
			} else {
				for($j = 0; $j < $this->max_col; ++$j) {
					if(isset($this->cells[$i][$j]))
						echo '| x ';
					else
						echo '|   ';
				}
				echo '|'."\n";
			}
		}
		*/
	}

	protected function setCell($row, $col, $str) {
		if(!isset($this->cells[$row]))
			$this->cells[$row] = array($col => $str);
		else
			$this->cells[$row][$col] = $str;

		if($row > $this->max_row)
			$this->max_row = $row;

		if($col > $this->max_col)
			$this->max_col = $col;
	}

	public function getCell($row, $col) {
		if(isset($this->cells[$row])) {
			if(isset($this->cells[$row][$col]))
				return $this->cells[$row][$col];
		}
		return null;
	}

	public function getColumnRange() {
		return $this->max_col;
	}

	public function getRowRange() {
		return $this->max_row;
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
