<?php

/**
 * Reads OLE containers
 */

class OLEReader {

	const
		STGTY_INVALID = 0,
		STGTY_STORAGE = 1,
		STGTY_STREAM = 2,
		STGTY_LOCKBYTES = 3,
		STGTY_PROPERTY = 4,
		STGTY_ROOT = 5;

	protected
		$block_size = null,
		$small_block_size = null,
		$nr_fat_sectors = 0,
		$dir_start = 0,
		$fd = null,
		$file = null;

	protected
		$streams = array();

	public function __construct($file) {
		$this->file = $file;
		if(is_readable($this->file)) {
			$this->fd = fopen($this->file, 'r');
			$this->parse();
		}
	}

	public function getStreams() {
		return $this->streams;
	}

	public function getBlockSize() {
		return $this->block_size;
	}

	public function getNextBlock($block) {
		if(isset($this->fat[$block]))
			return $this->fat[$block];
		return null;
	}

	protected function parse() {
		$streams = array();

		if(!$this->readHeader())
			throw new Exception('Unable to read header');

		/* directories are aligned on 128 bytes */
		$stats = fstat($this->fd);
		for($pos = 512 + $this->dir_start * $this->block_size; $pos < $stats[7]; $pos += 128) {
			fseek($this->fd, $pos);
			$name = fread($this->fd, 64);
			$name_length = $this->readUnsignedShort();
			$name = utf8_encode(str_replace("\x00", '', substr($name, 0, $name_length - 2)));
			$type = $this->readByte();
			$color = $this->readByte();

			$left_sib = $this->readUnsignedLong();
			$right_sib = $this->readUnsignedLong();
			$child = $this->readUnsignedLong();

			fseek($this->fd, 36, SEEK_CUR);

			$stream_start = $this->readUnsignedLong();
			$stream_size = $this->readUnsignedLong();

			switch($type) {
				case self::STGTY_STREAM:
					debug('Found stream '.$name.' starting at sector '.sprintf('0x%X', $stream_start).' of size '.$stream_size);
					$streams[] = new OLEStream($name, $this, $this->fd, $stream_start, $stream_size);
				break;
				case self::STGTY_ROOT:
					debug('Found root '.$name);
				break;
			}
		}

		$this->streams = $streams;
	}

	protected function readHeader() {
		$data = fread($this->fd, 8);
		$header = unpack('C8', $data);

		if($data == "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1") {
			/* ensure Intel byte-ordering */
			fseek($this->fd, 28);
			if(fread($this->fd, 2) != "\xFE\xFF")
				throw new Exception('Unable to read OLE container; only little endian encoding is supported');

			$this->block_size = pow(2, $this->readUnsignedShort());
			$this->small_block_size = pow(2, $this->readUnsignedShort());

			fseek($this->fd, 44);
			$this->nr_fat_sectors = $this->readUnsignedLong();
			$this->dir_start = $this->readUnsignedLong();

			fseek($this->fd, 56);
			$mini_fat_cutoff = $this->readUnsignedLong();
			$mini_fat_start = $this->readUnsignedLong();
			$nr_mini_fat = $this->readUnsignedLong();
			$dif_start = $this->readUnsignedLong();
			$nr_dif = $this->readUnsignedLong();

			debug(
				'block size              : '.$this->block_size."\n".
				'small_block size        : '.$this->small_block_size."\n".
				'nr FAT sectors          : '.$this->nr_fat_sectors."\n".
				'first sect in DIR chain : '.$this->dir_start."\n".
				'mini FAT cutoff         : '.$mini_fat_cutoff."\n".
				'mini FAT start          : '.sprintf('0x%x', $mini_fat_start)."\n".
				'nr mini FAT             : '.$nr_mini_fat."\n".
				'DIF start               : '.sprintf('0x%x', $dif_start)."\n".
				'nr DIF                  : '.$nr_dif."\n");

			/* read first 109 fat sectors */
			$blocks = array();
			for($i = 0; $i < 109; ++$i)
				$blocks[$i] = $this->readUnsignedLong();
			$this->blocks = $blocks;

			/* should read rest of table */

			$fat = array();
			for($i = 0; $i < $this->nr_fat_sectors; ++$i) {
				$pos = $blocks[$i] * $this->block_size + 512;
				fseek($this->fd, $pos);
				for($j = 0; $j < $this->block_size / 4; ++$j) {
					$addr = $this->readUnsignedLong();
					$fat[] = $addr;
				}
			}
			$this->fat = $fat;

			return true;
		}

		return false;
	}

	protected function readByte() {
		list(, $tmp) = unpack('c', fread($this->fd, 1));
		return $tmp;
	}

	protected function readUnsignedShort() {
		list(, $tmp) = unpack('v', fread($this->fd, 2));
		return $tmp;
	}

	protected function readUnsignedLong() {
		list(, $tmp) = unpack('V', fread($this->fd, 4));
		return $tmp;
	}
}


