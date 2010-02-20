<?php

/**
 * Represents a OLE stream
 */

class OLEStream {

	protected
		$name = '',
		$block = 0,
		$block_size = 512,
		$fd = null,
		$reader = null;

	protected
		$bfd = null;

	public function __construct($name, OLEReader $reader, $fd, $block, $size) {
		$this->name = $name;
		$this->reader = $reader;
		$this->fd = $fd;
		$this->block = $block;
		$this->block_size = $this->reader->getBlockSize();

		debug(
			'Stream name: '.$name."\n".
			'Start block: '.sprintf('0x%x', $block)."\n".
			'Size: '.$size." bytes\n");
		
		/* buffer */
		$this->bfd = fopen('php://temp', 'r+');
		fwrite($this->bfd, $this->getContents());
		rewind($this->bfd);
	}

	public function getName() {
		return $this->name;
	}

	public function feof() {
		return feof($this->bfd);
	}

	public function read($count) {
		return fread($this->bfd, $count);
	}

	public function getContents() {
		$content = '';
		$block_id = $this->block;

		while($block_id != 0xfffffffe) {
			$pos = $block_id * $this->block_size + 512;
			fseek($this->fd, $pos);
			$content .= fread($this->fd, $this->block_size);
			if(!($block_id = $this->reader->getNextBlock($block_id)))
				break;
		}

		return $content;
	}
	
	public function seek($offset, $whence = SEEK_SET) {
		//
	}

}
