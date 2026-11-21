<?php

// SPDX-License-Identifier: MIT OR Apache-2.0

declare(strict_types=1);

namespace OCA\OxiDocsPreview\Preview;

use OCA\OxiDocsPreview\AppInfo\Application;
use OCA\OxiDocsPreview\Engine;
use OCP\Files\File;
use OCP\Files\FileInfo;
use OCP\IAppConfig;
use OCP\IImage;
use OCP\Image;
use OCP\ITempManager;
use OCP\Preview\IProviderV2;
use Psr\Log\LoggerInterface;

/**
 * Thumbnails of Office files: the first page, slide or printed sheet page,
 * drawn by oxidocs. One subclass per kind says which files and which command.
 */
abstract class OfficePreview implements IProviderV2 {
	/** Larger documents are left without a thumbnail (bytes; app config `max_size`). */
	private const DEFAULT_MAX_SIZE = 50 * 1024 * 1024;

	/**
	 * Nextcloud first asks for its largest preview (4096 px by default) and
	 * scales every thumbnail down from that. A page drawn that large takes
	 * seconds to encode and decode, and shows no more of a document than one
	 * at 2048 px (app config `max_pixels`).
	 */
	private const DEFAULT_MAX_PIXELS = 2048;

	public function __construct(
		private Engine $engine,
		private IAppConfig $appConfig,
		private ITempManager $tempManager,
		private LoggerInterface $logger,
	) {
	}

	/** The oxidocs command that draws this kind, such as `docx-to-png`. */
	abstract protected function command(): string;

	/**
	 * The extension a temporary copy gets. oxidocs tells the kinds apart by
	 * it, so a copy of a deck must not be called `.docx`.
	 */
	abstract protected function suffix(): string;

	/** @return list<string> further options for the command */
	protected function options(): array {
		return [];
	}

	public function getMimeType(): string {
		return static::MIME_REGEX;
	}

	public function isAvailable(FileInfo $file): bool {
		$max = $this->appConfig->getValueInt(Application::APP_ID, 'max_size', self::DEFAULT_MAX_SIZE);
		return $file->getSize() > 0 && ($max <= 0 || $file->getSize() <= $max) && $this->engine->binary() !== null;
	}

	public function getThumbnail(File $file, int $maxX, int $maxY): ?IImage {
		$cleanup = [];
		try {
			$input = $this->localPath($file, $cleanup);
			if ($input === null) {
				return null;
			}
			$png = $this->tempManager->getTemporaryFile('.png');
			if ($png === false) {
				return null;
			}
			$cleanup[] = $png;
			$cap = $this->appConfig->getValueInt(Application::APP_ID, 'max_pixels', self::DEFAULT_MAX_PIXELS);
			if ($cap > 0) {
				$maxX = min($maxX, $cap);
				$maxY = min($maxY, $cap);
			}
			if (!$this->engine->drawFirstPage($this->command(), $input, $png, $maxX, $maxY, $this->options())) {
				return null;
			}
			$image = new Image();
			$image->loadFromFile($png);
			if (!$image->valid()) {
				return null;
			}
			$image->scaleDownToFit($maxX, $maxY);
			return $image;
		} catch (\Throwable $e) {
			$this->logger->info('No preview for ' . $file->getPath() . ': ' . $e->getMessage(), [
				'app' => Application::APP_ID,
				'exception' => $e,
			]);
			return null;
		} finally {
			foreach ($cleanup as $path) {
				@unlink($path);
			}
		}
	}

	/**
	 * A path oxidocs can read: the file itself on local, unencrypted storage,
	 * otherwise a temporary copy (added to $cleanup).
	 *
	 * @param list<string> $cleanup
	 */
	private function localPath(File $file, array &$cleanup): ?string {
		$storage = $file->getStorage();
		if (!$file->isEncrypted() && $storage->isLocal()) {
			$path = $storage->getLocalFile($file->getInternalPath());
			if (is_string($path) && is_file($path)) {
				return $path;
			}
		}
		$copy = $this->tempManager->getTemporaryFile($this->suffix());
		if ($copy === false) {
			return null;
		}
		$cleanup[] = $copy;
		$in = $file->fopen('r');
		$out = fopen($copy, 'wb');
		if ($in === false || $out === false) {
			return null;
		}
		try {
			if (stream_copy_to_stream($in, $out) === false) {
				return null;
			}
		} finally {
			fclose($in);
			fclose($out);
		}
		return $copy;
	}
}
