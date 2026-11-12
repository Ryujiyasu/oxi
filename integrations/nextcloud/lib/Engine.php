<?php

// SPDX-License-Identifier: MIT OR Apache-2.0

declare(strict_types=1);

namespace OCA\OxiDocsPreview;

use OCA\OxiDocsPreview\AppInfo\Application;
use OCP\App\IAppManager;
use OCP\IAppConfig;
use OCP\IBinaryFinder;
use OCP\ITempManager;
use Psr\Log\LoggerInterface;

/**
 * The oxidocs program: where it is, and drawing a page with it.
 *
 * The app ships static Linux builds in bin/. They are run from a copy under
 * Nextcloud's temporary directory, never in place: an app unpacked by the
 * App Store may have lost its execute bits, and oxidocs writes its bundled
 * fonts next to the binary on first use, which inside a signed app would
 * fail the integrity check.
 *
 * An administrator can point the app at another build instead:
 *   occ config:app:set oxidocs_preview binary --value=/usr/local/bin/oxidocs
 */
class Engine {
	/** Seconds a single thumbnail may take before it is abandoned. */
	private const DEFAULT_TIMEOUT = 30;

	private ?string $resolved = null;
	private bool $looked = false;

	public function __construct(
		private IAppConfig $appConfig,
		private IAppManager $appManager,
		private IBinaryFinder $binaryFinder,
		private ITempManager $tempManager,
		private LoggerInterface $logger,
	) {
	}

	/** The program to run, or null when this server has none usable. */
	public function binary(): ?string {
		if (!$this->looked) {
			$this->looked = true;
			$this->resolved = $this->configured() ?? $this->bundled() ?? $this->onPath();
			if ($this->resolved === null) {
				$this->logger->warning(
					'No usable oxidocs program for ' . php_uname('m') . ': Word previews are off. '
					. 'Install one and run: occ config:app:set oxidocs_preview binary --value=/path/to/oxidocs',
					['app' => Application::APP_ID],
				);
			}
		}
		return $this->resolved;
	}

	/**
	 * Draw page 1 of $docx into $png, fitted inside $maxX × $maxY pixels.
	 * False (and a log line) when the program fails or runs out of time.
	 */
	public function drawFirstPage(string $docx, string $png, int $maxX, int $maxY): bool {
		$binary = $this->binary();
		if ($binary === null) {
			return false;
		}
		$command = [$binary, 'docx-to-png', '--width', (string)max(1, $maxX), '--height', (string)max(1, $maxY), $docx, $png];
		[$status, $stderr] = $this->run($command, $this->timeout());
		if ($status === 0 && is_file($png) && filesize($png) > 0) {
			return true;
		}
		$this->logger->info(
			'oxidocs could not draw a preview (status ' . var_export($status, true) . '): ' . $this->tail($stderr),
			['app' => Application::APP_ID],
		);
		return false;
	}

	private function timeout(): int {
		$value = $this->appConfig->getValueInt(Application::APP_ID, 'timeout', self::DEFAULT_TIMEOUT);
		return $value > 0 ? $value : self::DEFAULT_TIMEOUT;
	}

	private function configured(): ?string {
		$path = $this->appConfig->getValueString(Application::APP_ID, 'binary', '');
		if ($path === '') {
			return null;
		}
		if ($this->works($path)) {
			return $path;
		}
		$this->logger->warning('The configured oxidocs program ' . $path . ' does not run', ['app' => Application::APP_ID]);
		return null;
	}

	private function onPath(): ?string {
		$path = $this->binaryFinder->findBinaryPath('oxidocs');
		return is_string($path) && $this->works($path) ? $path : null;
	}

	/** The build in bin/ for this machine, copied out to where it may run. */
	private function bundled(): ?string {
		$arch = match (strtolower(php_uname('m'))) {
			'x86_64', 'amd64' => 'x86_64',
			'aarch64', 'arm64' => 'aarch64',
			default => null,
		};
		if ($arch === null) {
			return null;
		}
		$source = $this->appManager->getAppPath(Application::APP_ID) . '/bin/oxidocs-linux-' . $arch;
		if (!is_file($source)) {
			return null;
		}
		$version = $this->appManager->getAppVersion(Application::APP_ID);
		$dir = rtrim($this->tempManager->getTempBaseDir(), '/') . '/oxidocs_preview-' . $version . '-' . $arch;
		$target = $dir . '/oxidocs';
		if (!is_file($target) || filesize($target) !== filesize($source)) {
			if (!is_dir($dir) && !@mkdir($dir, 0700, true) && !is_dir($dir)) {
				return null;
			}
			// Copy under a private name, then rename: two requests may race here.
			$partial = $target . '.' . bin2hex(random_bytes(4));
			if (!@copy($source, $partial) || !@chmod($partial, 0755) || !@rename($partial, $target)) {
				@unlink($partial);
				$this->logger->warning('Could not copy oxidocs to ' . $dir, ['app' => Application::APP_ID]);
				return null;
			}
		}
		if ($this->works($target)) {
			return $target;
		}
		$this->logger->warning(
			'The bundled oxidocs does not run from ' . $dir . ' (a noexec temporary directory?)',
			['app' => Application::APP_ID],
		);
		return null;
	}

	/** Whether $path runs at all: without arguments oxidocs prints its usage and exits 1. */
	private function works(string $path): bool {
		if (!is_file($path) || !is_executable($path)) {
			return false;
		}
		[$status, $stderr] = $this->run([$path], 10);
		return $status === 1 && str_contains($stderr, 'oxi docx-to-png');
	}

	/**
	 * Run $command without a shell. Returns [exit status or null on timeout or
	 * failure to start, stderr]. Output goes to files, not pipes: oxidocs
	 * reports freely on stderr, and an undrained pipe would stall it.
	 *
	 * @param list<string> $command
	 * @return array{0: ?int, 1: string}
	 */
	private function run(array $command, int $timeout): array {
		$err = $this->tempManager->getTemporaryFile('.log');
		if ($err === false) {
			return [null, 'no temporary file'];
		}
		$process = @proc_open($command, [
			0 => ['file', '/dev/null', 'r'],
			1 => ['file', '/dev/null', 'w'],
			2 => ['file', $err, 'w'],
		], $pipes);
		if (!is_resource($process)) {
			@unlink($err);
			return [null, 'could not start ' . $command[0]];
		}
		$deadline = microtime(true) + $timeout;
		$status = null;
		while (true) {
			$state = proc_get_status($process);
			if (!$state['running']) {
				$status = $state['exitcode'];
				break;
			}
			if (microtime(true) > $deadline) {
				proc_terminate($process, 9);
				break;
			}
			usleep(20000);
		}
		proc_close($process);
		$stderr = (string)@file_get_contents($err);
		@unlink($err);
		if ($status === null) {
			$stderr .= "\n(timed out after {$timeout} s)";
		}
		return [$status, $stderr];
	}

	private function tail(string $text): string {
		$text = trim($text);
		return strlen($text) > 500 ? '…' . substr($text, -500) : $text;
	}
}
