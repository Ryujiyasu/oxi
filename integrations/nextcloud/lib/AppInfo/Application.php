<?php

// SPDX-License-Identifier: MIT OR Apache-2.0

declare(strict_types=1);

namespace OCA\OxiDocsPreview\AppInfo;

use OCA\OxiDocsPreview\Preview\DocxPreview;
use OCA\OxiDocsPreview\Preview\PptxPreview;
use OCA\OxiDocsPreview\Preview\XlsxPreview;
use OCP\AppFramework\App;
use OCP\AppFramework\Bootstrap\IBootContext;
use OCP\AppFramework\Bootstrap\IBootstrap;
use OCP\AppFramework\Bootstrap\IRegistrationContext;

class Application extends App implements IBootstrap {
	public const APP_ID = 'oxidocs_preview';

	public function __construct() {
		parent::__construct(self::APP_ID);
	}

	public function register(IRegistrationContext $context): void {
		// An app's provider is not subject to `enabledPreviewProviders` (that
		// list gates the core ones), and the longer regex is tried before
		// core's office pattern, so these files come here once enabled.
		foreach ([DocxPreview::class, PptxPreview::class, XlsxPreview::class] as $provider) {
			$context->registerPreviewProvider($provider, $provider::MIME_REGEX);
		}
	}

	public function boot(IBootContext $context): void {
	}
}
