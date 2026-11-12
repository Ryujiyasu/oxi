<?php

// SPDX-License-Identifier: MIT OR Apache-2.0

declare(strict_types=1);

namespace OCA\OxiDocsPreview\AppInfo;

use OCA\OxiDocsPreview\Preview\DocxPreview;
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
		// core's office pattern, so Word documents come here once enabled.
		$context->registerPreviewProvider(DocxPreview::class, DocxPreview::MIME_REGEX);
	}

	public function boot(IBootContext $context): void {
	}
}
