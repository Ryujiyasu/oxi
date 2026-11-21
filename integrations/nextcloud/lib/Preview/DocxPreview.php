<?php

// SPDX-License-Identifier: MIT OR Apache-2.0

declare(strict_types=1);

namespace OCA\OxiDocsPreview\Preview;

/** Word documents: page 1. */
class DocxPreview extends OfficePreview {
	/** .docx, .dotx and their macro-enabled forms. */
	public const MIME_REGEX = '/^application\/(vnd\.openxmlformats-officedocument\.wordprocessingml\.(document|template)|vnd\.ms-word\.(document|template)\.macroenabled\.12)$/i';

	protected function command(): string {
		return 'docx-to-png';
	}

	protected function suffix(): string {
		return '.docx';
	}
}
