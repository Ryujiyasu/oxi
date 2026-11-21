<?php

// SPDX-License-Identifier: MIT OR Apache-2.0

declare(strict_types=1);

namespace OCA\OxiDocsPreview\Preview;

/** PowerPoint decks: slide 1. */
class PptxPreview extends OfficePreview {
	/** .pptx, .potx, .ppsx and their macro-enabled forms. */
	public const MIME_REGEX = '/^application\/(vnd\.openxmlformats-officedocument\.presentationml\.(presentation|template|slideshow)|vnd\.ms-powerpoint\.(presentation|template|slideshow)\.macroenabled\.12)$/i';

	protected function command(): string {
		return 'pptx-to-png';
	}

	protected function suffix(): string {
		return '.pptx';
	}
}
