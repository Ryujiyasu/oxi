<?php

// SPDX-License-Identifier: MIT OR Apache-2.0

declare(strict_types=1);

namespace OCA\OxiDocsPreview\Preview;

/** Excel workbooks: the first printed page, as Excel would print it. */
class XlsxPreview extends OfficePreview {
	/** .xlsx, .xltx and their macro-enabled forms. */
	public const MIME_REGEX = '/^application\/(vnd\.openxmlformats-officedocument\.spreadsheetml\.(sheet|template)|vnd\.ms-excel\.(sheet|template)\.macroenabled\.12)$/i';

	protected function command(): string {
		return 'xlsx-to-png';
	}

	protected function suffix(): string {
		return '.xlsx';
	}

	/**
	 * A printed sheet is mostly paper; draw only the part with ink, so the
	 * table fills the thumbnail instead of sitting in a corner of an A4 page.
	 */
	protected function options(): array {
		return ['--trim'];
	}
}
