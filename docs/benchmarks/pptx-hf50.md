# PPTX blind set - 50 HuggingFace decks

Oxi, LibreOffice and ONLYOFFICE scored against **Microsoft PowerPoint's own render** of
the same files. This is the per-document detail behind the PPTX table in the
[README](../../README.md#powerpoint-pptx-blind-set--50-huggingface-documents).
Raw scores: [`pptx-hf50-results.json`](pptx-hf50-results.json).

## Method

| | |
|---|---|
| Corpus | HuggingFace noxneural/pptx_collection_templates |
| Sampling | random.sample(pool_of_999_pptx, 50) with seed 20260809, frozen before any measurement |
| Ground truth | Microsoft PowerPoint (Microsoft 365 16.0.20131.20154) COM export to PDF |
| Score | SSIM at 150 DPI over same-index slides, LANCZOS resize-to-match, structural_similarity(channel_axis=2, data_range=255) |
| Excluded | 2 of 50 files are corrupt zip containers that PowerPoint COM, LibreOffice, ONLYOFFICE x2t and Oxi all reject; 48 measured |
| Measured | Oxi 2026-08-20 - LibreOffice 2026-08-20 (26.2.1.2) - ONLYOFFICE 2026-08-13 (x2t 9.3.1.8) |

The set was frozen before any engine ran against it, and no document in it has been
opened, anatomized or fixed against. It is re-measured as the engine changes - in both
directions.

## Result

| Engine | mean SSIM vs PowerPoint (per doc) | slide count matches PowerPoint |
|--------|-----------------------------------|--------------------------------|
| **Oxi** | **0.953** | **48 / 48** |
| LibreOffice 26.2.1.2 | 0.913 | 48 / 48 |
| ONLYOFFICE 9.3.1.8 | 0.908 | 48 / 48 |

Paired differences over the 48 documents:

| Comparison | mean difference | std. error | \|t\| | ahead on |
|---|---|---|---|---|
| Oxi - LibreOffice | +0.0398 | 0.0046 | 8.71 | Oxi on 44 / 48 |
| Oxi - ONLYOFFICE | +0.0451 | 0.0078 | 5.79 | Oxi on 46 / 48 |
| LibreOffice - ONLYOFFICE | +0.0053 | 0.0085 | 0.63 | statistically tied |

At its first measurement (2026-08-13, the renderer's first PPTX benchmark) Oxi scored
**0.679** on this set and lost to both suites on 47 of the 48
documents. A week of renderer work - slide-master placeholder inheritance, group
transforms under mirroring, embedded fonts, color emoji, mixed-face line boxes - moved it
to 0.953. **All 48 documents improved and none regressed.**

## Per document

Sorted by Oxi's score, worst first. "Oxi (first)" is the same document at the 2026-08-13
measurement.

| idx | slides | Oxi | LibreOffice | ONLYOFFICE | Oxi (first) | HuggingFace path |
|---|---|---|---|---|---|---|
| 47 | 39 | **0.8801** | 0.8640 | 0.8344 | 0.3802 | Salisbury · SlidesCarnival.pptx |
| 12 | 20 | **0.9117** | 0.8572 | 0.8703 | 0.2358 | Ethics of Artificial Intelligence.pptx |
| 49 | 21 | **0.9147** | 0.8846 | 0.9040 | 0.8000 | Cruise Ship Line Company Profile.pptx |
| 4 | 39 | **0.9249** | 0.9279 | 0.9152 | 0.7597 | Adrian · SlidesCarnival.pptx |
| 28 | 20 | **0.9281** | 0.9226 | 0.8741 | 0.5477 | Dentist Career Day Presentation.pptx |
| 50 | 5 | **0.9288** | 0.9056 | 0.9182 | 0.7758 | June Calendar Planner.pptx |
| 15 | 18 | **0.9302** | 0.8424 | 0.8237 | 0.6314 | Futuristic Engineering Center Modern Presentation(2).pptx |
| 44 | 39 | **0.9375** | 0.9039 | 0.9154 | 0.7882 | Feeble · SlidesCarnival(2).pptx |
| 24 | 18 | **0.9383** | 0.9036 | 0.9153 | 0.8599 | Agribusiness Newsletter Modern Minimal Presentation.pptx |
| 20 | 18 | **0.9393** | 0.9021 | 0.9031 | 0.7559 | Canadian Thanksgiving Day Background Slides.pptx |
| 38 | 20 | **0.9395** | 0.9127 | 0.8991 | 0.6704 | Elegant New Year's Eve MiniTheme Presentation(3).pptx |
| 1 | 20 | **0.9427** | 0.9165 | 0.9172 | 0.7296 | Biology Subject for Elementary_ Tropical Rainforest Wildlife Pink and Green Color Blocks Presentation(1).pptx |
| 36 | 22 | **0.9430** | 0.9517 | 0.9525 | 0.7661 | Colourful Playful Infographic Presentation(1).pptx |
| 37 | 22 | **0.9430** | 0.9517 | 0.9525 | 0.7661 | Colourful Playful Infographic Presentation.pptx |
| 34 | 4 | **0.9448** | 0.8733 | 0.9236 | 0.6962 | Back to School Open to Registration Poster.pptx |
| 31 | 33 | **0.9450** | 0.9346 | 0.9275 | 0.4984 | Green and Beige Retail Startup Pitch Deck.pptx |
| 10 | 18 | **0.9451** | 0.9201 | 0.9133 | 0.6701 | Boho Autumn Leaves Personal Organizer.pptx |
| 13 | 4 | **0.9458** | 0.9358 | 0.9425 | 0.8109 | Family Feud Answer Board Background(1).pptx |
| 42 | 4 | **0.9458** | 0.9358 | 0.9425 | 0.8109 | Family Feud Answer Board Background.pptx |
| 6 | 19 | **0.9461** | 0.8065 | 0.9158 | 0.5754 | Illustrated Pattern April Daily Calendar 2024 Presentation(1).pptx |
| 8 | 19 | **0.9462** | 0.8065 | 0.9159 | 0.5754 | Illustrated Pattern April Daily Calendar 2024 Presentation.pptx |
| 9 | 39 | **0.9488** | 0.9215 | 0.5698 | 0.6780 | Isabella · SlidesCarnival.pptx |
| 3 | 20 | **0.9497** | 0.9202 | 0.9116 | 0.5727 | Fun Illustrative Happy Indian Independence Day!.pptx |
| 25 | 4 | **0.9529** | 0.8672 | 0.9172 | 0.5605 | Christmas Giveaway Poster.pptx |
| 27 | 4 | **0.9569** | 0.9592 | 0.9276 | 0.7749 | Simple Minimal Formal Research Poster(1).pptx |
| 21 | 39 | **0.9580** | 0.9373 | 0.9275 | 0.4125 | Eleanor · SlidesCarnival(2).pptx |
| 33 | 21 | **0.9591** | 0.9095 | 0.9035 | 0.7041 | Modern Abstract Trip Planning Infographic.pptx |
| 2 | 19 | **0.9595** | 0.8959 | 0.8886 | 0.2583 | Self Introduction for High School Students(1).pptx |
| 39 | 12 | **0.9600** | 0.8897 | 0.9001 | 0.7386 | Cupid God of Love Storybook.pptx |
| 7 | 4 | **0.9619** | 0.9011 | 0.9010 | 0.8505 | Real Estate Agent CV Resume.pptx |
| 43 | 20 | **0.9623** | 0.9258 | 0.9180 | 0.6488 | Simple Illustrative Science Lesson for Elementary_ Ozone Layer.pptx |
| 40 | 24 | **0.9631** | 0.9425 | 0.9162 | 0.7057 | Illustrated Landscape Marketing Christmas Background(1).pptx |
| 48 | 15 | **0.9643** | 0.9131 | 0.9051 | 0.6827 | Roadmap Infographics.pptx |
| 35 | 20 | **0.9649** | 0.9352 | 0.9146 | 0.5974 | Retro Y2K SEO Specialist Resume Presentation.ppt(1).pptx |
| 18 | 22 | **0.9655** | 0.8818 | 0.8871 | 0.6898 | Fall October Marketing Calendar(1).pptx |
| 14 | 3 | **0.9676** | 0.9148 | 0.9175 | 0.7980 | Classroom Seating Chart .pptx |
| 41 | 5 | **0.9694** | 0.9617 | 0.9587 | 0.7229 | January Calendar Planner.pptx |
| 45 | 13 | **0.9726** | 0.9152 | 0.9299 | 0.6177 | Volumes of Composite Solids Lesson for High School.pptx |
| 32 | 28 | **0.9731** | 0.9475 | 0.9280 | 0.7239 | Colorful Geometric Company Founder About Me Creative Presentation · SlidesCarnival.pptx |
| 26 | 15 | **0.9736** | 0.9307 | 0.9295 | 0.7322 | Impacted Wisdom Teeth Slides.pptx |
| 11 | 28 | **0.9742** | 0.9307 | 0.9248 | 0.6676 | Colorful Cute Simple Illustrative Doodles Funny About Me Creative Presentation · SlidesCarnival.pptx |
| 29 | 10 | **0.9751** | 0.9203 | 0.9468 | 0.7135 | Cute Scrapbook Baby Milestones Photo Album.pptx |
| 30 | 20 | **0.9767** | 0.9263 | 0.9189 | 0.8521 | Volunteer Onboarding Presentation.pptx |
| 19 | 20 | **0.9780** | 0.9281 | 0.9262 | 0.7230 | Pre-K Outdoor Activities to Celebrate National Play Outside Day Presentation.pptx |
| 23 | 26 | **0.9784** | 0.9386 | 0.9314 | 0.7912 | Red, Blue and Yellow Cute Doodle Science Subject for Elementary School Magnetism Presentation.pptx |
| 46 | 31 | **0.9788** | 0.9501 | 0.9390 | 0.8574 | Soft Green and Beige Simple Doodles Product Timeline Presentation.pptx |
| 17 | 28 | **0.9795** | 0.9576 | 0.9249 | 0.7888 | Blue Red and Yellow Cute Simple Illustrative Doodles Pretty Social Media Creative Presentation · SlidesCarnival.pptx |
| 5 | 28 | **0.9831** | 0.9375 | 0.9218 | 0.6160 | Black Purple and Green Geometric Corporate Resume Creative Presentation · SlidesCarnival.pptx |

## Reproduce

1. Fetch the 50 decks from the public HuggingFace dataset with the sampling rule above
   (`random.sample` over the 999 `.pptx` in the pool, seed `20260809`).
2. Ground truth: export each deck to PDF with PowerPoint (Windows + Microsoft 365) and
   rasterise at 150 DPI.
3. Candidates: `tools/oxi-pptx-renderer` (PNG at 150 DPI), LibreOffice
   `soffice --headless --convert-to pdf`, ONLYOFFICE `x2t`.
4. Score each same-index slide with scikit-image
   `structural_similarity(channel_axis=2, data_range=255)`, LANCZOS-resizing the candidate
   to the reference when raster sizes differ; a document's score is the mean over its
   slides.

`tools/metrics/pptx_ssim_floor.py` runs exactly this scoring loop against a local
PowerPoint PDF set — it is the development-corpus tool, with the same conventions as the
blind-set harness.
