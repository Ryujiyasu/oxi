// How many people took Oxi, and how many are still running it — without the app
// sending anything.
//
// GitHub counts every download of every release asset, and two of those counts
// answer different questions:
//
//   the installers   how many people took it, per platform
//   latest.json      how many times an installed copy asked whether it was out
//                    of date, which it does once at every launch
//
// The second is the closer thing to "how many are using it", and it is free:
// the updater already fetches that file, so the count exists whether or not
// anyone looks at it. Read it, do not trust it too far — see the notes at the
// bottom of the output.
//
//   node tools/release-downloads.mjs [owner/repo]
//
// Needs the gh CLI, logged in.
import { execFileSync } from 'node:child_process';

const repo = process.argv[2] || 'Ryujiyasu/oxi';

const releases = JSON.parse(
  execFileSync('gh', [
    'api', `repos/${repo}/releases`, '--paginate',
    '--jq', '[.[] | {tag: .tag_name, draft: .draft, published: .published_at, assets: [.assets[] | {name: .name, downloads: .download_count}]}]',
  ], { encoding: 'utf8' }),
);

/// Which platform an asset's name says it is for. The signatures beside each
/// installer are the updater's, not something a person downloads.
function platformOf(name) {
  if (/\.sig$/.test(name)) return null;
  if (name === 'latest.json') return null;
  if (/\.(exe|msi)$/i.test(name)) return 'Windows';
  if (/\.(dmg|app\.tar\.gz)$/i.test(name)) return 'macOS';
  if (/\.(AppImage|deb|rpm)$/i.test(name)) return 'Linux';
  return 'other';
}

let anyPublished = false;
for (const release of releases) {
  const checks = release.assets.find((a) => a.name === 'latest.json');
  const byPlatform = new Map();
  for (const asset of release.assets) {
    const where = platformOf(asset.name);
    if (!where) continue;
    byPlatform.set(where, (byPlatform.get(where) || 0) + asset.downloads);
  }
  const took = [...byPlatform.values()].reduce((sum, n) => sum + n, 0);
  if (!release.draft) anyPublished = true;

  console.log(
    `${release.tag}${release.draft ? '  (draft — nothing counts yet)' : ''}`
    + `${release.published ? `  ${release.published.slice(0, 10)}` : ''}`,
  );
  if (byPlatform.size === 0) {
    console.log('    no installers attached');
  } else {
    for (const [where, n] of [...byPlatform].sort((a, b) => b[1] - a[1])) {
      console.log(`    ${String(n).padStart(7)}  downloads  ${where}`);
    }
    console.log(`    ${String(took).padStart(7)}  downloads  all platforms`);
  }
  if (checks) {
    console.log(`    ${String(checks.downloads).padStart(7)}  update checks (latest.json)`);
  }
  console.log();
}

console.log('What these numbers are not:');
console.log('  · An update check is one launch, not one machine. Someone who opens Oxi');
console.log('    every morning is thirty checks a month, not thirty people.');
console.log('  · A download is not an install. Some are never run; some are run on');
console.log('    several machines.');
console.log('  · A copy taken from anywhere but this release page — a mirror, a');
console.log('    software directory carrying its own copy — is invisible here.');
if (!anyPublished) {
  console.log('  · Nothing is published yet, so every count above is zero by definition.');
}
