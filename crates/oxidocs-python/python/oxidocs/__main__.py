# SPDX-License-Identifier: MIT OR Apache-2.0
"""`python -m oxidocs ...` runs the bundled oxidocs command line."""

import subprocess
import sys

from oxidocs import binary_path

if __name__ == "__main__":
    sys.exit(subprocess.call([binary_path(), *sys.argv[1:]]))
