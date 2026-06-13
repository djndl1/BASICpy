#!/usr/bin/env bash
set -euo pipefail

if [ "$#" -ne 1 ]; then
    echo "usage: $0 <project-root>" >&2
    exit 1
fi

root="$1"
if [ ! -d "$root" ]; then
    echo "error: '$root' is not a directory" >&2
    exit 1
fi

find "$root" -type f \( -name '*.cls' -o -name '*.bas' \) -exec unix2dos {} \;
