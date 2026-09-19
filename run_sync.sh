#!/usr/bin/env bash
# ==============================================================================
# Script Jalan Cepat Sinkronisasi & Generator Pas Foto 3x4
#
# Penggunaan:
#   ./run_sync.sh            -> Menjalankan sinkronisasi sekali jalan (on-demand)
#   ./run_sync.sh --watch    -> Menjalankan pemantau otomatis berkala (background)
#   ./run_sync.sh --force    -> Memaksa regenerasi dokumen PDF & DOCX
# ==============================================================================

cd "$(dirname "$0")" || exit 1

if [ "$1" == "--force" ]; then
    python3 sync_and_generate.py --force-regen --upload-drive
elif [ "$1" == "--watch" ]; then
    shift
    python3 sync_and_generate.py --watch --interval 5 --upload-drive "$@"
else
    python3 sync_and_generate.py --upload-drive "$@"
fi

