#!/usr/bin/env python3
"""Download a local GGUF model for AutoByte Agent and verify its checksum.

Default: Qwen2.5-7B-Instruct Q4_K_M (~4.7 GB) - fits a 16 GB laptop, CPU-only.
The expected SHA-256 comes from --sha256 if given, otherwise from the Hugging Face
LFS metadata for the pinned --revision. The download is rejected on mismatch.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import pathlib
import sys

HERE = pathlib.Path(__file__).resolve().parent
DEFAULT_REPO = "unsloth/Qwen2.5-7B-Instruct-GGUF"
DEFAULT_FILE = "Qwen2.5-7B-Instruct-Q4_K_M.gguf"
DEFAULT_REVISION = "main"  # pin to a commit SHA for reproducible fleet installs


def sha256_file(path: pathlib.Path, chunk: int = 8 << 20) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        while block := fh.read(chunk):
            h.update(block)
    return h.hexdigest()


def expected_sha256(repo: str, filename: str, revision: str, token: str | None) -> str:
    from huggingface_hub import HfApi

    info = HfApi().get_paths_info(repo, [filename], expand=True, revision=revision, token=token)
    lfs = getattr(info[0], "lfs", None) if info else None
    sha = getattr(lfs, "sha256", None)
    if not sha:
        raise SystemExit(f"No LFS sha256 published for {repo}/{filename}; pass --sha256.")
    return str(sha).lower()


def write_config(model_path: pathlib.Path) -> pathlib.Path:
    cfg_path = HERE / "config.json"
    base = HERE / "config.example.json"
    src = cfg_path if cfg_path.exists() else base
    cfg = json.loads(src.read_text(encoding="utf-8")) if src.exists() else {}
    cfg["mode"] = "local"
    cfg["model_path"] = str(model_path)
    cfg.pop("api_key", None)  # keys belong in OPENAI_API_KEY, never on disk
    cfg_path.write_text(json.dumps(cfg, indent=2) + "\n", encoding="utf-8")
    return cfg_path


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--repo", default=DEFAULT_REPO)
    ap.add_argument("--file", default=DEFAULT_FILE)
    ap.add_argument("--revision", default=DEFAULT_REVISION)
    ap.add_argument("--sha256", default="", help="expected SHA-256 (overrides HF metadata)")
    ap.add_argument("--out", default=str(HERE / "models"))
    args = ap.parse_args(argv)
    token = os.environ.get("HF_TOKEN") or None

    try:
        from huggingface_hub import hf_hub_download
    except ImportError:
        sys.exit("Missing huggingface_hub. Run: pip install -r requirements-setup.txt")

    out = pathlib.Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    os.environ.setdefault("HF_HUB_DISABLE_XET", "1")  # xet stalled in the past; use plain LFS

    want = args.sha256.lower() or expected_sha256(args.repo, args.file, args.revision, token)
    print(f"Downloading {args.repo}/{args.file}@{args.revision} -> {out}")
    path = pathlib.Path(
        hf_hub_download(
            repo_id=args.repo, filename=args.file, revision=args.revision, local_dir=str(out), token=token
        )
    )
    print("Verifying SHA-256 ...")
    got = sha256_file(path)
    if got != want:
        path.unlink(missing_ok=True)
        sys.exit(f"Checksum mismatch: expected {want}, got {got}. File removed.")
    print(f"Verified {got}")
    print(f"{write_config(path)} updated -> mode=local. Run: python agent.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
