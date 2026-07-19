#!/usr/bin/env python3
"""
Download a small local GGUF model for AutoByte Agent.

Default: Qwen2.5-7B-Instruct Q4_K_M (~4.7GB) — fits a 16GB laptop, CPU-only,
private, no token limit. Override with --repo / --file / --out.

Uses huggingface_hub; xet is disabled (it stalled in the past) so plain LFS is used.
Requires a HF token only for gated repos; Qwen2.5-Instruct is open.
"""
import argparse, pathlib, os, sys

DEFAULT_REPO = "unsloth/Qwen2.5-7B-Instruct-GGUF"
DEFAULT_FILE = "Qwen2.5-7B-Instruct-Q4_K_M.gguf"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", default=DEFAULT_REPO)
    ap.add_argument("--file", default=DEFAULT_FILE)
    ap.add_argument("--out", default=str(pathlib.Path(__file__).resolve().parent / "models"))
    ap.add_argument("--token", default=os.environ.get("HF_TOKEN", ""))
    args = ap.parse_args()

    try:
        from huggingface_hub import hf_hub_download
    except ImportError:
        sys.exit("Missing huggingface_hub. Run: pip install huggingface_hub")

    out = pathlib.Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    os.environ["HF_HUB_DISABLE_XET"] = "1"
    os.environ["HF_XET_DISABLED"] = "1"
    print(f"Downloading {args.repo}/{args.file} -> {out}")
    path = hf_hub_download(
        repo_id=args.repo, filename=args.file,
        local_dir=str(out), token=args.token or None,
    )
    print("Downloaded:", path)
    # write config so agent.py picks it up
    cfg = pathlib.Path(__file__).resolve().parent / "config.json"
    import json
    c = json.loads(cfg.read_text()) if cfg.exists() else {}
    c["mode"] = "local"
    c["model_path"] = str(path)
    cfg.write_text(json.dumps(c, indent=2))
    print("config.json updated -> mode=local, model_path set. Run: python agent.py")


if __name__ == "__main__":
    main()
