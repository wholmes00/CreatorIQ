#!/usr/bin/env python3
"""
On-Screen Text Extraction Pipeline
====================================
Downloads TikTok videos, extracts key frames, and stores
on-screen text data in Supabase.

Frame analysis is done externally (by AI vision) — this script
handles the download/extraction/storage pipeline.

Usage:
    python extract_onscreen_text.py --batch-size 10 --start-offset 0
    python extract_onscreen_text.py --video-id UUID
    python extract_onscreen_text.py --status  (show progress)
"""

import argparse
import json
import os
import subprocess
import sys
import glob
from pathlib import Path
from supabase import create_client

# --- Config ---
SUPABASE_URL = "https://owklfaoaxdrggmbtcwpn.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im93a2xmYW9heGRyZ2dtYnRjd3BuIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzM0NDQyNjcsImV4cCI6MjA4OTAyMDI2N30.EQkJzeS4MYG4QO6aH9c_zbF7BNuH_bKwZIKQpTXvw1Y"

WORK_DIR = Path("/sessions/fervent-admiring-brown/tiktok_engine/processing")
FRAMES_DIR = WORK_DIR / "frames"
YTDLP = "/sessions/fervent-admiring-brown/.local/bin/yt-dlp"

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)


def get_videos_missing_onscreen_text(limit=None, offset=0):
    """Get videos that don't have on-screen text entries yet."""
    # Get all video IDs that already have on-screen text
    existing = supabase.table("video_onscreen_text").select("video_id").execute().data
    existing_ids = {r["video_id"] for r in existing}

    # Get all videos
    videos = supabase.table("tiktok_videos").select(
        "id, tiktok_url, product_name, content_type, duration_seconds"
    ).order("likes", desc=True).execute().data

    # Filter to missing ones
    missing = [v for v in videos if v["id"] not in existing_ids]

    if offset:
        missing = missing[offset:]
    if limit:
        missing = missing[:limit]

    return missing


def download_video(tiktok_url, output_path):
    """Download a TikTok video using yt-dlp."""
    try:
        # First get available formats
        result = subprocess.run(
            [YTDLP, "--list-formats", tiktok_url],
            capture_output=True, text=True, timeout=30
        )

        # Prefer h264 540p for speed/size balance
        format_id = None
        for line in result.stdout.split('\n'):
            if 'h264_540p' in line:
                format_id = line.split()[0]
                break

        if not format_id:
            # Fallback: any h264 format
            for line in result.stdout.split('\n'):
                if 'h264' in line and 'mp4' in line:
                    format_id = line.split()[0]
                    break

        if not format_id:
            format_id = "best"

        # --force-overwrites: CRITICAL — prevents stale cached files from
        # being served when a different video is downloaded to the same path.
        # Without this, yt-dlp skips download if file exists, which can
        # cause one video's analysis to be applied to a completely different video.
        # --no-cache-dir: Prevents yt-dlp from using cached metadata.
        result = subprocess.run(
            [YTDLP, "-f", format_id, "--no-warnings", "--force-overwrites",
             "--no-cache-dir", "-o", str(output_path), tiktok_url],
            capture_output=True, text=True, timeout=120
        )

        if result.returncode != 0:
            print(f"    yt-dlp error: {result.stderr[:200]}")
            return False

        return output_path.exists()

    except subprocess.TimeoutExpired:
        print(f"    Download timed out")
        return False
    except Exception as e:
        print(f"    Download error: {e}")
        return False


def extract_frames(video_path, frame_dir, interval=3):
    """Extract frames from video at given interval using ffmpeg."""
    frame_dir.mkdir(parents=True, exist_ok=True)

    # CRITICAL: Clean up any stale frames from previous runs before extracting.
    # Without this, leftover frames from a different video could contaminate
    # the analysis if the new video is shorter (fewer frames extracted).
    for old_frame in frame_dir.glob("frame_*.*"):
        old_frame.unlink()

    try:
        result = subprocess.run(
            ["ffmpeg", "-i", str(video_path), "-vf", f"fps=1/{interval}",
             "-q:v", "2", str(frame_dir / "frame_%03d.jpg"), "-y"],
            capture_output=True, text=True, timeout=60
        )

        frames = sorted(frame_dir.glob("frame_*.jpg"))
        return frames

    except Exception as e:
        print(f"    Frame extraction error: {e}")
        return []


def store_onscreen_text(video_id, entries):
    """Store on-screen text entries in Supabase."""
    for entry in entries:
        supabase.table("video_onscreen_text").insert({
            "video_id": video_id,
            "timestamp_seconds": entry["timestamp_seconds"],
            "text_content": entry["text_content"],
            "text_type": entry["text_type"],
            "is_persistent": entry.get("is_persistent", False),
        }).execute()


def process_single_video(video):
    """Download, extract frames, and return frame paths for analysis."""
    video_id = video["id"]
    url = video["tiktok_url"]
    name = video["product_name"]

    print(f"\n  Processing: {name}")
    print(f"  URL: {url}")

    # Setup directories
    video_dir = WORK_DIR / video_id
    video_dir.mkdir(parents=True, exist_ok=True)
    video_path = video_dir / "video.mp4"
    frame_dir = video_dir / "frames"

    # Download
    print(f"    Downloading...")
    if not download_video(url, video_path):
        print(f"    ✗ Download failed")
        return None

    # Extract frames
    print(f"    Extracting frames (every 3s)...")
    frames = extract_frames(video_path, frame_dir, interval=3)
    if not frames:
        print(f"    ✗ No frames extracted")
        return None

    print(f"    ✓ {len(frames)} frames extracted")

    return {
        "video_id": video_id,
        "product_name": name,
        "video_dir": str(video_dir),
        "frame_dir": str(frame_dir),
        "frames": [str(f) for f in frames],
        "frame_count": len(frames),
    }


def show_status():
    """Show current on-screen text extraction progress."""
    existing = supabase.table("video_onscreen_text").select("video_id").execute().data
    existing_ids = len({r["video_id"] for r in existing})
    total = supabase.table("tiktok_videos").select("id", count="exact").execute()
    total_count = total.count if total.count else len(total.data)

    print(f"\n  On-Screen Text Extraction Progress")
    print(f"  {'─' * 40}")
    print(f"  Complete:   {existing_ids}/{total_count}")
    print(f"  Remaining:  {total_count - existing_ids}")
    print(f"  Coverage:   {existing_ids/total_count*100:.1f}%\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="On-screen text extraction pipeline")
    parser.add_argument("--batch-size", type=int, default=10)
    parser.add_argument("--start-offset", type=int, default=0)
    parser.add_argument("--video-id", type=str, default=None)
    parser.add_argument("--status", action="store_true")
    args = parser.parse_args()

    if args.status:
        show_status()
        sys.exit(0)

    if args.video_id:
        videos = supabase.table("tiktok_videos").select(
            "id, tiktok_url, product_name, content_type, duration_seconds"
        ).eq("id", args.video_id).execute().data
    else:
        videos = get_videos_missing_onscreen_text(
            limit=args.batch_size, offset=args.start_offset
        )

    print(f"\n{'═' * 60}")
    print(f"  ON-SCREEN TEXT EXTRACTION PIPELINE")
    print(f"  Videos to process: {len(videos)}")
    print(f"{'═' * 60}")

    results = []
    for i, video in enumerate(videos):
        print(f"\n  [{i+1}/{len(videos)}] ", end="")
        result = process_single_video(video)
        if result:
            results.append(result)

    # Save batch manifest for frame analysis
    manifest_path = WORK_DIR / "batch_manifest.json"
    with open(manifest_path, "w") as f:
        json.dump(results, f, indent=2)

    print(f"\n{'═' * 60}")
    print(f"  BATCH COMPLETE")
    print(f"  Downloaded & extracted: {len(results)}/{len(videos)}")
    print(f"  Manifest saved: {manifest_path}")
    print(f"  Next step: Analyze frames for on-screen text")
    print(f"{'═' * 60}\n")
