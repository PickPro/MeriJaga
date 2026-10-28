#!/usr/bin/env python3
"""
run_credit_risk_2ppts.py
========================
Automated 2-PPT pipeline for Credit Risk Metrics Slide 2 generation.

Takes 2 presentations:
  1. Old Template PPTX (--template): The deck containing the proven Slide 2 structure.
  2. New Deck PPTX (--target / --source): The new presentation with an empty Slide 2.
  3. Audit JSON (--audit): The validated metrics report (.audit.json).

Workflow:
  Step 1: Uses PowerPoint COM to natively copy Slide 2 from the Template into the New Deck,
          cleanly replacing the empty Slide 2 placeholder while leaving all other slides untouched.
  Step 2: Updates Title, Commentary, Top-Right Summary Counts, and Detailed 9x7 Table with the
          latest quarter metrics, executive category dividing lines, and bold colors.
  Step 3: Automatically exports a high-resolution preview PNG for immediate visual verification.

Usage:
  python run_credit_risk_2ppts.py --template "old_template.pptx" --target "new_deck.pptx" --audit "audit.json" --output "completed_deck.pptx"
"""

import sys
import os
import argparse
import subprocess
from pathlib import Path
from patch_original_ppt import patch_pptx


def find_default_file(patterns, search_dirs):
    for d in search_dirs:
        if not os.path.exists(d):
            continue
        for root, _, files in os.walk(d):
            for f in files:
                for pat in patterns:
                    if pat.lower() in f.lower():
                        return os.path.abspath(os.path.join(root, f))
    return None


def export_slide_preview(pptx_path, output_png_path, slide_num=2):
    try:
        import win32com.client
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        pres = None
        try:
            pres = ppt_app.Presentations.Open(os.path.abspath(str(pptx_path)), WithWindow=False)
            if pres.Slides.Count >= slide_num:
                s = pres.Slides(slide_num)
                s.Export(os.path.abspath(str(output_png_path)), "PNG", 1920, 1080)
                return True
        finally:
            if pres is not None:
                try: pres.Close()
                except Exception: pass
            try:
                if ppt_app.Presentations.Count == 0:
                    ppt_app.Quit()
            except Exception: pass
    except Exception as e:
        print(f"Note: Could not export preview PNG ({e})", file=sys.stderr)
    return False


def main():
    parser = argparse.ArgumentParser(
        description="2-PPT Pipeline: Copy Slide 2 from Template into New Deck and Patch with Audit JSON"
    )
    parser.add_argument(
        "-t", "--template", default=None,
        help="Path to old template PPTX (containing proven Slide 2 structure)"
    )
    parser.add_argument(
        "-s", "--target", "--source", default=None, dest="target",
        help="Path to new PPTX deck (with empty Slide 2 placeholder, e.g. Final Deck.pptx)"
    )
    parser.add_argument(
        "-a", "--audit", default=None,
        help="Path to audit JSON file (e.g. final_3q26_verification.audit.json)"
    )
    parser.add_argument(
        "-o", "--output", default=None,
        help="Path for output presentation (e.g. Final_Deck_3Q26_Completed.pptx)"
    )
    parser.add_argument(
        "--preview", action="store_true", default=True,
        help="Export high-resolution Slide 2 preview PNG (default: True)"
    )

    args = parser.parse_args()

    search_dirs = [os.getcwd(), "c:/chatgpt_recovery", "c:/Users/newzs/Downloads", "c:/Users/newzs/Documents"]

    # Resolve Template
    template_path = args.template
    if not template_path:
        template_path = find_default_file(["crm report", "test_patched_final.pptx", "ideal_slide"], search_dirs)
    if not template_path or not os.path.isfile(template_path):
        print("ERROR: Template PPTX not found. Please provide with --template <file.pptx>", file=sys.stderr)
        sys.exit(1)

    # Resolve Target (New Deck)
    target_path = args.target
    if not target_path:
        target_path = find_default_file(["final deck", "test_empty_slide_deck.pptx"], search_dirs)
    if not target_path or not os.path.isfile(target_path):
        print("ERROR: Target PPTX not found. Please provide with --target <file.pptx>", file=sys.stderr)
        sys.exit(1)

    # Resolve Audit JSON
    audit_path = args.audit
    if not audit_path:
        audit_path = find_default_file(["final_3q26_verification.audit.json", ".audit.json"], search_dirs)
    if not audit_path or not os.path.isfile(audit_path):
        print("ERROR: Audit JSON not found. Please provide with --audit <file.audit.json>", file=sys.stderr)
        sys.exit(1)

    # Resolve Output Path
    output_path = args.output
    if not output_path:
        base, ext = os.path.splitext(target_path)
        output_path = f"{base}_3Q26_Updated{ext}"

    print("=" * 65)
    print("   3Q26 Credit Risk Metrics - 2-PPT Automated Slide 2 Pipeline")
    print("=" * 65)
    print(f"1. Template Deck (Source of Slide 2 Layout) : {template_path}")
    print(f"2. Target Deck   (Deck with Empty Slide 2)  : {target_path}")
    print(f"3. Audit Report  (Latest Quarterly Metrics) : {audit_path}")
    print(f"4. Output Deck   (Final Completed Deck)     : {output_path}")
    print("-" * 65)

    print("[Step 1/3] Copying Slide 2 natively from Template into New Deck...")
    # patch_pptx handles template copying natively when template_pptx is passed
    res = patch_pptx(
        source=target_path,
        output=output_path,
        audit=audit_path,
        same_slide=True,
        template_pptx=template_path
    )

    print("[Step 2/3] Patching latest metrics, summary counts, and colors...")
    layout = res.get("layout_info", {})
    print(f"           - Summary table updated: {layout.get('summary_table_updated')}")
    print(f"           - Detailed table height: {layout.get('table_height_in')} in")
    print(f"           - Table bottom position: {layout.get('table_bottom_in')} in (clean footer buffer)")

    preview_path = os.path.splitext(output_path)[0] + "_Slide2_Preview.png"
    if args.preview:
        print("[Step 3/3] Exporting Slide 2 high-res preview...")
        ok = export_slide_preview(output_path, preview_path, slide_num=2)
        if ok:
            print(f"           - Preview saved: {preview_path}")

    print("=" * 65)
    print("SUCCESS! Slide 2 completed and verified.")
    print(f"Output File: {output_path}")
    if os.path.exists(preview_path):
        print(f"Preview PNG: {preview_path}")
    print("=" * 65)


if __name__ == "__main__":
    main()
