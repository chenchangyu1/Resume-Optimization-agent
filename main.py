import argparse
from pathlib import Path

from agent_service import optimize_resume_docx, save_optimized_resume_docx, save_word_units_snapshot


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Use LangChain + GLM to optimize a DOCX resume for a target job."
    )
    parser.add_argument("--resume", required=True, help="Path to resume DOCX or DOC")
    parser.add_argument(
        "--job",
        required=False,
        help="Path to job requirement file (.txt/.md) or image (.png/.jpg/.jpeg/.webp/.bmp)",
    )
    parser.add_argument(
        "--job-text",
        required=False,
        help="Direct job requirement text content",
    )
    parser.add_argument("--out", default="output/optimized_resume.docx", help="Output DOCX path")
    args = parser.parse_args()

    if not args.job and not args.job_text:
        raise ValueError("You must provide either --job (file path) or --job-text (plain text)")

    if args.job and args.job_text:
        raise ValueError("Please provide only one: --job or --job-text")

    optimized = optimize_resume_docx(
        resume_path=args.resume,
        job_path=args.job,
        job_text=args.job_text,
    )

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    save_optimized_resume_docx(
        optimized_units=optimized["optimized_units"],
        source_resume_path=optimized.get("resolved_resume_path", args.resume),
        output_docx_path=str(out_path),
    )

    units_path = out_path.with_suffix(".units.json")
    save_word_units_snapshot(optimized["word_units"], str(units_path))

    print(f"Optimized resume saved to: {out_path}")


if __name__ == "__main__":
    main()
