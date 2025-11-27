from __future__ import annotations

import argparse
from datetime import timedelta

from .market_fetcher import MarketClient, MarketFetcher
from .pipeline import MarketPipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run market refresh pipeline.")
    parser.add_argument("--force-refresh", action="store_true", help="Always refetch markets regardless of freshness.")
    parser.add_argument(
        "--stale-minutes",
        type=int,
        default=60,
        help="Minutes after which cached markets are considered stale.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=200,
        help="Page size when listing markets (raises default from previous runs).",
    )
    return parser


def run_pipeline(
    client: MarketClient,
    *,
    processors: list | None = None,
    args: argparse.Namespace | None = None,
) -> None:
    """Entrypoint to execute the market pipeline with CLI-friendly overrides."""

    parsed_args = args or build_parser().parse_args()
    fetcher = MarketFetcher(
        client=client,
        default_limit=parsed_args.limit,
        stale_after=timedelta(minutes=parsed_args.stale_minutes),
    )
    pipeline = MarketPipeline(fetcher=fetcher, processors=processors)
    pipeline.run(force_refresh=parsed_args.force_refresh)


if __name__ == "__main__":  # pragma: no cover
    parser = build_parser()
    ns = parser.parse_args()
    raise SystemExit("Provide a concrete client implementation before running the pipeline.")
