from datetime import datetime, timedelta

from poly.market_fetcher import MarketFetcher, MarketRecord
from poly.pipeline import MarketPipeline


class StaticClient:
    def __init__(self, payloads):
        self.payloads = payloads

    def fetch_page(self, *, cursor=None, limit=0):
        ids = list(self.payloads.keys())
        if cursor:
            return [], None
        return ([{"id": market_id} for market_id in ids], None)

    def fetch_detail(self, market_id: str):
        return self.payloads[market_id]


def test_pipeline_updates_last_updated_and_step_timestamps():
    clock_calls = []

    def clock():
        now = datetime(2024, 1, 1, 12, 0, 0) + timedelta(seconds=len(clock_calls))
        clock_calls.append(now)
        return now

    client = StaticClient({"mkt-1": {"name": "alpha"}})
    fetcher = MarketFetcher(client, clock=clock)

    def processor(record: MarketRecord):
        record.payload["processed"] = True

    pipeline = MarketPipeline(fetcher, processors=[processor], clock=clock)
    result = pipeline.run()

    record = result.markets["mkt-1"]
    assert record.payload["processed"] is True
    assert "processor" in record.step_timestamps
    assert result.last_updated["mkt-1"] == record.last_updated
    assert "fetch_completed" in result.step_timestamps
    assert "processor_completed" in result.step_timestamps
    assert "pipeline_completed" in result.step_timestamps


def test_pipeline_carries_forward_existing_records_and_marks_run_time():
    clock_calls = []

    def clock():
        now = datetime(2024, 1, 1, 12, 0, 0) + timedelta(seconds=len(clock_calls))
        clock_calls.append(now)
        return now

    existing_record = MarketRecord(
        market_id="mkt-1",
        payload={"name": "alpha"},
        last_updated=datetime(2023, 12, 31, 23, 59, 0),
        step_timestamps={},
    )
    client = StaticClient({"mkt-1": {"name": "alpha"}})
    fetcher = MarketFetcher(client, stale_after=timedelta(hours=2), clock=clock)
    pipeline = MarketPipeline(fetcher, processors=[], clock=clock)

    result = pipeline.run(existing={"mkt-1": existing_record})

    assert result.last_updated["mkt-1"] > existing_record.last_updated
    assert result.step_timestamps["pipeline_started"] <= result.step_timestamps["pipeline_completed"]
