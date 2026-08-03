"""
Performance Analyzer Module

Analyzes persisted conversion metrics to surface actionable insights:
- Throughput trend detection (improving / degrading / stable)
- Bottleneck identification (slow files, slow tools)
- Outlier detection (unusually fast/slow conversions)
- Recommendations based on observed patterns

Phase 4 Week 5: Performance Tuning & Metrics Persistence
"""

import logging
import statistics
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Dict, Optional, Callable, Any

try:
    from metrics_store import MetricsStore, ConversionRecord, create_metrics_store
    METRICS_STORE_AVAILABLE = True
except ImportError:
    METRICS_STORE_AVAILABLE = False

logger = logging.getLogger(__name__)


class TrendDirection(Enum):
    IMPROVING = "improving"
    DEGRADING  = "degrading"
    STABLE     = "stable"
    INSUFFICIENT_DATA = "insufficient_data"


@dataclass
class ThroughputTrend:
    """Result of throughput trend analysis."""
    direction: TrendDirection
    recent_avg_mbps: float       # Average of latest half of data
    baseline_avg_mbps: float     # Average of earlier half of data
    change_percent: float        # Positive = improvement
    sample_count: int

    def summary(self) -> str:
        if self.direction == TrendDirection.INSUFFICIENT_DATA:
            return "Insufficient data for trend analysis (need ≥4 conversions)"
        arrow = {"improving": "↑", "degrading": "↓", "stable": "→"}[self.direction.value]
        sign = "+" if self.change_percent >= 0 else ""
        return (
            f"Throughput trend: {arrow} {self.direction.value.upper()} "
            f"({sign}{self.change_percent:.1f}%) — "
            f"recent {self.recent_avg_mbps:.2f} MB/s vs baseline {self.baseline_avg_mbps:.2f} MB/s"
        )


@dataclass
class Bottleneck:
    """A detected performance bottleneck."""
    category: str          # "slow_tool", "large_file", "low_throughput", "high_failure_rate"
    description: str
    severity: str          # "low", "medium", "high"
    affected_items: List[str] = field(default_factory=list)
    recommendation: str = ""


@dataclass
class OutlierRecord:
    """A conversion identified as an outlier."""
    record: ConversionRecord
    reason: str            # "slow", "fast", "large", "small_output"
    z_score: float         # Standard deviations from mean (absolute value)


@dataclass
class AnalysisReport:
    """Complete analysis report."""
    trend: ThroughputTrend
    bottlenecks: List[Bottleneck]
    outliers: List[OutlierRecord]
    recommendations: List[str]
    stats_snapshot: Dict[str, Any]

    def format_report(self) -> str:
        """Format full analysis as a readable string."""
        lines = ["=" * 60, "📊 ROM Converter Performance Analysis Report", "=" * 60]

        # Stats snapshot
        s = self.stats_snapshot
        if s:
            total = s.get("total", 0)
            successful = s.get("successful", 0) or 0
            rate = f"{(successful / total * 100):.1f}%" if total else "N/A"
            avg_tp = s.get("avg_throughput_mbps") or 0
            lines += [
                "",
                "── Overall Stats ──",
                f"  Total conversions : {total}",
                f"  Success rate      : {rate}",
                f"  Avg throughput    : {avg_tp:.2f} MB/s",
            ]

        # Trend
        lines += ["", "── Throughput Trend ──", f"  {self.trend.summary()}"]

        # Bottlenecks
        if self.bottlenecks:
            lines += ["", "── Bottlenecks ──"]
            for b in self.bottlenecks:
                lines.append(f"  [{b.severity.upper()}] {b.description}")
                if b.recommendation:
                    lines.append(f"    → {b.recommendation}")
        else:
            lines += ["", "── Bottlenecks ──", "  None detected ✅"]

        # Outliers
        if self.outliers:
            lines += ["", "── Outliers ──"]
            for o in self.outliers[:5]:
                lines.append(
                    f"  {o.reason.upper()}: {o.record.file_name} "
                    f"(z={o.z_score:.1f}, {o.record.throughput_mbps:.2f} MB/s)"
                )
            if len(self.outliers) > 5:
                lines.append(f"  ... and {len(self.outliers) - 5} more")

        # Recommendations
        if self.recommendations:
            lines += ["", "── Recommendations ──"]
            for i, r in enumerate(self.recommendations, 1):
                lines.append(f"  {i}. {r}")

        lines.append("=" * 60)
        return "\n".join(lines)


class PerformanceAnalyzer:
    """
    Analyzes metrics from MetricsStore to surface performance insights.
    """

    # Thresholds
    OUTLIER_Z_THRESHOLD = 2.0        # Z-score to flag as outlier
    TREND_CHANGE_THRESHOLD = 10.0    # % change to label improving/degrading
    LOW_THROUGHPUT_MBPS = 5.0        # Below this is flagged as low throughput
    HIGH_FAILURE_RATE = 0.20         # Above 20% failure rate triggers warning

    def __init__(
        self,
        store: Optional["MetricsStore"] = None,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        """
        Initialize the analyzer.

        Args:
            store: MetricsStore to read from. If None, creates an in-memory one.
            log_callback: Optional logging callback
        """
        self.log_callback = log_callback
        if store is not None:
            self.store = store
        elif METRICS_STORE_AVAILABLE:
            self.store = create_metrics_store(db_path=":memory:")
        else:
            self.store = None

    def _log(self, message: str) -> None:
        if self.log_callback:
            self.log_callback(message)
        logger.info(message)

    # ------------------------------------------------------------------
    # Trend analysis
    # ------------------------------------------------------------------

    def analyze_throughput_trend(self, limit: int = 50) -> ThroughputTrend:
        """
        Detect if throughput is improving, degrading, or stable.

        Splits the most recent `limit` successful conversions into two
        halves and compares average throughput.

        Args:
            limit: Number of recent records to consider

        Returns:
            ThroughputTrend result
        """
        if self.store is None:
            return ThroughputTrend(
                direction=TrendDirection.INSUFFICIENT_DATA,
                recent_avg_mbps=0, baseline_avg_mbps=0,
                change_percent=0, sample_count=0,
            )

        data = self.store.get_throughput_trend(limit=limit, successful_only=True)

        if len(data) < 4:
            return ThroughputTrend(
                direction=TrendDirection.INSUFFICIENT_DATA,
                recent_avg_mbps=0, baseline_avg_mbps=0,
                change_percent=0, sample_count=len(data),
            )

        mid = len(data) // 2
        baseline_vals = [d["throughput_mbps"] for d in data[:mid] if d["throughput_mbps"] > 0]
        recent_vals   = [d["throughput_mbps"] for d in data[mid:] if d["throughput_mbps"] > 0]

        if not baseline_vals or not recent_vals:
            return ThroughputTrend(
                direction=TrendDirection.INSUFFICIENT_DATA,
                recent_avg_mbps=0, baseline_avg_mbps=0,
                change_percent=0, sample_count=len(data),
            )

        baseline_avg = statistics.mean(baseline_vals)
        recent_avg   = statistics.mean(recent_vals)
        change_pct   = ((recent_avg - baseline_avg) / baseline_avg * 100) if baseline_avg else 0

        if change_pct > self.TREND_CHANGE_THRESHOLD:
            direction = TrendDirection.IMPROVING
        elif change_pct < -self.TREND_CHANGE_THRESHOLD:
            direction = TrendDirection.DEGRADING
        else:
            direction = TrendDirection.STABLE

        return ThroughputTrend(
            direction=direction,
            recent_avg_mbps=recent_avg,
            baseline_avg_mbps=baseline_avg,
            change_percent=change_pct,
            sample_count=len(data),
        )

    # ------------------------------------------------------------------
    # Bottleneck detection
    # ------------------------------------------------------------------

    def detect_bottlenecks(self) -> List[Bottleneck]:
        """
        Identify performance bottlenecks.

        Checks:
        - Low average throughput
        - High failure rate
        - Tools with significantly lower throughput than average
        - Disproportionately slow file formats

        Returns:
            List of detected Bottleneck objects
        """
        if self.store is None:
            return []

        bottlenecks: List[Bottleneck] = []
        stats = self.store.get_stats()
        tool_stats = self.store.get_tool_stats()

        if not stats or not stats.get("total"):
            return bottlenecks

        total    = stats["total"]
        failed   = stats.get("failed") or 0
        avg_tp   = stats.get("avg_throughput_mbps") or 0

        # 1. Low average throughput
        if avg_tp > 0 and avg_tp < self.LOW_THROUGHPUT_MBPS:
            bottlenecks.append(Bottleneck(
                category="low_throughput",
                description=f"Average throughput is low ({avg_tp:.2f} MB/s)",
                severity="medium",
                recommendation=(
                    "Consider reducing parallel workers to lower I/O contention, "
                    "or move temp files to a faster drive."
                ),
            ))

        # 2. High failure rate
        if total > 0:
            failure_rate = failed / total
            if failure_rate > self.HIGH_FAILURE_RATE:
                bottlenecks.append(Bottleneck(
                    category="high_failure_rate",
                    description=f"High failure rate: {failure_rate*100:.1f}% ({failed}/{total})",
                    severity="high" if failure_rate > 0.4 else "medium",
                    recommendation=(
                        "Check the failure summary for common error patterns. "
                        "Verify tool paths and input file integrity."
                    ),
                ))

        # 3. Slow tools compared to overall average
        if avg_tp > 0 and tool_stats:
            for tool in tool_stats:
                tool_tp = tool.get("avg_throughput_mbps") or 0
                tool_name = tool.get("tool_name") or "(unknown)"
                if tool_tp > 0 and tool_tp < avg_tp * 0.5:
                    bottlenecks.append(Bottleneck(
                        category="slow_tool",
                        description=(
                            f"Tool '{tool_name}' throughput ({tool_tp:.2f} MB/s) is "
                            f">50% below average ({avg_tp:.2f} MB/s)"
                        ),
                        severity="medium",
                        affected_items=[tool_name],
                        recommendation=(
                            f"Check if '{tool_name}' supports multi-threading flags, "
                            "or if an updated version is available."
                        ),
                    ))

        return bottlenecks

    # ------------------------------------------------------------------
    # Outlier detection
    # ------------------------------------------------------------------

    def detect_outliers(self, limit: int = 200) -> List[OutlierRecord]:
        """
        Flag conversions that are statistical outliers.

        Uses Z-score on throughput. Records beyond OUTLIER_Z_THRESHOLD
        standard deviations from the mean are flagged.

        Args:
            limit: Records to analyse

        Returns:
            List of OutlierRecord
        """
        if self.store is None:
            return []

        records = self.store.get_recent(limit=limit)
        successful = [r for r in records if r.success and r.duration_seconds > 0]

        if len(successful) < 4:
            return []

        throughputs = [r.throughput_mbps for r in successful]
        mean = statistics.mean(throughputs)
        try:
            stdev = statistics.stdev(throughputs)
        except statistics.StatisticsError:
            return []

        if stdev == 0:
            return []

        outliers: List[OutlierRecord] = []
        for rec in successful:
            z = abs(rec.throughput_mbps - mean) / stdev
            if z >= self.OUTLIER_Z_THRESHOLD:
                reason = "fast" if rec.throughput_mbps > mean else "slow"
                outliers.append(OutlierRecord(record=rec, reason=reason, z_score=z))

        # Sort by z-score descending (most extreme first)
        outliers.sort(key=lambda o: o.z_score, reverse=True)
        return outliers

    # ------------------------------------------------------------------
    # Recommendations
    # ------------------------------------------------------------------

    def generate_recommendations(
        self,
        trend: ThroughputTrend,
        bottlenecks: List[Bottleneck],
        outliers: List[OutlierRecord],
    ) -> List[str]:
        """
        Generate prioritized recommendations based on analysis.

        Args:
            trend: ThroughputTrend from analyze_throughput_trend
            bottlenecks: From detect_bottlenecks
            outliers: From detect_outliers

        Returns:
            Ordered list of recommendation strings
        """
        recs: List[str] = []

        if trend.direction == TrendDirection.DEGRADING:
            recs.append(
                "Throughput is declining — check for disk fragmentation, "
                "background processes competing for I/O, or tool version regressions."
            )

        high_severity = [b for b in bottlenecks if b.severity == "high"]
        for b in high_severity:
            if b.recommendation:
                recs.append(b.recommendation)

        medium_severity = [b for b in bottlenecks if b.severity == "medium"]
        for b in medium_severity:
            if b.recommendation:
                recs.append(b.recommendation)

        slow_outliers = [o for o in outliers if o.reason == "slow"]
        if len(slow_outliers) > 3:
            names = [o.record.file_name for o in slow_outliers[:3]]
            recs.append(
                f"Several unusually slow conversions detected "
                f"({', '.join(names)}, …). Verify input files are not corrupt."
            )

        if not recs:
            recs.append("No performance issues detected — system is operating normally. ✅")

        return recs

    # ------------------------------------------------------------------
    # Full analysis
    # ------------------------------------------------------------------

    def run_full_analysis(self) -> AnalysisReport:
        """
        Run all analyses and return a combined report.

        Returns:
            AnalysisReport with trend, bottlenecks, outliers, recommendations
        """
        stats = self.store.get_stats() if self.store else {}
        trend = self.analyze_throughput_trend()
        bottlenecks = self.detect_bottlenecks()
        outliers = self.detect_outliers()
        recommendations = self.generate_recommendations(trend, bottlenecks, outliers)

        report = AnalysisReport(
            trend=trend,
            bottlenecks=bottlenecks,
            outliers=outliers,
            recommendations=recommendations,
            stats_snapshot=stats,
        )
        self._log("Performance analysis complete")
        return report


def create_performance_analyzer(
    store: Optional["MetricsStore"] = None,
    log_callback: Optional[Callable[[str], None]] = None,
) -> PerformanceAnalyzer:
    """
    Factory function to create a PerformanceAnalyzer.

    Args:
        store: MetricsStore to read from (creates in-memory one if None)
        log_callback: Optional logging callback

    Returns:
        PerformanceAnalyzer instance
    """
    return PerformanceAnalyzer(store=store, log_callback=log_callback)
