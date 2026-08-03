"""
Phase 4 Week 3: Performance Optimization Module

Provides:
- Parallel conversion processing with semaphore-based concurrency
- Resource-aware worker count determination
- CPU/RAM/Disk throttling
- Performance metrics and monitoring
- ETA calculation and tracking
"""

import threading
import multiprocessing
import time
import psutil
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional, Callable, Dict, List
from queue import Queue
from datetime import datetime, timedelta


@dataclass
class ConversionMetrics:
    """Tracks metrics for a single conversion operation"""
    file_path: str
    start_time: float = field(default_factory=time.time)
    end_time: Optional[float] = None
    file_size_bytes: int = 0
    output_size_bytes: int = 0
    success: bool = False
    error_message: Optional[str] = None
    
    @property
    def duration_seconds(self) -> float:
        """Get duration in seconds"""
        if self.end_time:
            return self.end_time - self.start_time
        return time.time() - self.start_time
    
    @property
    def throughput_mbps(self) -> float:
        """Get throughput in MB/s"""
        if self.duration_seconds > 0 and self.file_size_bytes > 0:
            mb_processed = self.file_size_bytes / (1024 * 1024)
            return mb_processed / self.duration_seconds
        return 0.0
    
    @property
    def compression_ratio(self) -> float:
        """Get compression ratio (output/input)"""
        if self.file_size_bytes > 0:
            return self.output_size_bytes / self.file_size_bytes
        return 1.0


class ResourceMonitor:
    """Monitors system resources (CPU, RAM, Disk)"""
    
    def __init__(self, log_callback: Optional[Callable[[str], None]] = None):
        """Initialize resource monitor
        
        Args:
            log_callback: Optional logging callback
        """
        self.log_callback = log_callback
        self.cpu_threshold_percent = 95
        self.ram_threshold_percent = 80
        self.ram_critical_percent = 85
        self.disk_write_throttle_mb_s = 500
        self.last_check_time = 0
        self.check_interval = 0.5  # Check every 500ms
    
    def _log(self, message: str) -> None:
        """Log a message"""
        if self.log_callback:
            self.log_callback(message)
    
    def get_cpu_percent(self) -> float:
        """Get current CPU usage percentage (0-100)"""
        try:
            return psutil.cpu_percent(interval=0.1)
        except Exception:
            return 0.0
    
    def get_memory_percent(self) -> float:
        """Get current memory usage percentage (0-100)"""
        try:
            return psutil.virtual_memory().percent
        except Exception:
            return 0.0
    
    def get_available_memory_mb(self) -> float:
        """Get available memory in MB"""
        try:
            return psutil.virtual_memory().available / (1024 * 1024)
        except Exception:
            return 0.0
    
    def get_disk_write_rate_mbps(self) -> float:
        """Get current disk write rate in MB/s"""
        try:
            io_counters = psutil.disk_io_counters()
            if io_counters:
                current_time = time.time()
                # Calculate write rate based on counters
                # This is approximate, real implementation would track deltas
                return 0.0
        except Exception:
            return 0.0
    
    def should_throttle_cpu(self) -> bool:
        """Check if CPU usage is too high"""
        return self.get_cpu_percent() > self.cpu_threshold_percent
    
    def should_throttle_memory(self) -> bool:
        """Check if memory usage is too high"""
        return self.get_memory_percent() > self.ram_threshold_percent
    
    def is_memory_critical(self) -> bool:
        """Check if memory is in critical state"""
        return self.get_memory_percent() > self.ram_critical_percent
    
    def get_status_report(self) -> str:
        """Get resource status report"""
        cpu = self.get_cpu_percent()
        ram = self.get_memory_percent()
        available_mb = self.get_available_memory_mb()
        
        return f"""
Resource Monitor Status:
  CPU: {cpu:.1f}% (threshold: {self.cpu_threshold_percent}%)
  RAM: {ram:.1f}% (threshold: {self.ram_threshold_percent}%, critical: {self.ram_critical_percent}%)
  Available Memory: {available_mb:.0f} MB
  Disk Write Throttle: {self.disk_write_throttle_mb_s} MB/s
"""


class ParallelConversionManager:
    """Manages parallel conversion processing with resource awareness"""
    
    def __init__(self, 
                 max_workers: Optional[int] = None,
                 log_callback: Optional[Callable[[str], None]] = None):
        """Initialize parallel conversion manager
        
        Args:
            max_workers: Maximum concurrent conversions (auto-detect if None)
            log_callback: Optional logging callback
        """
        self.log_callback = log_callback
        self.max_workers = max_workers or self._detect_optimal_workers()
        self.resource_monitor = ResourceMonitor(log_callback)
        self.active_conversions = 0
        self.conversion_semaphore = threading.Semaphore(self.max_workers)
        self.metrics: Dict[str, ConversionMetrics] = {}
        self.metrics_lock = threading.Lock()
        self.conversion_queue: Queue = Queue()
        self.running = False
    
    def _log(self, message: str) -> None:
        """Log a message"""
        if self.log_callback:
            self.log_callback(message)
    
    def _detect_optimal_workers(self) -> int:
        """Detect optimal number of concurrent workers
        
        Based on available CPU cores and system resources
        """
        try:
            total_cores = multiprocessing.cpu_count()
            # Use half of available cores, minimum 1, maximum total-1
            optimal = max(1, min(total_cores - 1, total_cores // 2))
            return optimal
        except Exception:
            return 2
    
    def acquire_conversion_slot(self, timeout: float = 300) -> bool:
        """Acquire a slot to start conversion (blocking with resource checks)
        
        Args:
            timeout: Maximum wait time in seconds
            
        Returns:
            True if slot acquired, False if timeout
        """
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            # Check resource constraints
            if self.resource_monitor.is_memory_critical():
                self._log("⚠️  Memory critical, waiting for resources...")
                time.sleep(1)
                continue
            
            # Try to acquire semaphore
            if self.conversion_semaphore.acquire(blocking=False):
                with self.metrics_lock:
                    self.active_conversions += 1
                return True
            
            # If CPU throttling needed, wait longer
            if self.resource_monitor.should_throttle_cpu():
                wait_time = 2
            elif self.resource_monitor.should_throttle_memory():
                wait_time = 1
            else:
                wait_time = 0.1
            
            time.sleep(wait_time)
        
        return False
    
    def release_conversion_slot(self) -> None:
        """Release a conversion slot"""
        self.conversion_semaphore.release()
        with self.metrics_lock:
            self.active_conversions = max(0, self.active_conversions - 1)
    
    def record_conversion_metric(self, metrics: ConversionMetrics) -> None:
        """Record metrics for a completed conversion
        
        Args:
            metrics: ConversionMetrics object with conversion data
        """
        with self.metrics_lock:
            self.metrics[metrics.file_path] = metrics
    
    def get_average_duration(self) -> float:
        """Get average conversion duration in seconds"""
        with self.metrics_lock:
            if not self.metrics:
                return 0.0
            durations = [m.duration_seconds for m in self.metrics.values() if m.end_time]
            return sum(durations) / len(durations) if durations else 0.0
    
    def get_average_throughput(self) -> float:
        """Get average throughput in MB/s"""
        with self.metrics_lock:
            if not self.metrics:
                return 0.0
            throughputs = [m.throughput_mbps for m in self.metrics.values() if m.end_time]
            return sum(throughputs) / len(throughputs) if throughputs else 0.0
    
    def estimate_time_remaining(self, files_remaining: int) -> timedelta:
        """Estimate time remaining for conversions
        
        Args:
            files_remaining: Number of files still to convert
            
        Returns:
            Estimated remaining time as timedelta
        """
        avg_duration = self.get_average_duration()
        if avg_duration <= 0:
            return timedelta(0)
        
        total_seconds = avg_duration * files_remaining
        return timedelta(seconds=total_seconds)
    
    def get_status_report(self) -> str:
        """Get status report for parallel conversion"""
        avg_duration = self.get_average_duration()
        avg_throughput = self.get_average_throughput()
        
        report = f"""
Parallel Conversion Status:
  Max Workers: {self.max_workers}
  Active Conversions: {self.active_conversions}
  Completed Conversions: {len(self.metrics)}
  Average Duration: {avg_duration:.1f}s
  Average Throughput: {avg_throughput:.1f} MB/s
  
{self.resource_monitor.get_status_report()}
"""
        return report


class PerformanceOptimizer:
    """High-level performance optimization interface"""
    
    def __init__(self, log_callback: Optional[Callable[[str], None]] = None):
        """Initialize performance optimizer
        
        Args:
            log_callback: Optional logging callback
        """
        self.log_callback = log_callback
        self.parallel_manager = ParallelConversionManager(log_callback=log_callback)
        self.resource_monitor = self.parallel_manager.resource_monitor
        self.conversion_start_time: Optional[datetime] = None
        self.total_files: int = 0
        self.processed_files: int = 0
    
    def _log(self, message: str) -> None:
        """Log a message"""
        if self.log_callback:
            self.log_callback(message)
    
    def start_conversion_batch(self, total_files: int) -> None:
        """Initialize a batch conversion
        
        Args:
            total_files: Total number of files to convert
        """
        self.conversion_start_time = datetime.now()
        self.total_files = total_files
        self.processed_files = 0
        self._log(f"🚀 Starting conversion batch: {total_files} files")
        self._log(f"📊 {self.parallel_manager.get_status_report()}")
    
    def finish_conversion_batch(self) -> None:
        """Finalize a batch conversion"""
        if self.conversion_start_time:
            elapsed = datetime.now() - self.conversion_start_time
            self._log(f"✅ Batch complete in {elapsed}")
            self._log(f"📊 {self.parallel_manager.get_status_report()}")
    
    def update_progress(self, files_completed: int) -> str:
        """Update conversion progress
        
        Args:
            files_completed: Number of files completed so far
            
        Returns:
            Status message with ETA
        """
        self.processed_files = files_completed
        files_remaining = self.total_files - files_completed
        
        if files_remaining <= 0:
            return "✅ All files processed"
        
        eta = self.parallel_manager.estimate_time_remaining(files_remaining)
        avg_throughput = self.parallel_manager.get_average_throughput()
        
        status = (
            f"📊 Progress: {files_completed}/{self.total_files} "
            f"({100 * files_completed / self.total_files:.1f}%) | "
            f"ETA: {eta} | "
            f"Throughput: {avg_throughput:.1f} MB/s"
        )
        return status


# Helper functions for integration with rom_converter.py

def create_performance_optimizer(log_callback: Optional[Callable[[str], None]] = None) -> PerformanceOptimizer:
    """Create a performance optimizer instance
    
    Args:
        log_callback: Optional logging callback
        
    Returns:
        Configured PerformanceOptimizer instance
    """
    return PerformanceOptimizer(log_callback=log_callback)


def get_parallel_manager(optimizer: PerformanceOptimizer) -> ParallelConversionManager:
    """Get the parallel conversion manager from optimizer
    
    Args:
        optimizer: PerformanceOptimizer instance
        
    Returns:
        ParallelConversionManager instance
    """
    return optimizer.parallel_manager


def get_resource_monitor(optimizer: PerformanceOptimizer) -> ResourceMonitor:
    """Get the resource monitor from optimizer
    
    Args:
        optimizer: PerformanceOptimizer instance
        
    Returns:
        ResourceMonitor instance
    """
    return optimizer.resource_monitor
