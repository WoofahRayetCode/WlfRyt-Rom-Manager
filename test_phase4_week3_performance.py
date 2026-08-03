"""
Phase 4 Week 3 Performance Tests

Tests for performance optimization features:
- Parallel conversion processing
- Resource monitoring
- ETA calculation
- Metrics tracking
"""

import pytest
import time
import threading
from pathlib import Path
from unittest.mock import MagicMock, patch
from datetime import timedelta

from performance_optimizer import (
    ConversionMetrics,
    ResourceMonitor,
    ParallelConversionManager,
    PerformanceOptimizer,
    create_performance_optimizer,
)


class TestConversionMetrics:
    """Test ConversionMetrics data tracking"""
    
    def test_metrics_initialization(self):
        """Test creating conversion metrics"""
        metrics = ConversionMetrics(
            file_path="/tmp/test.iso",
            file_size_bytes=1024 * 1024,  # 1 MB
        )
        assert metrics.file_path == "/tmp/test.iso"
        assert metrics.file_size_bytes == 1024 * 1024
        assert metrics.success is False
    
    def test_metrics_duration_calculation(self):
        """Test duration calculation"""
        start_time = time.time()
        metrics = ConversionMetrics(
            file_path="/tmp/test.iso",
            start_time=start_time,
            end_time=start_time + 10.0,  # 10 seconds
        )
        assert metrics.duration_seconds == pytest.approx(10.0, abs=0.1)
    
    def test_metrics_throughput_calculation(self):
        """Test throughput calculation"""
        metrics = ConversionMetrics(
            file_path="/tmp/test.iso",
            file_size_bytes=1024 * 1024 * 100,  # 100 MB
            start_time=time.time() - 10.0,  # 10 seconds ago
            end_time=time.time(),
        )
        # Should be approximately 10 MB/s
        assert metrics.throughput_mbps >= 9.0
        assert metrics.throughput_mbps <= 11.0
    
    def test_metrics_compression_ratio(self):
        """Test compression ratio calculation"""
        metrics = ConversionMetrics(
            file_path="/tmp/test.iso",
            file_size_bytes=1024 * 1024 * 100,  # 100 MB input
            output_size_bytes=1024 * 1024 * 50,  # 50 MB output
        )
        assert metrics.compression_ratio == pytest.approx(0.5)


class TestResourceMonitor:
    """Test system resource monitoring"""
    
    def test_resource_monitor_initialization(self):
        """Test resource monitor creation"""
        monitor = ResourceMonitor()
        assert monitor is not None
        assert monitor.cpu_threshold_percent == 95
        assert monitor.ram_threshold_percent == 80
    
    def test_resource_monitor_with_callback(self):
        """Test resource monitor with logging callback"""
        messages = []
        def log_fn(msg):
            messages.append(msg)
        
        monitor = ResourceMonitor(log_callback=log_fn)
        monitor._log("Test message")
        assert "Test message" in messages
    
    def test_get_cpu_percent(self):
        """Test CPU percentage retrieval"""
        monitor = ResourceMonitor()
        cpu_percent = monitor.get_cpu_percent()
        assert isinstance(cpu_percent, float)
        assert 0 <= cpu_percent <= 100
    
    def test_get_memory_percent(self):
        """Test memory percentage retrieval"""
        monitor = ResourceMonitor()
        mem_percent = monitor.get_memory_percent()
        assert isinstance(mem_percent, float)
        assert 0 <= mem_percent <= 100
    
    def test_get_available_memory_mb(self):
        """Test available memory retrieval"""
        monitor = ResourceMonitor()
        available_mb = monitor.get_available_memory_mb()
        assert isinstance(available_mb, float)
        assert available_mb > 0
    
    def test_should_throttle_cpu(self):
        """Test CPU throttling check"""
        monitor = ResourceMonitor()
        # Set threshold very high so we don't throttle
        monitor.cpu_threshold_percent = 999
        assert monitor.should_throttle_cpu() is False
    
    def test_should_throttle_memory(self):
        """Test memory throttling check"""
        monitor = ResourceMonitor()
        # Set threshold very high so we don't throttle
        monitor.ram_threshold_percent = 999
        assert monitor.should_throttle_memory() is False
    
    def test_get_status_report(self):
        """Test status report generation"""
        monitor = ResourceMonitor()
        report = monitor.get_status_report()
        assert isinstance(report, str)
        assert "Resource Monitor Status" in report
        assert "CPU:" in report
        assert "RAM:" in report


class TestParallelConversionManager:
    """Test parallel conversion manager"""
    
    def test_manager_initialization(self):
        """Test parallel manager creation"""
        manager = ParallelConversionManager()
        assert manager is not None
        assert manager.max_workers >= 1
        assert manager.active_conversions == 0
    
    def test_manager_with_callback(self):
        """Test parallel manager with logging"""
        messages = []
        def log_fn(msg):
            messages.append(msg)
        
        manager = ParallelConversionManager(log_callback=log_fn)
        manager._log("Test message")
        assert "Test message" in messages
    
    def test_manager_custom_workers(self):
        """Test setting custom worker count"""
        manager = ParallelConversionManager(max_workers=4)
        assert manager.max_workers == 4
    
    def test_acquire_conversion_slot(self):
        """Test acquiring conversion slot"""
        manager = ParallelConversionManager(max_workers=2)
        
        # Should acquire first slot
        assert manager.acquire_conversion_slot(timeout=1) is True
        assert manager.active_conversions == 1
        
        # Should acquire second slot
        assert manager.acquire_conversion_slot(timeout=1) is True
        assert manager.active_conversions == 2
        
        # Third should timeout (max 2 workers)
        assert manager.acquire_conversion_slot(timeout=0.5) is False
    
    def test_release_conversion_slot(self):
        """Test releasing conversion slot"""
        manager = ParallelConversionManager(max_workers=2)
        
        # Acquire both slots
        manager.acquire_conversion_slot(timeout=1)
        manager.acquire_conversion_slot(timeout=1)
        assert manager.active_conversions == 2
        
        # Release one
        manager.release_conversion_slot()
        assert manager.active_conversions == 1
        
        # Should now be able to acquire another
        assert manager.acquire_conversion_slot(timeout=1) is True
        assert manager.active_conversions == 2
    
    def test_record_conversion_metric(self):
        """Test recording conversion metrics"""
        manager = ParallelConversionManager()
        
        metrics = ConversionMetrics(
            file_path="/tmp/test1.iso",
            file_size_bytes=1024 * 1024 * 100,
        )
        metrics.end_time = time.time()
        
        manager.record_conversion_metric(metrics)
        assert "/tmp/test1.iso" in manager.metrics
    
    def test_get_average_duration(self):
        """Test average duration calculation"""
        manager = ParallelConversionManager()
        
        # Record some metrics
        for i in range(3):
            metrics = ConversionMetrics(
                file_path=f"/tmp/test{i}.iso",
                start_time=time.time() - 10.0,
                end_time=time.time(),
            )
            manager.record_conversion_metric(metrics)
        
        avg_duration = manager.get_average_duration()
        assert avg_duration > 0
        assert 9.0 <= avg_duration <= 11.0
    
    def test_get_average_throughput(self):
        """Test average throughput calculation"""
        manager = ParallelConversionManager()
        
        # Record some metrics
        for i in range(3):
            metrics = ConversionMetrics(
                file_path=f"/tmp/test{i}.iso",
                file_size_bytes=1024 * 1024 * 100,  # 100 MB
                start_time=time.time() - 10.0,
                end_time=time.time(),
            )
            manager.record_conversion_metric(metrics)
        
        avg_throughput = manager.get_average_throughput()
        assert avg_throughput > 0
        # Should be approximately 10 MB/s
        assert 9.0 <= avg_throughput <= 11.0
    
    def test_estimate_time_remaining(self):
        """Test ETA estimation"""
        manager = ParallelConversionManager()
        
        # Record a metric to establish baseline
        metrics = ConversionMetrics(
            file_path="/tmp/test.iso",
            start_time=time.time() - 10.0,
            end_time=time.time(),
        )
        manager.record_conversion_metric(metrics)
        
        # Estimate time for 5 remaining files
        eta = manager.estimate_time_remaining(5)
        assert isinstance(eta, timedelta)
    
    def test_get_status_report(self):
        """Test status report generation"""
        manager = ParallelConversionManager()
        report = manager.get_status_report()
        assert isinstance(report, str)
        assert "Parallel Conversion Status" in report
        assert "Max Workers" in report


class TestPerformanceOptimizer:
    """Test high-level performance optimizer"""
    
    def test_optimizer_initialization(self):
        """Test performance optimizer creation"""
        optimizer = PerformanceOptimizer()
        assert optimizer is not None
        assert optimizer.parallel_manager is not None
        assert optimizer.resource_monitor is not None
    
    def test_optimizer_with_callback(self):
        """Test performance optimizer with logging"""
        messages = []
        def log_fn(msg):
            messages.append(msg)
        
        optimizer = PerformanceOptimizer(log_callback=log_fn)
        optimizer._log("Test message")
        assert "Test message" in messages
    
    def test_create_performance_optimizer(self):
        """Test factory function"""
        optimizer = create_performance_optimizer()
        assert optimizer is not None
        assert isinstance(optimizer, PerformanceOptimizer)
    
    def test_start_conversion_batch(self):
        """Test batch initialization"""
        optimizer = PerformanceOptimizer()
        optimizer.start_conversion_batch(10)
        
        assert optimizer.total_files == 10
        assert optimizer.processed_files == 0
        assert optimizer.conversion_start_time is not None
    
    def test_update_progress(self):
        """Test progress updates"""
        optimizer = PerformanceOptimizer()
        optimizer.start_conversion_batch(10)
        
        # Record a metric for ETA calculation
        metrics = ConversionMetrics(
            file_path="/tmp/test.iso",
            start_time=time.time() - 10.0,
            end_time=time.time(),
        )
        optimizer.parallel_manager.record_conversion_metric(metrics)
        
        # Update progress
        status = optimizer.update_progress(5)
        assert isinstance(status, str)
        assert "Progress" in status
        assert "50.0%" in status
    
    def test_finish_conversion_batch(self):
        """Test batch finalization"""
        optimizer = PerformanceOptimizer()
        optimizer.start_conversion_batch(10)
        time.sleep(0.1)  # Small delay
        optimizer.finish_conversion_batch()
        
        # Should complete without error
        assert optimizer is not None


class TestEndToEndPerformance:
    """End-to-end performance tests"""
    
    def test_full_conversion_workflow(self):
        """Test complete conversion workflow with metrics"""
        optimizer = create_performance_optimizer()
        
        # Start batch
        optimizer.start_conversion_batch(3)
        
        # Simulate 3 conversions
        for i in range(3):
            # Acquire slot
            assert optimizer.parallel_manager.acquire_conversion_slot(timeout=1)
            
            # Record metric
            metrics = ConversionMetrics(
                file_path=f"/tmp/test{i}.iso",
                file_size_bytes=1024 * 1024 * 50,
                start_time=time.time(),
            )
            time.sleep(0.1)  # Simulate work
            metrics.end_time = time.time()
            metrics.success = True
            
            optimizer.parallel_manager.record_conversion_metric(metrics)
            optimizer.parallel_manager.release_conversion_slot()
            
            # Update progress
            status = optimizer.update_progress(i + 1)
            assert "Progress" in status or "All files processed" in status
        
        # Finish batch
        optimizer.finish_conversion_batch()
        
        # Verify metrics
        assert len(optimizer.parallel_manager.metrics) == 3
        assert optimizer.parallel_manager.get_average_throughput() > 0
    
    def test_parallel_execution_safety(self):
        """Test thread-safe parallel execution"""
        optimizer = create_performance_optimizer()
        optimizer.start_conversion_batch(10)
        
        def simulate_conversion(file_id):
            # Acquire slot
            if optimizer.parallel_manager.acquire_conversion_slot(timeout=5):
                try:
                    metrics = ConversionMetrics(
                        file_path=f"/tmp/test{file_id}.iso",
                        file_size_bytes=1024 * 1024 * 50,
                    )
                    time.sleep(0.05)  # Simulate work
                    metrics.end_time = time.time()
                    optimizer.parallel_manager.record_conversion_metric(metrics)
                finally:
                    optimizer.parallel_manager.release_conversion_slot()
        
        # Spawn multiple threads
        threads = []
        for i in range(10):
            t = threading.Thread(target=simulate_conversion, args=(i,))
            threads.append(t)
            t.start()
        
        # Wait for all threads
        for t in threads:
            t.join(timeout=10)
        
        # Verify all conversions recorded
        assert len(optimizer.parallel_manager.metrics) == 10


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
