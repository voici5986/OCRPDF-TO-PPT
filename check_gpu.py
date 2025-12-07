# -*- coding: utf-8 -*-
"""
GPU Environment Check Script
"""
import sys
import subprocess

print("=" * 60)
print("GPU Environment Check")
print("=" * 60)

# 1. Check PaddlePaddle Version
print("\n[1] Checking PaddlePaddle Installation...")
try:
    import paddle
    print(f"[OK] PaddlePaddle Version: {paddle.__version__}")

    # Check if GPU version
    if 'post' in paddle.__version__ or 'gpu' in str(paddle.__version__):
        print("[OK] GPU version installed")
    else:
        print("[!!] CPU version installed!")
        print("     Need GPU version: pip install paddlepaddle-gpu")
except ImportError:
    print("[!!] PaddlePaddle not installed")
    sys.exit(1)

# 2. Check CUDA availability
print("\n[2] Checking CUDA...")
try:
    is_cuda = paddle.device.is_compiled_with_cuda()
    print(f"[OK] CUDA compiled: {is_cuda}")
    if is_cuda:
        cuda_ver = paddle.version.cuda()
        print(f"[OK] CUDA version: {cuda_ver}")
    else:
        print("[!!] PaddlePaddle not compiled with CUDA (CPU version)")
except Exception as e:
    print(f"[!!] Check failed: {e}")

# 3. Check GPU count
print("\n[3] Checking GPU devices...")
try:
    gpu_count = paddle.device.cuda.device_count()
    print(f"[OK] Found {gpu_count} GPU device(s)")

    if gpu_count == 0:
        print("[!!] No GPU detected")
        print("     Possible reasons:")
        print("     1. No NVIDIA GPU")
        print("     2. GPU driver not installed")
        print("     3. CPU version of PaddlePaddle")
    else:
        for i in range(gpu_count):
            props = paddle.device.cuda.get_device_properties(i)
            print(f"     GPU {i}: {props.name}")
except Exception as e:
    print(f"[!!] Check failed: {e}")

# 4. Test device setting
print("\n[4] Testing device setting...")
try:
    # Try to set GPU
    paddle.set_device('gpu:0')
    device = paddle.get_device()
    print(f"[OK] Current device: {device}")

    if 'gpu' in device.lower():
        print("[OK] GPU is available!")
    else:
        print("[!!] Switched to CPU")
except Exception as e:
    print(f"[!!] GPU not available: {e}")
    print("     Using CPU")

# 5. Check NVIDIA driver
print("\n[5] Checking NVIDIA driver...")
try:
    result = subprocess.run(['nvidia-smi'], capture_output=True, text=True, timeout=5, encoding='utf-8', errors='ignore')
    if result.returncode == 0:
        print("[OK] NVIDIA driver installed")
        # Extract GPU info
        lines = result.stdout.split('\n')
        for line in lines:
            if any(keyword in line for keyword in ['NVIDIA', 'GeForce', 'RTX', 'GTX', 'Tesla', 'Quadro']):
                print(f"     {line.strip()}")
    else:
        print("[!!] Cannot run nvidia-smi")
except FileNotFoundError:
    print("[!!] nvidia-smi not found")
    print("     Possible reasons:")
    print("     1. No NVIDIA GPU")
    print("     2. NVIDIA driver not installed")
except Exception as e:
    print(f"[!!] Check failed: {e}")

# Summary
print("\n" + "=" * 60)
print("Summary")
print("=" * 60)

try:
    if paddle.device.is_compiled_with_cuda() and paddle.device.cuda.device_count() > 0:
        print("[OK] GPU environment is ready!")
        print("\nRecommendation: Use GPU in PPT Editor settings")
    else:
        print("[!!] GPU not available")
        print("\nSolutions:")

        # Check if CPU version
        if not paddle.device.is_compiled_with_cuda():
            print("\n1. Current: CPU version of PaddlePaddle")
            print("   Solution: Reinstall GPU version")
            print("   >>> pip uninstall paddlepaddle")
            print("   >>> pip install paddlepaddle-gpu")

        # Check GPU count
        if paddle.device.cuda.device_count() == 0:
            print("\n2. No GPU device detected")
            print("   Please check:")
            print("   - Do you have NVIDIA GPU?")
            print("   - Is NVIDIA driver installed?")
            print("   - Try running: nvidia-smi")

        print("\nCurrent recommendation: Use CPU mode in PPT Editor")
except Exception as e:
    print(f"Error during diagnosis: {e}")

print("=" * 60)
