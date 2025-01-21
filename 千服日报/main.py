import subprocess
import sys


def run_script(script_name):
    try:
        result = subprocess.run(["python", script_name], check=True, capture_output=True, text=True)
        print(result.stdout)
        print(result.stderr)
    except subprocess.CalledProcessError as e:
        print(f"Error running {script_name}: {e}", file=sys.stderr)
        print(e.output, file=sys.stderr)


if __name__ == "__main__":
    run_script("autoDownloadEdge.py")
    run_script("MergeData.py")

    # 保持窗口打开
    input("Press Enter to exit...")
