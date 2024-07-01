import subprocess

# 执行a.py文件
subprocess.run(['python', 'autoDownloadEdge.py'])

# 执行完成后再执行b.py文件
subprocess.run(['python', 'MergeData.py'])