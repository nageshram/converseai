import pyaudio

p = pyaudio.PyAudio()
for i in range(p.get_device_count()):
    device_info = p.get_device_info_by_index(i)
    print(f"Index {i}: {device_info['name']}")
p.terminate()
