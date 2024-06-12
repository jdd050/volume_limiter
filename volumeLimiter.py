'''
volumeLimiter.py
Provides a temporary solution for people experiencing issues in windows where the volume suddenly increases to full randomly
'''
import tkinter as tk
import os
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

class main:
    def __init__(self):
        # initialize application
        self.root = tk.Tk()
        self.root.title("Volume Limiter V1.0")
        # get device, interface, and volume vars
        self.devices = AudioUtilities.GetSpeakers()
        self.interface = self.devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(self.interface, POINTER(IAudioEndpointVolume))
        # get percentage to dB conversions
        with open("decibels.txt", "r") as dbText:
            decibels = [float(num) for num in dbText.readlines()]
            pass
        with open("percentages.txt", "r") as pcText:
            percentages = [int(num) for num in pcText.readlines()]
        # create dictionary with conversions
        self.conversions = {}
        for db in decibels:
            for pc in percentages:
                self.conversions[pc] = db
                percentages.remove(pc)
                break

    def makeGUI(self):
        # make title
        titleFrame = tk.Frame(self.root)
        titleText = tk.Label(titleFrame, text="Volume Limiter for Windows V1.0", font=("Helvetica", 20)).pack(padx=20, pady=20)
        titleFrame.pack(side=tk.TOP)
        # desired volume
        volumeFrame = tk.Frame(self.root)
        volumeText = tk.Label(volumeFrame, text="Maximum volume (0%-100%):").pack()
        self.entry = tk.Entry(volumeFrame)
        self.entry.pack()
        volumeFrame.pack(side=tk.TOP, padx=5)
        return None
    
    def limitVolume(self):
        volumeNum = int(self.entry.get())
        self.root.destroy()
        print(f"Maintaining volume level {volumeNum}")
        while True:
            current = self.volume.GetMasterVolumeLevel()
            if current != volumeNum:
                self.volume.SetMasterVolumeLevel(self.conversions[volumeNum], None)
            else:
                continue
        return None

    pass

if __name__ == "__main__":
    volume = main()
    volume.makeGUI()
    # create button to start script
    btnFrame = tk.Frame(volume.root)
    startBtn = tk.Button(btnFrame, text="Start", command=volume.limitVolume).pack(padx=5, pady=5)
    btnFrame.pack(side=tk.TOP)
    volume.root.mainloop()