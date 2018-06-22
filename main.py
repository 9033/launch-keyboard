import pywintypes
import win32com.directsound.directsound as ds
import win32event
import wave
import math
import struct
import sys

#p/f = t second
def aaaaa(p,f):
    if p/f < 0.01:
        return p/f*100 #fade in
    elif p/f > 0.04:
        return (-100)*(p/f-0.04)+1 #fade out
    else:
        return 1

def datawav(wavefname,freq):
    w=wave.open(wavefname,"wb")
    totaltime=0.05 #seconds
    channels=2
    samplewidth=2
    frate=44100 #hz
    datasize=int(totaltime*channels*samplewidth*frate)
    data=bytearray(datasize)

    #freq=500 #hz
    #time=1/frate
    afreq=2*math.pi*(freq/frate)
    amp=(1<<(samplewidth*8-1))-1 #signed
    num2byte = lambda fnum:(int(fnum)).to_bytes(2,byteorder='little',signed=True)
    x=0
    while x < datasize:
        #phase=(x/4) % frate  #(nchannels+samplewidth=4)
        phase=(x/4)  #(nchannels+samplewidth=4)
        #left
        num=(math.sin(afreq*phase)*aaaaa(phase,frate))*amp
        num=num2byte(num)
        data[x]=num[0]
        x+=1
        data[x]=num[1]
        x+=1
        #right
        num=(math.cos(afreq*phase)*aaaaa(phase,frate))*amp
        num=num2byte(num)
        data[x]=num[0]
        x+=1
        data[x]=num[1]
        x+=1
        
    w.setframerate(frate)
    w.setnchannels(channels)
    w.setsampwidth(samplewidth)  
    w.writeframes(data)
    w.close()

def chkwav(wavefname):    
    f=open(wavefname,"rb")
    WAV_HEADER_SIZE = struct.calcsize('<4sl4s4slhhllhh4sl')
    h=struct.unpack('<4sl4s4slhhllhh4sl', f.read(WAV_HEADER_SIZE))
    print(h)
    print('fmtsize', h[4])
    print('format', h[5])
    print('nchannels', h[6])
    print('samplespersecond', h[7])
    print('datarate', h[8])
    print('blockalign', h[9])
    print('bitspersample', h[10])
    f.close()
    
    print()
    
    w=wave.open(wavefname,"rb")
    #print('fmtsize', 4)
    #print('format', 5)
    print('nchannels', w.getnchannels())
    print('samplespersecond', w.getframerate())
    print('datarate', w.getnchannels()*w.getsampwidth()*w.getframerate())
    print('blockalign', w.getnchannels()*w.getsampwidth())
    print('bitspersample', w.getsampwidth()*8)
    print('getnframes', w.getnframes())
    print('wave bytes', w.getnframes()*w.getnchannels()*w.getsampwidth())

def loadwav(wavefname):    
    d = ds.DirectSoundCreate(None, None)
    d.SetCooperativeLevel(None, ds.DSSCL_PRIORITY)

    w=wave.open(wavefname,"rb")

    sdesc = ds.DSBUFFERDESC()
    sdesc.lpwfxFormat = pywintypes.WAVEFORMATEX()
    sdesc.lpwfxFormat.wFormatTag = pywintypes.WAVE_FORMAT_PCM
    sdesc.lpwfxFormat.nChannels = w.getnchannels()
    sdesc.lpwfxFormat.nSamplesPerSec = w.getframerate()
    sdesc.lpwfxFormat.nAvgBytesPerSec = w.getnchannels()*w.getsampwidth()*w.getframerate()
    sdesc.lpwfxFormat.nBlockAlign = w.getnchannels()*w.getsampwidth()
    sdesc.lpwfxFormat.wBitsPerSample = w.getsampwidth()*8
    sdesc.dwBufferBytes = w.getnframes()*w.getnchannels()*w.getsampwidth()
    sdesc.dwFlags = ds.DSBCAPS_STICKYFOCUS | ds.DSBCAPS_CTRLPOSITIONNOTIFY

    buffer = d.CreateSoundBuffer(sdesc, None)

    buffer.Update(0, w.readframes(w.getnframes()))
    w.close()
    return buffer

def playwav(buffer, flag=True):
    if flag:
        event = win32event.CreateEvent(None, 0, 0, None)
        notify = buffer.QueryInterface(ds.IID_IDirectSoundNotify)
        notify.SetNotificationPositions((ds.DSBPN_OFFSETSTOP, event))

    buffer.Play(0)
    
    if flag:
        win32event.WaitForSingleObject(event, -1)

        
        
def fileexist(filename):
    return [False,True][1]

def genwavfile():
    keycode=0
    while keycode <= 0xff:
        if not fileexist("data"+str(keycode)+".wav"):
            datawav("data"+str(keycode)+".wav",keycode*10)
        keycode+=1
    if not fileexist("data500.wav"):
        datawav("data500.wav",500)
    if not fileexist("data1000.wav"):
        datawav("data1000.wav",1000)
        
#main
sbuf = {}
def wavemain():
    #num2byte = lambda fnum:(int(fnum) & 0xffff).to_bytes(2,byteorder='little',signed=True)
    #print(num2byte(32766))
    genwavfile()
    #chkwav()
    
    #bind wave to buffer
    sbuf['500']=loadwav("data500.wav")
    sbuf['1000']=loadwav("data1000.wav")
    
    keycode=0
    while keycode <= 0xff:
        sbuf[keycode]=loadwav("data"+str(keycode)+".wav")
        keycode+=1

    rungui()

    #playwav(buf)
    #sys.exit()

from tkinter import *
from tkinter import ttk
def rungui():
    #from tkinter import messagebox
    root = Tk()

    var = StringVar()
    label = Message( root, textvariable=var, width=200, justify=RIGHT,relief=RAISED)
    var.set('keyboard')
    label.pack()


    var1 = StringVar()
    label = Message( root, textvariable=var1, width=100, justify=RIGHT,relief=RAISED)
    var1.set("mouse")
    label.pack()


    def b_callback():
        playwav(sbuf[ '1000' ],False)
        #print("click")

    b = Button(root, text="OK", command=b_callback)
    #b.bind('<Enter>', print("mouse inside"))
    #b.bind('<Leave>', print("mouse outside"))
    b.pack()

    def key(event):
        #print ("pressed", repr(event.char))
        #print ("pressed", event.keycode)
        var.set("keyboard"+" "+ str(event.keycode))
        playwav(sbuf[ event.keycode ],False)

    def callback(event):
        frame.focus_set()
        var1.set("clicked at"+" "+ str(event.x)+" "+ str(event.y))
        playwav(sbuf[ '500' ],False)
        #print ("clicked at", event.x, event.y)

    frame = Frame(root, width=100, height=100)
    frame.bind("<Key>", key)
    frame.bind("<Button-1>", callback)
    frame.pack()

    #w=Tk.Toplevel()
    #top.bind('<Enter>', print("mouse inside"))
    #top.bind('<Leave>', print("mouse inside"))
    #root.pack()
    root.mainloop()

if __name__=='__main__':
    wavemain()
